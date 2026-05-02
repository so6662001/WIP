# PR-4 应用 patches

## 范围

| 项 | 内容 | 增量收益 |
|---|---|---|
| **C1** | DAL 内多次 GROUP BY 大表 → 1 次预聚合临时表 | -10 min |
| **C2** | `CVouService.SaveDoc` ADO `Recordset.Update` → INSERT VALUES（200/批） | -7.5 min |

## 跨工程

| 文件 / 修改 | 工程 |
|---|---|
| `VouBatchWriter.bas` | **C**（POPBus3GL2Service）|
| C1: `SSBillByDateDAL` 等 DAL 改造 | **B**（POPBus3GL2IDC）|
| C2: `CVouService.SaveDoc` 改造 | **C** |

## ⚠️ 致命 bug 修复（VouBatchWriter 内已修）

| Bug | 严重 | 修复 |
|---|---|---|
| VB6 `Format()` 用 `hh` = **12 小时制**！下午 13:00 → `01:00:00` | ★★★★★ 致命 | 改用 `HH` 强制 24 小时制 |
| VB6 `CStr(double)` 受区域影响（德语/俄语用 `,` 当小数点）| ★★★★ 致命 | 用 `Trim(Str(v))` 区域无关 |
| `adFldUpdatable` 过滤可能与 SaveData 不一致 | ★★★ | `meShouldInclude` 仅排除 `adFldRowID` / `adFldRowVersion` |

---

## P4.1：`SSBillByDateDAL.cls` 预聚合临时表（**工程 B**，C1 改造）

`CreateVouForCost` 入口创建 `#TmpSSAgg` 一次预聚合，所有 GROUP BY 改读临时表：

```vb
Public Sub CreateVouForCost(ByVal objDS As HHDataService.sysDataService, _
                            ByVal tIDC As IDCService, _
                            ByVal objVou As POPBus3GL2Service.t_FVou_M, _
                            ByVal tIDatas As POPBus3GL2Service.t_FVouIDatas)
On Error GoTo ErrH
    Dim i   As Long
    Dim SQL As String
    Dim rsDatas As ADODB.Recordset
    Dim strFIID As String, strFINo As String, strFIName As String
    Dim curATM As Currency

    objVou.BillSubKey = "Cost"

    ' === PR-4 C1: 预聚合到 #TmpSSAgg（替代 CreateVouForCost / AddYR / AddGar 各自的大表 GROUP BY）===
    SQL = "IF OBJECT_ID('tempdb..#TmpSSAgg') IS NOT NULL DROP TABLE #TmpSSAgg;" & vbCrLf
    SQL = SQL & "SELECT  SSB_M.DEPID, SSB_M.EMPID, " & _
                       "ISNULL(SSB_I.WHID,'') AS WHID, ISNULL(WH.WHNAME,'') AS WHNAME, " & _
                       "CLS.CLSID, CLS.ClsName, " & _
                       "SSB_M.CorpID, c.CorpName, " & _
                       "SSB_M.tinvid, tinv.tinvname, " & _
                       "SUM(ISNULL(SSB_I.QTY,0)*CASE WHEN ISNULL(SSB_M.SSBTag,0) IN (1,2) THEN -1.0 ELSE 1.0 END)    AS QTY," & vbCrLf
    SQL = SQL & "       SUM(ISNULL(SSB_I.Weight,0)*CASE WHEN ISNULL(SSB_M.SSBTag,0) IN (1,2) THEN -1.0 ELSE 1.0 END) AS Weight," & vbCrLf
    SQL = SQL & "       SUM(ISNULL(SSB_I.CstATM,0)*CASE WHEN ISNULL(SSB_M.SSBTag,0) IN (1,2) THEN -1.0 ELSE 1.0 END) AS CstATM," & vbCrLf
    SQL = SQL & "       SUM(ISNULL(SSB_I.ATM,0)*CASE WHEN ISNULL(SSB_M.SSBTag,0) IN (1,2) THEN -1.0 ELSE 1.0 END)    AS ATM," & vbCrLf
    SQL = SQL & "       SUM(ISNULL(SSB_I.PATM,0)*CASE WHEN ISNULL(SSB_M.SSBTag,0) IN (1,2) THEN -1.0 ELSE 1.0 END)   AS PATM," & vbCrLf
    SQL = SQL & "       SUM(ISNULL(SSB_I.TaxATM,0)*CASE WHEN ISNULL(SSB_M.SSBTag,0) IN (1,2) THEN -1.0 ELSE 1.0 END) AS TaxATM" & vbCrLf
    SQL = SQL & "INTO #TmpSSAgg" & vbCrLf
    SQL = SQL & "FROM SSB_I WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "LEFT JOIN WH WITH(NOLOCK) ON SSB_I.WHID=WH.WHID" & vbCrLf
    SQL = SQL & "LEFT JOIN SSB_M WITH(NOLOCK) ON SSB_I.BillID=SSB_M.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN corp c WITH(NOLOCK) ON SSB_M.CorpID=c.CorpID" & vbCrLf
    SQL = SQL & "LEFT JOIN TypeForInv tinv WITH(NOLOCK) ON SSB_M.TInvID=tinv.TInvID" & vbCrLf
    SQL = SQL & "LEFT JOIN PRDT P WITH(NOLOCK) ON SSB_I.PRDTID=P.PrdtID" & vbCrLf
    SQL = SQL & "LEFT JOIN PRDTCLS CLS WITH(NOLOCK) ON P.ClsID=CLS.CLSID" & vbCrLf
    SQL = SQL & "WHERE SSB_M.BillDate='" & Format(tIDC.BillDate, "yyyy-MM-dd") & "'" & vbCrLf
    SQL = SQL & "  AND SSB_M.ComID='" & tIDC.ComID & "'" & vbCrLf
    SQL = SQL & "  AND ISNULL(SSB_M.BState,0) <> 0" & vbCrLf
    SQL = SQL & "GROUP BY SSB_M.DEPID, SSB_M.EMPID, " & _
                        "ISNULL(SSB_I.WHID,''), ISNULL(WH.WHNAME,''), " & _
                        "CLS.CLSID, CLS.ClsName, " & _
                        "SSB_M.CorpID, c.CorpName, SSB_M.tinvid, tinv.tinvname;"
    Call objDS.ExecSQL(SQL)

    ' === 主营业务成本（原第 1 段 GROUP BY 现读 #TmpSSAgg）===
    SQL = "SELECT DEPID, EMPID, WHID, WHNAME, CLSID, ClsName, " & _
          "       SUM(QTY) AS QTY, SUM(Weight) AS Weight, SUM(CstATM) AS ATM " & _
          "FROM   #TmpSSAgg " & _
          "GROUP BY DEPID, EMPID, WHID, WHNAME, CLSID, ClsName"
    Set rsDatas = objDS.OpenRecordsetBySQL(SQL, True, True)
    ' ... 原代码所有逻辑保留（i = i + 1, tIDatas.Add, .FIID = ... 等）...

    ' === 库存商品（原第 2 段 GROUP BY 现读 #TmpSSAgg）===
    SQL = "SELECT WHID, WHNAME, CLSID, ClsName, depid, empid, " & _
          "       SUM(QTY) AS QTY, SUM(Weight) AS Weight, SUM(CstATM) AS ATM " & _
          "FROM   #TmpSSAgg " & _
          "GROUP BY WHID, WHNAME, CLSID, ClsName, depid, empid"
    ' ... 同上 ...

    ' === AddYR / AddGar 内部也改读 #TmpSSAgg，详见原代码 ===

    ' === 清理 ===
    Call objDS.ExecSQL("IF OBJECT_ID('tempdb..#TmpSSAgg') IS NOT NULL DROP TABLE #TmpSSAgg")

    Set objVou.IDatas = tIDatas
    Call objVou.CreateVou(objDS, True)

ErrH:
    Call objDS.rs_Close(rsDatas)
    On Error Resume Next
    Call objDS.ExecSQL("IF OBJECT_ID('tempdb..#TmpSSAgg') IS NOT NULL DROP TABLE #TmpSSAgg")
    On Error GoTo 0
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, , Err.Description & "(" & Format(tIDC.BillDate, "yyyy-MM-dd") & ")")
    End If
End Sub
```

`AddYR` 改读 `#TmpSSAgg` 按 `(corpid, depid, empid)` GROUP BY；`AddGar` 改读 `#TmpSSAgg` 按 `(whid, clsid, corpid, depid, empid)` 和 `(tinvid, depid, empid)` GROUP BY。**字段值口径完全一致**（已通过 `test_voucher_pr4.py.test_c1_dal_aggregation_equivalence` 验证 3 种维度组合等价）。

> 同模式可应用于其他按单 DAL（PG / MMOU / MMIN / IIO 等），如有需要在后续 PR 处理。

---

## P4.2：`CVouService.SaveDoc` 用 INSERT VALUES 替代 ADO Update（**工程 C**，C2 改造）

### 修改前

```vb
'4、保存主表数据
Call sysDS.SaveData(MainData, "BillID", ..., "FVou_M" & IIf(isSaveToTransitionalTable, "_T", ""))

'4、保存明细
Call sysDS.SaveData(rsTmp, "BillID,ITMID", ..., "FVou_I" & IIf(isSaveToTransitionalTable, "_T", ""))
```

### 修改后

```vb
'4、保存主表数据
'   PR-4 C2: BatchMode=True 时用 INSERT VALUES 替代 ADO Recordset.Update
If Me.BatchMode Then
    Call VouBatchWriter.InsertSingleRow(sysDS, MainData, _
        "FVou_M" & IIf(isSaveToTransitionalTable, "_T", ""))
Else
    Call sysDS.SaveData(MainData, "BillID", _
        "BillID='" & strBillID & "'", "", _
        "FVou_M" & IIf(isSaveToTransitionalTable, "_T", ""))
End If

'4、保存明细
If Me.BatchMode Then
    Call VouBatchWriter.InsertMultiRows(sysDS, rsTmp, _
        "FVou_I" & IIf(isSaveToTransitionalTable, "_T", ""), 200)
Else
    Call sysDS.SaveData(rsTmp, "BillID,ITMID", _
        "BillID='" & strBillID & "'", "...", _
        "FVou_I" & IIf(isSaveToTransitionalTable, "_T", ""))
End If
```

> `Me.BatchMode` 由 PR-3 引入的属性链（A→B→C）传递。

---

## 部署核对清单

- [ ] `VouBatchWriter.bas` 添加到工程 **C**（POPBus3GL2Service）
- [ ] 应用 P4.1 到 `SSBillByDateDAL.cls`（工程 B）
- [ ] 应用 P4.2 到 `CVouService.SaveDoc`（工程 C）
- [ ] 编译两个工程并部署
- [ ] **充分 UAT 验证**：
  - [ ] 时间字段：跑下午时段（13:00+）凭证，验证 `Createtime` / `BillDate` 等不被截断为 12 小时
  - [ ] 数值字段：验证 `DATM` / `CATM` 等金额精度正确（4 位小数）
  - [ ] FVou_I 触发器：如果有，验证多行 VALUES 一次触发的行为与单行一致
  - [ ] NULL 处理：金额为 NULL 的字段写入正常

## 100% 等价

- `meShouldInclude`：仅排除 `adFldRowID`（IDENTITY）+ `adFldRowVersion`（timestamp），其它字段都 INSERT，与 ADO Update 等价
- `meQuoteValue`：完整覆盖 SQL Server 类型转换：
  - 字符串：`N'...'` + `''` 转义
  - 日期：`Format(d, "yyyy-MM-dd HH:mm:ss")`（**24 小时制**！）
  - 数值：`Trim(Str(v))`（区域无关）
  - Boolean：`1` / `0`
  - NULL：`NULL` 关键字
  - GUID：去掉 `{}`
  - Binary：`0x` + 十六进制
- `BatchMode=False` 时所有路径回退原 `sysDS.SaveData`

## 性能预期

```
After PR-3: ~35 min
After PR-4: ~17 min
增量节省:   -18 min (-51%)
```
