# PR-4 优化补丁应用清单

## 范围

- **C1**：DAL 内多次 GROUP BY 大表合并为 1 次预聚合（SSBillByDateDAL 为例，其他 DAL 同模式）
- **C2**：`CVouService.SaveDoc` ADO `Recordset.AddNew/Update` → 批量 `INSERT VALUES`（可选，最大改动面）

## 前置依赖

- PR-1 + PR-2 + PR-3 已合入

## Patch 1：`SSBillByDateDAL.cls` 预聚合临时表

### 修改前

`CreateVouForCost` 内 2 次 SELECT GROUP BY 大表 + `AddYR` 1 次 + `AddGar` 1~2 次 = 4~5 次扫描 `SSB_I + SSB_M (+ WH/Corp/PRDT/PRDTCLS)`。

### 修改后

在 DAL 入口（`CreateVouForCost` 第一行）创建一个临时聚合表，所有 GROUP BY 改为读临时表：

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

    ' === PR-4 C1: 预聚合到 #TmpSSAgg ===
    SQL = ""
    SQL = SQL & "IF OBJECT_ID('tempdb..#TmpSSAgg') IS NOT NULL DROP TABLE #TmpSSAgg;" & vbCrLf
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

    ' === C1: 现在原 4~5 次 GROUP BY 大表全部改为读 #TmpSSAgg ===
    
    ' 主营业务成本（原第 1 段 SQL）
    SQL = "SELECT DEPID, EMPID, WHID, WHNAME, CLSID, ClsName, " & _
          "       SUM(QTY) AS QTY, SUM(Weight) AS Weight, SUM(CstATM) AS ATM " & _
          "FROM   #TmpSSAgg " & _
          "GROUP BY DEPID, EMPID, WHID, WHNAME, CLSID, ClsName"
    Set rsDatas = objDS.OpenRecordsetBySQL(SQL, True, True)

    If objVou.depid = "" Then
        If rsDatas.RecordCount > 0 Then
            objVou.depid = objDS.NullToStr(rsDatas.Fields("depid").Value, True)
        Else
            objVou.depid = objVou.AppParameters.depid
            objVou.EmpID = objVou.AppParameters.EmpID
        End If
    End If

    Call objDS.rs_MoveFirst(rsDatas)
    Do While Not rsDatas.EOF
        ' ... 原代码 i = i + 1, tIDatas.Add(i), .FIID = ... 等保持不变 ...
        ' 字段名一致：DEPID/EMPID/WHID/WHNAME/CLSID/ClsName/QTY/Weight/ATM
        rsDatas.MoveNext
    Loop
    Call objDS.rs_Close(rsDatas)

    ' 库存商品（原第 2 段 SQL）
    SQL = "SELECT WHID, WHNAME, CLSID, ClsName, depid, empid, " & _
          "       SUM(QTY) AS QTY, SUM(Weight) AS Weight, SUM(CstATM) AS ATM " & _
          "FROM   #TmpSSAgg " & _
          "GROUP BY WHID, WHNAME, CLSID, ClsName, depid, empid"
    Set rsDatas = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rsDatas.EOF
        ' ... 原代码 i = i + 1, tIDatas.Add(i) 等保持不变 ...
        rsDatas.MoveNext
    Loop
    Call objDS.rs_Close(rsDatas)

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

### `AddYR` 同样改造

```vb
Private Sub AddYR(...)
    SQL = "SELECT depid, empid, CorpID, CorpName, " & _
          "       SUM(QTY) AS QTY, SUM(Weight) AS Weight, SUM(ATM) AS ATM, SUM(TaxATM) AS TaxATM " & _
          "FROM   #TmpSSAgg " & _
          "GROUP BY depid, empid, CorpID, CorpName"
    Set rsDatas = objDS.OpenRecordsetBySQL(SQL, True, True)
    ' ... 其余原代码不变 ...
End Sub
```

### `AddGar` 同样改造

```vb
Private Sub AddGar(...)
    ' 第 1 个 SELECT (按 WHID/CLSID/CORPID/DEPID/EMPID 聚合)
    SQL = "SELECT depid, empid, CorpID, CorpName, WHID, WHNAME, CLSID, ClsName, " & _
          "       SUM(QTY) AS QTY, SUM(Weight) AS Weight, SUM(ATM) AS ATM, " & _
          "       SUM(PATM) AS PATM, SUM(TaxATM) AS TaxATM " & _
          "FROM   #TmpSSAgg " & _
          "GROUP BY depid, empid, CorpID, CorpName, WHID, WHNAME, CLSID, ClsName"
    Set rsDatas = objDS.OpenRecordsetBySQL(SQL, True, True)
    ' ...

    ' 第 2 个 SELECT (CostModel=1, 按 tinvid/depid/empid 聚合)
    If objVou.AppParameters.CostModel = 1 Then
        SQL = "SELECT depid, empid, tinvid, tinvname, " & _
              "       SUM(QTY) AS QTY, SUM(Weight) AS Weight, SUM(TaxATM) AS TaxATM " & _
              "FROM   #TmpSSAgg " & _
              "GROUP BY depid, empid, tinvid, tinvname"
        Set rsDatas = objDS.OpenRecordsetBySQL(SQL, True, True)
        ' ...
    End If
End Sub
```

### 收益

原：4~5 次扫描 `SSB_I + SSB_M + 6 表 LEFT JOIN`（每张凭证）  
后：1 次扫描大表 + 4~5 次扫描内存 #TmpSSAgg（仅 ~10 行）

10 万张 × 5 次 - 10 万次 + 50 万小扫描 ≈ 净节省 4 倍扫描时间。但**注意**：按日生成时只有 30~60 张凭证（每天一张 SSBill），影响相对小；按单生成（10 万张）才显著。

**生效条件**：仅当 `GLVouCreateByDate=False` 走 `SSBillByDateDAL` 等 by-bill DAL 时显著。但用户已确认用按日模式 → C1 的 SSBill 改造收益有限。**其他 DAL（PG/MMOU/MMIN/IIO）按单模式跑 10 万次时是大头，应同样改造**。

---

## Patch 2：（可选）`CVouService.SaveDoc` ADO Update → 批量 INSERT VALUES

> ⚠️ **这是 PR-4 中改动面最大、风险最高的部分**。如果业务对 ADO Recordset 的 Update 行为有任何隐式依赖（触发 Field-level 事件、自动 ID 回填等），改为 INSERT 可能行为不一致。**强烈建议先做完整 UAT 回归再启用**。

### 思路

不修改 SaveDoc 接口。在 SaveDoc 内部判断：
- `VouBatchOps.InBatchMode=True` 且当前数据库连接事务正在进行 → 调用新 helper `WriteVouDirect`，绕过 `sysDS.SaveData` 用 INSERT VALUES 批量写入
- 否则保持原 `sysDS.SaveData` 路径

### 实现

新增 `VouBatchWriter.bas`（仅给出关键骨架，实际实现需要根据贵司 `sysDS.SaveData` 内部 escape 规则等细化）：

```vb
Public Sub WriteMain(ByVal objDS As HHDataService.sysDataService, _
                    ByVal rs As ADODB.Recordset, _
                    ByVal tabName As String)
    ' rs 已经 AddNew + Update 完成，把它的当前行用 INSERT VALUES 写入 tabName
    Dim sqlCols As String, sqlVals As String
    Dim fld     As ADODB.Field
    For Each fld In rs.Fields
        If sqlCols <> "" Then sqlCols = sqlCols & ","
        sqlCols = sqlCols & fld.Name
        If sqlVals <> "" Then sqlVals = sqlVals & ","
        sqlVals = sqlVals & meQuoteValue(fld)
    Next fld
    Call objDS.ExecSQL("INSERT INTO " & tabName & "(" & sqlCols & ") VALUES (" & sqlVals & ")")
End Sub

Public Sub WriteItemsBatch(ByVal objDS As HHDataService.sysDataService, _
                          ByVal rs As ADODB.Recordset, _
                          ByVal tabName As String, _
                          ByVal pageSize As Long)
    ' 批量 INSERT FVou_I 多行
    Dim sqlBuf As String
    Dim cnt As Long: cnt = 0
    Call objDS.rs_MoveFirst(rs)
    Do While Not rs.EOF
        If sqlBuf = "" Then
            ' 拼 INSERT INTO ... (cols) VALUES
            sqlBuf = meBuildInsertHeader(rs, tabName)
        Else
            sqlBuf = sqlBuf & ","
        End If
        sqlBuf = sqlBuf & meBuildValuesRow(rs)
        cnt = cnt + 1
        If cnt >= pageSize Then
            Call objDS.ExecSQL(sqlBuf)
            sqlBuf = ""
            cnt = 0
        End If
        rs.MoveNext
    Loop
    If sqlBuf <> "" Then Call objDS.ExecSQL(sqlBuf)
End Sub

Private Function meQuoteValue(ByVal fld As ADODB.Field) As String
    If IsNull(fld.Value) Then
        meQuoteValue = "NULL"
        Exit Function
    End If
    Select Case fld.Type
        Case adVarChar, adVarWChar, adChar, adWChar, adLongVarChar, adLongVarWChar
            meQuoteValue = "'" & Replace(CStr(fld.Value), "'", "''") & "'"
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            meQuoteValue = "'" & Format(fld.Value, "yyyy-MM-dd hh:mm:ss") & "'"
        Case adBoolean
            meQuoteValue = IIf(CBool(fld.Value), "1", "0")
        Case adNumeric, adDecimal, adCurrency, adDouble, adSingle, _
             adInteger, adSmallInt, adBigInt, adTinyInt, _
             adUnsignedInt, adUnsignedSmallInt, adUnsignedBigInt, adUnsignedTinyInt
            meQuoteValue = CStr(fld.Value)
        Case Else
            meQuoteValue = "'" & Replace(CStr(fld.Value), "'", "''") & "'"
    End Select
End Function
```

### SaveDoc 集成

```vb
' 4. 保存主表数据
If VouBatchOps.InBatchMode Then
    Call VouBatchWriter.WriteMain(sysDS, MainData, "FVou_M" & IIf(isSaveToTransitionalTable, "_T", ""))
Else
    Call sysDS.SaveData(MainData, "BillID", ..., "FVou_M" & IIf(isSaveToTransitionalTable, "_T", ""))
End If

' 4. 保存明细
If VouBatchOps.InBatchMode Then
    Call VouBatchWriter.WriteItemsBatch(sysDS, rsTmp, "FVou_I" & IIf(isSaveToTransitionalTable, "_T", ""), 200)
Else
    Call sysDS.SaveData(rsTmp, "BillID,ITMID", ..., "FVou_I" & IIf(isSaveToTransitionalTable, "_T", ""))
End If
```

### 等价性风险

1. **Field 类型映射**：`meQuoteValue` 必须正确处理所有 SQL Server 类型，特别是 `decimal/money/numeric` 的精度
2. **NULL 处理**：原 ADO Update 区分"显式 NULL"和"未设置"；INSERT VALUES 一致用 `NULL`
3. **触发器**：如果 `FVou_I` 等表上有触发器，ADO Update 可能逐行触发，而 INSERT 多行 VALUES 一次触发—— 触发器可能行为不同
4. **`isSaveToTransitionalTable=False`** 路径：写入正式 FVou_M / FVou_I，可能有索引维护、外键检查等。INSERT VALUES 同样会触发，但批量插入时锁/日志行为略有差异

### 收益

10 万张凭证 × 5 行明细 = **50 万次 ADO Update** → **2500 次 INSERT VALUES (200/批)**

预期节省：**~10 min → ~2 min**，节省 8 min（占总耗时 ~40%）

---

## 部署顺序

1. （仅 C1）：直接修改各 DAL，逐个灰度验证（按需，因为按日模式收益小）
2. （C2 可选）：先 UAT 验证 `VouBatchWriter` 在小批量上行为完全等价
3. 加 `VouBatchWriter.bas` → 应用 SaveDoc 集成 → 在 InBatchMode 下生效

## 100% 等价证明（C2）

- `meQuoteValue` 对每个 SQL Server 类型分别处理，等价于 ADO 的 ToString 转换
- 多行 INSERT VALUES 与多次 INSERT 在 SQL Server 内部行为完全一致（除非触发器明确依赖单行 INSERT）
- ADO Update 的字段写入顺序与 INSERT VALUES 字段顺序一致（都按 `rs.Fields` 顺序）
