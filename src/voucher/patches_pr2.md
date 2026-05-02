# PR-2 应用 patches

## 范围

| 项 | 内容 | 改动位置 |
|---|---|---|
| **A2** | `getInfoByID("")` → 用 `VouSchemaCache.CloneXxxEmptyOrLoad`（懒加载克隆 schema），消除 30 万次 SELECT TOP 0（含 11 表 LEFT JOIN）| 工程 C |

## 跨工程

`VouSchemaCache.bas` **只加到工程 C（POPBus3GL2Service）**——`megetDocByID` 是 `t_FVou_M` 类的方法，属于工程 C，其他工程不需要此缓存。

懒加载：每次 `CloneMainEmptyOrLoad(objDS)` 内部首先 `EnsureLoaded(objDS)`，缓存命中即可 Clone；首次或失败时回退原 SELECT TOP 0 路径。

---

## P2.1：`t_FVou_M.cls` — `megetDocByID` cache fast path（**工程 C**）

### 修改前

```vb
Private Function megetDocByID(ByVal cid As String, ByVal objDS As HHDataService.sysDataService, _
                               Optional ByVal CIDIsEmptyRaiseErr As Boolean = False) As Boolean
On Error GoTo ErrHandler
    Dim SQL  As String

    If Trim(cid) = "" And CIDIsEmptyRaiseErr = True Then
        Call Err.Raise(ERROR_FORSYSTEM, , "无效的凭证ID")
    End If

    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " * FROM FVou_M with(nolock) WHERE BillID='" & cid & "'"
    Set mMainData = objDS.OpenRecordsetBySQL(SQL, False, True)

    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " dbo.FVou_I.*, FVou_I.ITMID as NewITMID, FVou_I.ITMID as oldITMID, ..."
    SQL = SQL & "FROM dbo.FVou_I with(nolock) LEFT OUTER JOIN ..."
    Set mItemsData = objDS.OpenRecordsetBySQL(SQL, False, True)

    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " * FROM FVou_II with(nolock) WHERE BillID='" & cid & "'"
    Set mIItemsData = objDS.OpenRecordsetBySQL(SQL, False, True)

    megetDocByID = True

ErrHandler:
    If Err.Number <> 0 Then Call Err.Raise(Err.Number)
End Function
```

### 修改后

```vb
Private Function megetDocByID(ByVal cid As String, ByVal objDS As HHDataService.sysDataService, _
                               Optional ByVal CIDIsEmptyRaiseErr As Boolean = False) As Boolean
On Error GoTo ErrHandler
    Dim SQL  As String

    If Trim(cid) = "" And CIDIsEmptyRaiseErr = True Then
        Call Err.Raise(ERROR_FORSYSTEM, , "无效的凭证ID")
    End If

    ' === PR-2 cache fast path（仅 cid="" 即创建新凭证场景）===
    If cid = "" Then
        Set mMainData = VouSchemaCache.CloneMainEmptyOrLoad(objDS)
        Set mItemsData = VouSchemaCache.CloneItemsEmptyOrLoad(objDS)
        Set mIItemsData = VouSchemaCache.CloneIItemsEmptyOrLoad(objDS)
        If Not mMainData Is Nothing And Not mItemsData Is Nothing And Not mIItemsData Is Nothing Then
            megetDocByID = True
            Exit Function
        End If
        ' 缓存克隆失败 → 回退原 SELECT TOP 0
        Set mMainData = Nothing: Set mItemsData = Nothing: Set mIItemsData = Nothing
    End If

    ' === 原路径（cid 非空，或 cache 加载失败）===
    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " * FROM FVou_M with(nolock) WHERE BillID='" & cid & "'"
    Set mMainData = objDS.OpenRecordsetBySQL(SQL, False, True)

    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " dbo.FVou_I.*, FVou_I.ITMID as NewITMID, FVou_I.ITMID as oldITMID, ..."
    Set mItemsData = objDS.OpenRecordsetBySQL(SQL, False, True)

    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " * FROM FVou_II with(nolock) WHERE BillID='" & cid & "'"
    Set mIItemsData = objDS.OpenRecordsetBySQL(SQL, False, True)

    megetDocByID = True

ErrHandler:
    If Err.Number <> 0 Then Call Err.Raise(Err.Number)
End Function
```

---

## 部署核对清单

- [ ] `VouSchemaCache.bas` 添加到工程 **C**（POPBus3GL2Service）
- [ ] 应用 P2.1 到 `t_FVou_M.megetDocByID`（工程 C）
- [ ] 编译工程 C 并部署
- [ ] 灰度测试

## 100% 等价性

- 缓存克隆失败时回退原 SELECT TOP 0 路径
- ADO `Fields.Append` 构造的 disconnected recordset 与 SELECT TOP 0 行为一致：
  - 字段名 / 类型 / DefinedSize / Attributes（仅保留 Append 接受的子集）/ Precision/NumericScale 完全复制
- 后续 `AddNew + Update + sysDS.SaveData` 行为与原 connected recordset 一致

## 性能收益

```
Baseline (after PR-1): 25.1 min
After PR-2:            20.1 min
增量节省:              5.0 min (20%)
```
