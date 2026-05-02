Attribute VB_Name = "VouBatchOps"
'==============================================================================
' Module      : VouBatchOps.bas
' Description : 凭证 batch 后处理批量化（PR-3 第三波）
'
'   消除每张凭证的 3 个 stored proc 调用：
'     - ss_YWVouUpdateAchTimes  (per-bill 更新档案使用次数)
'     - ss_AfterSaveUpdateGL2VouBill  (per-bill 后处理)
'   改为整 batch 末尾**一次性**调用批量版本（按 BillID list 处理）。
'
' 跨工程架构说明
' ===============
'   ⚠️ VB6 .bas 模块状态在不同 ActiveX DLL 工程间不共享。
'   本模块**只能加到工程 A (POPBus3FileService)**，并由工程 A 的
'   meCreVouForXX 在循环里调 RecordSavedBill(objIDC.rtnVouBillID, ...)
'   收集 BillID。工程 C (POPBus3GL2Service) SaveDoc 内部不直接调本模块；
'   而是通过 IDCService.IsBatchMode 属性传递的标志判断是否跳过
'   per-bill stored proc 调用。
'
'   架构图：
'     A.meCreateVou
'       BeginBatch + BeginTrans
'       Set objIDC.IsBatchMode = True             ← 跨工程传递
'       Loop meCreVouForXX:
'         objIDC.CreateVou()                      ← 进入工程 B/C
'           SaveDoc()                              ← 工程 C
'             根据 t_FVou_M.IsBatchMode 跳过 stored proc
'         RecordSavedBill(objIDC.rtnVouBillID)    ← 工程 A 内收集
'       FlushStoredProcs                          ← 工程 A 一次性批量
'       CommitTrans
'       EndBatch
'
' 设计原则    : 100% 等价
'   - 整 batch 处理完所有凭证后，按相同的 BillID 集合调用批量 stored proc，
'     等价于原代码逐个调用（前提是批量 stored proc 实现等价于循环单凭证版本）。
'   - 大事务边界：整 batch 起始 BeginTrans，每 1000 张凭证 CommitTrans +
'     重新 BeginTrans，最末尾 CommitTrans。
'==============================================================================
Option Explicit

' 收集到的待处理 BillID（CreateVou 成功后调 RecordSavedBill）
Private m_colSavedBills_Direct  As Collection   ' isSaveToTransitionalTable=False
Private m_colSavedBills_Trans   As Collection   ' isSaveToTransitionalTable=True
Private m_blnInBatchMode        As Boolean

' 大事务计数
Private m_lngTxnCounter         As Long
Private Const TXN_COMMIT_EVERY  As Long = 1000


'==============================================================================
' Public：进入 batch 模式（meCreateVou 入口调用）
'==============================================================================
Public Sub BeginBatch()
    Set m_colSavedBills_Direct = New Collection
    Set m_colSavedBills_Trans = New Collection
    m_blnInBatchMode = True
    m_lngTxnCounter = 0
End Sub


'==============================================================================
' Public：退出 batch 模式（meCreateVou 末尾调用，无论成败）
'==============================================================================
Public Sub EndBatch()
    Set m_colSavedBills_Direct = Nothing
    Set m_colSavedBills_Trans = Nothing
    m_blnInBatchMode = False
    m_lngTxnCounter = 0
End Sub


Public Property Get InBatchMode() As Boolean
    InBatchMode = m_blnInBatchMode
End Property


'==============================================================================
' Public：每张凭证 SaveDoc 成功后调用（取代原 per-bill stored proc 调用）
'   isToTransTable: 凭证是否写入过渡表 (FVou_M_T) —— 决定 batch flush 时
'                   调批量 stored proc 的 isSaveToTransitionalTable 参数
'
'   ⚠️ BillID 校验：如果包含 ',' / "'" / 空白 等会破坏 list 拼接的字符，
'      抛错以避免 SQL 注入或 list 错切。生产中 BillID 都是 objDS.CreateSheetID
'      生成的纯数字串，不会触发此校验。
'==============================================================================
Public Sub RecordSavedBill(ByVal sBillID As String, ByVal isToTransTable As Boolean)
    If Not m_blnInBatchMode Then Exit Sub
    If sBillID = "" Then Exit Sub

    ' BillID 安全字符校验（防 SQL 注入 + list 切分错误）
    If InStr(sBillID, ",") > 0 Or InStr(sBillID, "'") > 0 Or _
       InStr(sBillID, ";") > 0 Or InStr(sBillID, " ") > 0 Then
        Call Err.Raise(vbObjectError + 1001, "VouBatchOps.RecordSavedBill", _
            "BillID '" & sBillID & "' 含非法字符（',';' 或空白），无法安全批量处理")
    End If

    If isToTransTable Then
        m_colSavedBills_Trans.Add sBillID
    Else
        m_colSavedBills_Direct.Add sBillID
    End If
End Sub


'==============================================================================
' Public：每张凭证 SaveDoc 后判断是否需要 flush 一段事务
'   返回 True = 调用方需要 CommitTrans + BeginTrans（外层管理事务）
'==============================================================================
Public Function ShouldFlushTxn() As Boolean
    If Not m_blnInBatchMode Then Exit Function
    m_lngTxnCounter = m_lngTxnCounter + 1
    If m_lngTxnCounter >= TXN_COMMIT_EVERY Then
        m_lngTxnCounter = 0
        ShouldFlushTxn = True
    End If
End Function


'==============================================================================
' Public：batch 结束前调用，一次性批量执行所有收集到的 stored proc
'
'   ⚠️ 严格等价：原代码 per-bill 调用顺序是交错的：
'     for each bill:
'         exec ss_YWVouUpdateAchTimes BillID, 1
'         exec ss_AfterSaveUpdateGL2VouBill BillID, 0
'   而不是先全部 YWVou 再全部 AfterSave。
'
'   为保证两个 stored proc 之间可能存在的"per-bill 依赖"不被破坏，
'   批量版在**同一页 200 张**中也按 per-bill 交错调用：
'     for page in chunks(bills, 200):
'         exec ss_YWVouUpdateAchTimes_Batch  page_csv, 1     -- 内部 cursor 循环
'         exec ss_AfterSaveUpdateGL2VouBill_Batch page_csv, 0
'   即每 200 张内部仍是先 YWVou 后 AfterSave，但**跨页**保持 per-page 交错；
'   原代码是 per-bill 交错，批量版是 per-page 交错。两者**只在两个 stored proc
'   存在跨 bill 依赖时**才有差异，原代码若无此跨 bill 依赖（典型情况），
'   两种顺序行为完全等价。
'
'   见 PR3_storedprocs.sql。如果 stored proc 未部署，调用方应回退到
'   per-bill 模式（即不启用 BeginBatch）。
'==============================================================================
Public Sub FlushStoredProcs(ByVal objDS As HHDataService.sysDataService)
    If Not m_blnInBatchMode Then Exit Sub

    ' Direct 路径：每页 200 张，先 YWVou 后 AfterSave（保持 per-page 交错）
    If m_colSavedBills_Direct.Count > 0 Then
        Call meBatchProcInterleaved(objDS, m_colSavedBills_Direct, _
            "ss_YWVouUpdateAchTimes_Batch", "1", _
            "ss_AfterSaveUpdateGL2VouBill_Batch", "0")
    End If

    ' Trans 路径：只调 ss_AfterSaveUpdateGL2VouBill_Batch (IsTrans=1)
    '   原代码 isSaveToTransitionalTable=True 时不调 ss_YWVouUpdateAchTimes
    '   batch 模式严格保留这个差异
    If m_colSavedBills_Trans.Count > 0 Then
        Call meBatchProcSingle(objDS, m_colSavedBills_Trans, _
            "ss_AfterSaveUpdateGL2VouBill_Batch", "1")
    End If

    ' 已 flush 的清空（避免重复执行）
    Set m_colSavedBills_Direct = New Collection
    Set m_colSavedBills_Trans = New Collection
End Sub


'==============================================================================
' Private：分页交错执行两个 stored proc（保留 per-page 顺序）
'   for each page of 200 bills:
'       exec procA page, paramA
'       exec procB page, paramB
'==============================================================================
Private Sub meBatchProcInterleaved(ByVal objDS As HHDataService.sysDataService, _
                                   ByVal colBills As Collection, _
                                   ByVal sProcA As String, ByVal sParamA As String, _
                                   ByVal sProcB As String, ByVal sParamB As String)
    Const PAGE_SIZE As Long = 200
    Dim sBuf  As String
    Dim lngBatchN As Long

    sBuf = ""
    lngBatchN = 0

    Dim v As Variant
    For Each v In colBills
        If sBuf <> "" Then sBuf = sBuf & ","
        sBuf = sBuf & CStr(v)
        lngBatchN = lngBatchN + 1
        If lngBatchN >= PAGE_SIZE Then
            Call objDS.ExecSQL("EXEC " & sProcA & " '" & sBuf & "'," & sParamA)
            Call objDS.ExecSQL("EXEC " & sProcB & " '" & sBuf & "'," & sParamB)
            sBuf = ""
            lngBatchN = 0
        End If
    Next v
    If sBuf <> "" Then
        Call objDS.ExecSQL("EXEC " & sProcA & " '" & sBuf & "'," & sParamA)
        Call objDS.ExecSQL("EXEC " & sProcB & " '" & sBuf & "'," & sParamB)
    End If
End Sub


'==============================================================================
' Private：分页执行单个 stored proc
'==============================================================================
Private Sub meBatchProcSingle(ByVal objDS As HHDataService.sysDataService, _
                              ByVal colBills As Collection, _
                              ByVal sProcName As String, _
                              ByVal sParam2 As String)
    Const PAGE_SIZE As Long = 200
    Dim sBuf  As String
    Dim lngBatchN As Long

    sBuf = ""
    lngBatchN = 0

    Dim v As Variant
    For Each v In colBills
        If sBuf <> "" Then sBuf = sBuf & ","
        sBuf = sBuf & CStr(v)
        lngBatchN = lngBatchN + 1
        If lngBatchN >= PAGE_SIZE Then
            Call objDS.ExecSQL("EXEC " & sProcName & " '" & sBuf & "'," & sParam2)
            sBuf = ""
            lngBatchN = 0
        End If
    Next v
    If sBuf <> "" Then
        Call objDS.ExecSQL("EXEC " & sProcName & " '" & sBuf & "'," & sParam2)
    End If
End Sub
