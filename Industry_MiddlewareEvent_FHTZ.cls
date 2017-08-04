VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_MiddlewareEvent_FHTZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is IndustryMiddleWareEvents Interface Class, made by K3BOSPLUGINSWIZAED"
 
Public Function HookInvoke(ByVal EventID As Long, dctParams As KFO.Dictionary) As Long
 
    '定义插件对象接口. 必须具有的声明, 以此来获得事件
    '注意: 此方法必须存在, 请勿修改
    Select Case EventID
        Case 200001 '审核前事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'InterID                | 单据的内码
            'TranType               | 单据的事务类型
            'OperateDelete          | 默认值为0
            'ROB                    | 是否是红单
            'OperateCode            | 操作码:   1------审核
            '                       |           2------反审核
            '                       |           4------保存
            '                       |           8------删除
            '                       |           16-----反作废
            '                       |           32-----作废
            '--------------------------------------------------------
            
            '请在这里添加处理 '审核前' 代码
            
            If (dctParams("OperateCode") = 1 And dctParams("TranType") = 83) Then
                Set cn = New ADODB.Connection
                cn.CursorLocation = adUseClient
                connStr = Mid(dctParams("sDsn"), InStr(dctParams("sDsn"), "{") + 1, InStr(dctParams("sDsn"), "}") - (InStr(dctParams("sDsn"), "{") + 1))
                cn.Open connStr
                Dim sql As String
                
                sql = "SELECT FSourceInterId,FSourceEntryID,FItemID,FStockID,FBatchNo,FQty FROM SEOutStockEntry WHERE FInterID=" & CStr(dctParams("InterID"))
                Set rs = New ADODB.Recordset
                Set rs2 = New ADODB.Recordset
                rs.Open sql, cn
                
                If (Not rs.EOF) Then
                    For i = 1 To rs.RecordCount
                        On Error Resume Next
                        If ((rs("FSourceInterId") <> 0 And rs("FSourceEntryID") <> 0)) Then
                            sql = makeSQL(rs("FSourceInterId"), rs("FSourceEntryID"), rs("FItemID"), rs("FStockID"), rs("FBatchNo"), rs("FQty"))
                            rs2.Open sql, cn
                            If (rs2("Sate") = "N") Then
                                MsgBox ("第" & i & "行指定批号库存数量少于出库数量，不能出库！")
                                HookInvoke = 0
                                GoTo endFor
                            End If
                        Else
                                MsgBox ("第" & i & "行物料没有原单关联，不能审核！")
                                HookInvoke = 0
                                GoTo endFor
                        End If
                        rs.MoveNext
                    Next
                End If
                HookInvoke = 1
endFor:
 rs.Close
 rs = Nothing
 rs2.Close
 rs2 = Nothing
            Else
                HookInvoke = 1
            End If


 
    Case Else
    End Select

End Function
Public Function makeSQL(ByVal InterID As Long, ByVal EntryID As Long, ByVal ItemID As Long, ByVal StockID As Long, ByVal BatchNo As String, ByVal Qty As Double) As String
    Dim sql As String
    sql = "" & _
    " DECLARE @InterID INT,@EnteyID INT,@ItemID INT,@StockID INT,@BatchNo NVARCHAR(200),@Qty DECIMAL(18,6)" & _
    " SELECT  @InterID=" & InterID & _
    "         ,@EnteyID=" & EntryID & _
    "         ,@StockID=" & StockID & _
    "         ,@ItemID=" & ItemID & _
    "         ,@BatchNo='" & BatchNo & "'" & _
    "         ,@Qty=" & Qty & _
    " SELECT" & _
    "     INV.FStockID AS StockID" & _
    "     ,INV.FItemID AS ItemID" & _
    "     ,INV.FQty AS INVQty" & _
    "     ,ISNULL(BDSK.FQty,0) AS BDSK" & _
    "     ,ISNULL(QTSK.FQty,0) AS QTSK" & _
    "     ,(CASE WHEN INV.FQty-ISNULL(QTSK.FQty,0)>0 THEN 'Y' ELSE 'N' END) AS Sate" & _
    " FROM ICInventory AS INV" & _
    " LEFT JOIN(" & _
    "     SELECT FStockID,FItemID,FBatchNo,SUM(FQty) AS FQty" & _
    "     FROM t_LockStock" & _
    "     WHERE FInterID=@InterID AND FEntryID=@EnteyID"
    
    sql = sql & "" & _
    "     GROUP BY FStockID,FItemID,FBatchNo" & _
    " ) AS BDSK ON BDSK.FStockID=INV.FStockID AND BDSK.FItemID=INV.FItemID AND BDSK.FBatchNo=INV.FBatchNo" & _
    " LEFT JOIN (" & _
    "     SELECT FStockID,FItemID, FBatchNo,SUM(FQty) AS FQty" & _
    "     FROM t_LockStock" & _
    "     WHERE FInterID<>@InterID AND FEntryID<>@EnteyID" & _
    "     GROUP BY FStockID,FItemID,FBatchNo" & _
    " ) AS QTSK ON QTSK.FStockID=INV.FSecQty AND QTSK.FItemID=INV.FItemID AND BDSK.FBatchNo=INV.FBatchNo" & _
    " WHERE INV.FItemID=@ItemID AND INV.FStockID=@StockID AND INV.FBatchNo=@BatchNo"
    makeSQL = sql

End Function


