VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_FHTZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is Industry Bill Interface Class, made by K3BOSPLUGINSWIZAED"
 
'定义插件对象接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillTransfer   As K3BillTransfer.Bill
Attribute m_BillTransfer.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillTransfer As Object)
 
    '接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = oBillTransfer
 
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = Nothing

End Sub

Private Sub m_BillTransfer_BeforeSave(ByVal bNew As Boolean, ReturnCode As Long)


    'TODO: 请在此处添加代码响应事件 BeforeSave
    With m_BillTransfer
    On Error GoTo er
        Set cn = New ADODB.Connection
        cn.CursorLocation = adUseClient

        cn.Open .Cnnstring
        Dim sql As String
        Set rs = New ADODB.Recordset

        '定义列号
        Dim cInterID As Long
        Dim cEnterID As Long
        Dim cStockID As Long
        Dim cItemID As Long
        Dim cBatchNo As Long
        Dim cQty As Long
        Dim cSType As Long
        '定义内容变量
        Dim InterID As Long
        Dim EntryID As Long
        Dim StockID As Long
        Dim ItemID As Long
        Dim BatchNoID As String
        Dim Qty As Double
        Dim msg As String
        msg = ""

        For i = LBound(.m_EntryCtl) To UBound(.m_EntryCtl)
            If (.EntryCtl(i).FieldName = "FStockID") Then cStockID = i
            If (.EntryCtl(i).FieldName = "FItemID") Then cItemID = i
            If (.EntryCtl(i).FieldName = "FQty") Then cQty = i
            If (.EntryCtl(i).FieldName = "FBatchNo") Then cFBatchNo = i
            If (.EntryCtl(i).FieldName = "FSourceInterId") Then cInterID = i
            If (.EntryCtl(i).FieldName = "FSourceEntryID") Then cEnterID = i
            If (.EntryCtl(i).FieldName = "FSourceTranType") Then cSType = i
        Next

        For i = LBound(.m_EntryCtl(cItemID).DInterID) To UBound(.m_EntryCtl(cItemID).DInterID)
            If (.GetGridText(i, cItemID) <> "") Then
                AA = .GetGridText(i, cInterID)
                If (.GetGridText(i, cInterID) = "" Or (.GetGridText(i, cEnterID) = "")) Then
                    m = MsgBox("第" & i & "行物料没有原单单号将会跳过库存检测，是否继续？", vbYesNo + vbExclamation)
                    If m = 7 Then
                        ReturnCode = -1
                        GoTo ExitSub
                    Else
                        GoTo NextFor
                    End If
                Else
                    If CInt(.m_EntryCtl(cStockID).DInterID(i)) > 0 Then
                        InterID = .GetGridText(i, cInterID)
                        EntryID = .GetGridText(i, cEnterID)
                        StockID = .m_EntryCtl(cStockID).DInterID(i)
                        ItemID = .m_EntryCtl(cItemID).DInterID(i)
                        BatchNo = .GetGridText(i, cFBatchNo)
                        Qty = .GetGridText(i, cQty)
        
                        sql = makeSQL(InterID, EntryID, StockID, ItemID, BatchNo, Qty)
                        rs.Open sql, cn
                        
                        If (rs.RecordCount = 1) Then
                            If (rs("FH") = "N") Then
                                MsgBox ("第" & i & "行物料的当前批次库存不足，不能保存单据！")
                                ReturnCode = -1
                                GoTo ExitSub
                                '库存不足
                            ElseIf (rs("FH") = "Y" And rs("SK") = "Y") Then
                                '可以发货
                                GoTo NextFor
                            ElseIf (rs("FH") = "Y" And rs("SK") = "W") Then
                                '没有锁库
                                r = MsgBox("第" & i & "行物料没有锁库数量，是否继续！", vbYesNo)
                                If (r = 6) Then
                                    GoTo NextFor
                                Else
                                    ReturnCode = -1
                                    GoTo ExitSub
                                End If
                            ElseIf (rs("FH") = "Y" And rs("SK") = "N") Then
                                '锁库不足
                                MsgBox ("第" & i & "行物料的当前批次锁库数量" & rs("BDSK") & "小于发货数量，不能保存单据！")
                                ReturnCode = -1
                                GoTo ExitSub
                            Else
                            
                            End If
                        ElseIf (rs.RecordCount = 0) Then
                            MsgBox ("第" & i & "行物料库存不足，进行发货！ ")
                            ReturnCode = -1
                            GoTo ExitSub
                        Else
                            MsgBox ("语句出现问题，请联系管理员！ 错误行：" & i)
                            ReturnCode = -1
                            GoTo ExitSub
                        End If
'                        If (rs.RecordCount = 1 And rs("FH") = "N") Then
'                                MsgBox ("第" & i & "行物料的当前批次可用库存不足，不能保存单据！")
'                                ReturnCode = -1
'                                GoTo ExitSub
'
'                        ElseIf rs.RecordCount > 1 Then
'                            MsgBox ("语句出现问题，请联系管理员！ 错误行：" & i)
'                            ReturnCode = -1
'                            GoTo ExitSub
'                        End If
                        rs.Close
                    Else
                        MsgBox ("第" & i & "行物料未选择仓库，不能保存单据！")
                        ReturnCode = -1
                        GoTo ExitSub
                    End If
                End If
            Else
                Exit For
            End If
NextFor:
            rs.Close
        Next
        ReturnCode = 0
GoTo ExitSub
er:
    ReturnCode = -1
    'a = MsgBox("未知错误，请联系管理员！", vbOKOnly + vbCritical, "系统错误")
    MsgBox "未知错误，请联系管理员！" & Chr(13) & "错误号：" & err.Number & Chr(13) & err.Description, vbOKOnly + vbCritical '获取错误号和错误说明。
ExitSub:
    cn.Close
    End With
End Sub
Public Function makeSQL(ByVal InterID As Long, ByVal EntryID As Long, ByVal StockID As Long, ByVal ItemID As Long, ByVal BatchNo As String, ByVal Qty As Double) As String
    Dim sql As String
sql = "" & _
"   DECLARE @InterID INT,@EnteyID INT,@StockID INT,@ItemID INT,@BatchNo NVARCHAR(200),@Qty DECIMAL(18,6)" & Chr(13) & _
"   SELECT" & Chr(13) & _
"      @InterID=" & InterID & Chr(13) & _
"      ,@EnteyID=" & EntryID & Chr(13) & _
"      ,@StockID=" & StockID & Chr(13) & _
"      ,@ItemID=" & ItemID & Chr(13) & _
"      ,@BatchNo='" & BatchNo & "'" & Chr(13) & _
"      ,@Qty=" & Qty & Chr(13) & _
"  SELECT" & _
"       ICI.FItemID" & Chr(13) & _
"      ,INV.FBatchNo" & Chr(13) & _
"      ,INV.FStockID" & Chr(13) & _
"      ,INV.FQty AS INVQty" & Chr(13) & _
"      ,ISNULL(BDLS.FQty,0) AS BDSK" & Chr(13) & _
"      ,ISNULL(QTSK.FQty,0) AS QTSK" & Chr(13) & _
"      ,(CASE WHEN (ISNULL(INV.FQty,0)-ISNULL(QTSK.FQty,0)>=@Qty) THEN 'Y' ELSE 'N' END) AS FH--库存数量是否足够发货" & Chr(13) & _
"      ,(CASE WHEN ISNULL(BDLS.FQty,0)>=@Qty THEN 'Y' WHEN BDLS.FQty IS NULL THEN 'W' ELSE 'N' END) AS SK--锁库数量是否足够发货,如果锁库为NULL说明未锁库,提示确定" & Chr(13) & _
"  FROM t_ICItem AS ICI" & Chr(13) & _
"  LEFT JOIN (--即时库存" & Chr(13) & Chr(10)
sql = sql + "" & _
"     SELECT INV.FStockID,INV.FItemID,FBatchNo,SUM(FQty) AS FQty" & Chr(13) & _
"     FROM ICInventory AS INV" & Chr(13) & _
"     GROUP BY INV.FStockID,INV.FItemID,FBatchNo" & _
"  ) AS INV ON ICI.FItemID=INV.FItemID AND INV.FBatchNo=@BatchNo" & Chr(13) & _
"  LEFT JOIN (--锁库数量" & Chr(13) & _
"      SELECT LS.FStockID,LS.FItemID,SUM(FQty) AS FQty" & Chr(13) & _
"      From t_LockStock AS LS" & Chr(13) & _
"      GROUP BY LS.FStockID,LS.FItemID,LS.FBatchNo" & Chr(13) & _
"  ) AS BDLS ON INV.FItemID=BDLS.FItemID" & Chr(13) & _
"  LEFT JOIN (--其他锁库数量" & Chr(13) & _
"      SELECT LS.FStockID,LS.FItemID,FBatchNo,SUM(FQty) AS FQty" & Chr(13) & _
"      From t_LockStock LS" & Chr(13) & _
"      WHERE  FBatchNo=@BatchNo AND FQty>0 AND NOT(FInterID=@InterID AND FEntryID=@EnteyID)" & Chr(13) & _
"      GROUP BY LS.FStockID,LS.FItemID,FBatchNo" & Chr(13) & _
"  ) AS QTSK ON INV.FItemID=QTSK.FItemID" & Chr(13) & _
"  WHERE ICI.FItemID=@ItemID AND INV.FStockID=@StockID"

    makeSQL = sql
    
    
'  DECLARE @InterID INT,@EnteyID INT,@StockID INT,@ItemID INT,@BatchNo NVARCHAR(200),@Qty DECIMAL(18,6)
'  SELECT
'     @InterID=8520
'     ,@EnteyID=1
'     ,@StockID=22109
'     ,@ItemID=1392
'     ,@BatchNo=''
'     ,@Qty=4.13
' SELECT
'      ICI.FItemID
'     ,INV.FBatchNo
'     ,INV.FStockID
'     ,INV.FQty AS INVQty
'     ,ISNULL(BDLS.FQty,0) AS BDSK
'     ,ISNULL(QTSK.FQty,0) AS QTSK
'     ,(CASE WHEN (ISNULL(INV.FQty,0)-ISNULL(QTSK.FQty,0)>=@Qty) THEN 'Y' ELSE 'N' END) AS FH--库存数量是否足够发货
'     ,(CASE WHEN ISNULL(BDLS.FQty,0)>=@Qty THEN 'Y' WHEN BDLS.FQty IS NULL THEN 'W' ELSE 'N' END) AS SK--锁库数量是否足够发货,如果锁库为NULL说明未锁库,提示确定
' FROM t_ICItem AS ICI
' LEFT JOIN (--即时库存
'    SELECT INV.FStockID,INV.FItemID,FBatchNo,SUM(FQty) AS FQty
'    FROM ICInventory AS INV
'    GROUP BY INV.FStockID,INV.FItemID,FBatchNo
' ) AS INV ON ICI.FItemID=INV.FItemID AND INV.FBatchNo=@BatchNo
' LEFT JOIN (--锁库数量
'     SELECT LS.FStockID,LS.FItemID,SUM(FQty) AS FQty
'     From t_LockStock AS LS
'     GROUP BY LS.FStockID,LS.FItemID,LS.FBatchNo
' ) AS BDLS ON INV.FItemID=BDLS.FItemID
' LEFT JOIN (--其他锁库数量
'     SELECT LS.FStockID,LS.FItemID,FBatchNo,SUM(FQty) AS FQty
'     From t_LockStock LS
'     WHERE  FBatchNo=@BatchNo AND FQty>0 AND NOT(FInterID=@InterID AND FEntryID=@EnteyID)
'     GROUP BY LS.FStockID,LS.FItemID,FBatchNo
' ) AS QTSK ON INV.FItemID=QTSK.FItemID
' WHERE ICI.FItemID=@ItemID AND INV.FStockID=@StockID

End Function


