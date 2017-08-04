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
 
'����������ӿ�. ������е�����, �Դ�������¼�
Private WithEvents m_BillTransfer   As K3BillTransfer.Bill
Attribute m_BillTransfer.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillTransfer As Object)
 
    '�ӿ�ʵ��
    'ע��: �˷����������, �����޸�
    Set m_BillTransfer = oBillTransfer
 
End Sub

Private Sub Class_Terminate()
 
    '�ͷŽӿڶ���
    'ע��: �˷����������, �����޸�
    Set m_BillTransfer = Nothing

End Sub

Private Sub m_BillTransfer_BeforeSave(ByVal bNew As Boolean, ReturnCode As Long)


    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� BeforeSave
    With m_BillTransfer
    On Error GoTo er
        Set cn = New ADODB.Connection
        cn.CursorLocation = adUseClient

        cn.Open .Cnnstring
        Dim sql As String
        Set rs = New ADODB.Recordset

        '�����к�
        Dim cInterID As Long
        Dim cEnterID As Long
        Dim cStockID As Long
        Dim cItemID As Long
        Dim cBatchNo As Long
        Dim cQty As Long
        Dim cSType As Long
        '�������ݱ���
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
                    m = MsgBox("��" & i & "������û��ԭ�����Ž�����������⣬�Ƿ������", vbYesNo + vbExclamation)
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
                                MsgBox ("��" & i & "�����ϵĵ�ǰ���ο�治�㣬���ܱ��浥�ݣ�")
                                ReturnCode = -1
                                GoTo ExitSub
                                '��治��
                            ElseIf (rs("FH") = "Y" And rs("SK") = "Y") Then
                                '���Է���
                                GoTo NextFor
                            ElseIf (rs("FH") = "Y" And rs("SK") = "W") Then
                                'û������
                                r = MsgBox("��" & i & "������û�������������Ƿ������", vbYesNo)
                                If (r = 6) Then
                                    GoTo NextFor
                                Else
                                    ReturnCode = -1
                                    GoTo ExitSub
                                End If
                            ElseIf (rs("FH") = "Y" And rs("SK") = "N") Then
                                '���ⲻ��
                                MsgBox ("��" & i & "�����ϵĵ�ǰ������������" & rs("BDSK") & "С�ڷ������������ܱ��浥�ݣ�")
                                ReturnCode = -1
                                GoTo ExitSub
                            Else
                            
                            End If
                        ElseIf (rs.RecordCount = 0) Then
                            MsgBox ("��" & i & "�����Ͽ�治�㣬���з����� ")
                            ReturnCode = -1
                            GoTo ExitSub
                        Else
                            MsgBox ("���������⣬����ϵ����Ա�� �����У�" & i)
                            ReturnCode = -1
                            GoTo ExitSub
                        End If
'                        If (rs.RecordCount = 1 And rs("FH") = "N") Then
'                                MsgBox ("��" & i & "�����ϵĵ�ǰ���ο��ÿ�治�㣬���ܱ��浥�ݣ�")
'                                ReturnCode = -1
'                                GoTo ExitSub
'
'                        ElseIf rs.RecordCount > 1 Then
'                            MsgBox ("���������⣬����ϵ����Ա�� �����У�" & i)
'                            ReturnCode = -1
'                            GoTo ExitSub
'                        End If
                        rs.Close
                    Else
                        MsgBox ("��" & i & "������δѡ��ֿ⣬���ܱ��浥�ݣ�")
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
    'a = MsgBox("δ֪��������ϵ����Ա��", vbOKOnly + vbCritical, "ϵͳ����")
    MsgBox "δ֪��������ϵ����Ա��" & Chr(13) & "����ţ�" & err.Number & Chr(13) & err.Description, vbOKOnly + vbCritical '��ȡ����źʹ���˵����
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
"      ,(CASE WHEN (ISNULL(INV.FQty,0)-ISNULL(QTSK.FQty,0)>=@Qty) THEN 'Y' ELSE 'N' END) AS FH--��������Ƿ��㹻����" & Chr(13) & _
"      ,(CASE WHEN ISNULL(BDLS.FQty,0)>=@Qty THEN 'Y' WHEN BDLS.FQty IS NULL THEN 'W' ELSE 'N' END) AS SK--���������Ƿ��㹻����,�������ΪNULL˵��δ����,��ʾȷ��" & Chr(13) & _
"  FROM t_ICItem AS ICI" & Chr(13) & _
"  LEFT JOIN (--��ʱ���" & Chr(13) & Chr(10)
sql = sql + "" & _
"     SELECT INV.FStockID,INV.FItemID,FBatchNo,SUM(FQty) AS FQty" & Chr(13) & _
"     FROM ICInventory AS INV" & Chr(13) & _
"     GROUP BY INV.FStockID,INV.FItemID,FBatchNo" & _
"  ) AS INV ON ICI.FItemID=INV.FItemID AND INV.FBatchNo=@BatchNo" & Chr(13) & _
"  LEFT JOIN (--��������" & Chr(13) & _
"      SELECT LS.FStockID,LS.FItemID,SUM(FQty) AS FQty" & Chr(13) & _
"      From t_LockStock AS LS" & Chr(13) & _
"      GROUP BY LS.FStockID,LS.FItemID,LS.FBatchNo" & Chr(13) & _
"  ) AS BDLS ON INV.FItemID=BDLS.FItemID" & Chr(13) & _
"  LEFT JOIN (--������������" & Chr(13) & _
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
'     ,(CASE WHEN (ISNULL(INV.FQty,0)-ISNULL(QTSK.FQty,0)>=@Qty) THEN 'Y' ELSE 'N' END) AS FH--��������Ƿ��㹻����
'     ,(CASE WHEN ISNULL(BDLS.FQty,0)>=@Qty THEN 'Y' WHEN BDLS.FQty IS NULL THEN 'W' ELSE 'N' END) AS SK--���������Ƿ��㹻����,�������ΪNULL˵��δ����,��ʾȷ��
' FROM t_ICItem AS ICI
' LEFT JOIN (--��ʱ���
'    SELECT INV.FStockID,INV.FItemID,FBatchNo,SUM(FQty) AS FQty
'    FROM ICInventory AS INV
'    GROUP BY INV.FStockID,INV.FItemID,FBatchNo
' ) AS INV ON ICI.FItemID=INV.FItemID AND INV.FBatchNo=@BatchNo
' LEFT JOIN (--��������
'     SELECT LS.FStockID,LS.FItemID,SUM(FQty) AS FQty
'     From t_LockStock AS LS
'     GROUP BY LS.FStockID,LS.FItemID,LS.FBatchNo
' ) AS BDLS ON INV.FItemID=BDLS.FItemID
' LEFT JOIN (--������������
'     SELECT LS.FStockID,LS.FItemID,FBatchNo,SUM(FQty) AS FQty
'     From t_LockStock LS
'     WHERE  FBatchNo=@BatchNo AND FQty>0 AND NOT(FInterID=@InterID AND FEntryID=@EnteyID)
'     GROUP BY LS.FStockID,LS.FItemID,FBatchNo
' ) AS QTSK ON INV.FItemID=QTSK.FItemID
' WHERE ICI.FItemID=@ItemID AND INV.FStockID=@StockID

End Function

