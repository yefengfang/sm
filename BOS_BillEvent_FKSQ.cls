VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_FKSQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is BillEvent Interface Class, made by K3BOSPLUGINSWIZAED"
 
'���� BillEvent �ӿ�. ������е�����, �Դ�������¼�
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillInterface As Object)
 
    'BillEvent �ӿ�ʵ��
    'ע��: �˷����������, �����޸�
    Set m_BillInterface = oBillInterface
 
End Sub

Private Sub Class_Terminate()
 
    '�ͷŽӿڶ���
    'ע��: �˷����������, �����޸�
    Set m_BillInterface = Nothing

End Sub

Private Sub m_BillInterface_Change(ByVal dct As KFO.IDictionary, ByVal dctFld As KFO.IDictionary, ByVal Col As Long, ByVal Row As Long, Cancel As Boolean)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼�
'--SELECT
'--  FItemID,
'--  ISNULL(F_105,'') KHM,
'--  ISNULL(FBank,'') KHYH,
'--  ISNULL(FAccount,'') YHZH,
'--  ISNULL(FTaxNum,'') SWDJH,
'--  ISNULL(F_102,'') KHM_S,
'--  ISNULL(F_103,'') KHYH_S,
'--  ISNULL(F_104,'')YHZH_S
'--FROM t_Supplier
'--WHERE FItemID = 441
    
    On Error GoTo ExitFun
    With m_BillInterface
        If (dct("FCaption") = "������Ŀ") Then
            Dim CMD As String
            CMD = "" & _
"SELECT                          " & _
"    FItemID,                    " & _
"    ISNULL(F_105,'') KHM,       " & _
"    ISNULL(FBank,'') KHYH,      " & _
"    ISNULL(FAccount,'') YHZH,   " & _
"    ISNULL(FTaxNum,'') SWDJH,   " & _
"    ISNULL(F_102,'') KHM_S,     " & _
"    ISNULL(F_103,'') KHYH_S,    " & _
"    ISNULL(F_104,'') YHZH_S     " & _
"FROM t_Supplier                 " & _
"WHERE FItemID ="


            Set rs = .K3Lib.GetData(CMD + dctFld("FFLD"))
            If (rs.RecordCount > 0) Then
                a = .SetFieldValue("FText5", rs.Fields("KHM"))
                a = .SetFieldValue("FText2", rs.Fields("KHM_S"))
                a = .SetFieldValue("FText3", rs.Fields("KHYH_S"))
                a = .SetFieldValue("FText4", rs.Fields("YHZH_S"))
            Else
                MsgBox ("δ�ҵ���Ӧ������Ŀ��")
            End If

        End If
    End With
GoTo ExitSub
Catch:
    'HandleError Err
ExitSub:
    Set obj = Nothing
    Set oDataSrv = Nothing
    Set objLookup = Nothing
    Set dFilter = Nothing
    Set rs = Nothing
ExitFun:

End Sub

Private Sub m_BillTransfer_LoadBillEnd(ByVal ShowType As Long)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� LoadBillEnd
    
End Sub
    