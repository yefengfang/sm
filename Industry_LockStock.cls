VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_LockStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is Industry Bill Interface Class, made by K3BOSPLUGINSWIZAED"

Public Function MainFunction(ByVal sKey As String, oList As Object, ByRef bCancel As Boolean)
    '�˵���Ӧ
    
    Select Case sKey
        Case "BathLockStock" '���������ݿ�������FNAMEһ��
            '�����ļ�����
            Dim text As String
            text = ""
            '��ȡ��ѡ��ĵ�������ͷ�¼����
            For i = 1 To oList.ListSelectBillInfo.Size
                text = text + "I" + CStr(oList.ListSelectBillInfo(i)("FInterID"))
                text = text + "E" + CStr(oList.ListSelectBillInfo(i)("FEntryID"))
            Next
            
            If Dir(App.Path & "\temp.txt") = "" Then
            '������
                Open (App.Path & "\temp.txt") For Output As #1
                Print #1, text
                Close #1
            Else
            '����
                Kill (App.Path & "\temp.txt")
                Open (App.Path & "\temp.txt") For Output As #1
                Print #1, text
                Close #1
            End If
            frm_LockStock.Show vbModal
    End Select
End Function