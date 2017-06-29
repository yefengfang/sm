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
    '菜单响应
    
    Select Case sKey
        Case "BathLockStock" '和你在数据库里插入的FNAME一样
            '配置文件内容
            Dim text As String
            text = ""
            '读取被选择的单据内码和分录内码
            For i = 1 To oList.ListSelectBillInfo.Size
                text = text + "I" + CStr(oList.ListSelectBillInfo(i)("FInterID"))
                text = text + "E" + CStr(oList.ListSelectBillInfo(i)("FEntryID"))
            Next
            
            If Dir(App.Path & "\temp.txt") = "" Then
            '不存在
                Open (App.Path & "\temp.txt") For Output As #1
                Print #1, text
                Close #1
            Else
            '存在
                Kill (App.Path & "\temp.txt")
                Open (App.Path & "\temp.txt") For Output As #1
                Print #1, text
                Close #1
            End If
            frm_LockStock.Show vbModal
    End Select
End Function
