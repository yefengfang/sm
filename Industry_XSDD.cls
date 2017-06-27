VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_XSDD"
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

Private Sub m_BillTransfer_HeadChange(ByVal CtlIndex As Long, ByVal Value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
 
    'TODO: 请在此处添加代码响应事件 HeadChange
    With m_BillTransfer
        On Error GoTo ExFun
        Set cn = New ADODB.Connection
        cn.CursorLocation = adUseClient
        cn.Open .Cnnstring
        Set rs = New ADODB.Recordset
        Dim sql As String
        
        If (.m_HeadCtl(CtlIndex).Caption = "购货单位:") Then
            sql = "SELECT fendbalance FROM V_CustomerBalance WHERE FCustomer=" + CStr(.m_HeadCtl(CtlIndex).InterID)
            rs.Open sql, cn
            For I = LBound(.m_HeadCtl) To UBound(.m_HeadCtl)
                If (.m_HeadCtl(I).Caption = "应收账款") Then
                    r = .SetHead(I, rs("fendbalance"))
                    Exit For
                End If
                
            Next
        End If
    End With
 
ExFun:
    Set rs = Nothing
    cn.Close
    Set cn = Nothing

End Sub
