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

Private Sub m_BillTransfer_HeadChange(ByVal CtlIndex As Long, ByVal Value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� HeadChange
    
'SELECT
'    I.FNumber,B.FSupplyID,SUM(BE.FConsignAmount) AS SBAmount,
'    SUM(RB.FSettleAmount) AS RBAmount,
'    (SUM(BE.FConsignAmount)-SUM(RB.FSettleAmount)) AS YSAmount
'FROM ICStockBill AS B
'INNER JOIN ICStockBillEntry AS BE ON B.FInterID=BE.FInterID
'FULL JOIN t_RP_NewReceiveBill RB ON RB.FCustomer=B.FSupplyID
'INNER JOIN t_Item AS I ON I.FItemID=ISNULL(B.FSupplyID,RB.FCustomer)
'WHERE B.FTranType=21 AND I.FNumber=''
'GROUP BY I.FNumber,B.FSupplyID

 
    With m_BillTransfer
        On Error GoTo ExFun
        Set cn = New ADODB.Connection
        cn.CursorLocation = adUseClient
        cn.Open .Cnnstring
        Set rs = New ADODB.Recordset
        Dim sql As String
        sql = "" & _
"SELECT" & _
"    B.FSupplyID,SUM(BE.FConsignAmount) AS SBAmount, " & _
"    SUM(RB.FSettleAmount) AS RBAmount, " & _
"    (SUM(BE.FConsignAmount)-SUM(RB.FSettleAmount)) AS YSAmount " & _
"FROM ICStockBill AS B " & _
"INNER JOIN ICStockBillEntry AS BE ON B.FInterID=BE.FInterID " & _
"FULL JOIN t_RP_NewReceiveBill RB ON RB.FCustomer=B.FSupplyID " & _
"INNER JOIN t_Item AS I ON I.FItemID=ISNULL(B.FSupplyID,RB.FCustomer) " & _
"WHERE B.FTranType=21 AND I.FNumber='" + Value + "' " & _
"GROUP BY B.FSupplyID "

        rs.Open sql, cn
        
        If (.m_HeadCtl(CtlIndex).Caption = "������λ:") Then
            'MsgBox ("SQL")
            For i = LBound(.m_HeadCtl) To UBound(.m_HeadCtl)
                If (.m_HeadCtl(i).Caption = "Ӧ�����") Then
                    r = .SetHead(i, rs("YSAmount"))
                    Exit For
                End If
                
            Next
        End If
        
    End With
 
ExFun:

End Sub
Private Sub m_BillTransfer_HookInvoke(ByVal pEventID As Long, Para As KFO.IDictionary)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� HookInvoke
    With m_BillTransfer
    MsgBox ("Hello")
        For i = LBound(.m_EntryCtl) To UBound(.m_EntryCtl)
            If (.m_EntryCtl(i).Caption = "����") Then
                    MsgBox (i)
                Exit For
            End If
        Next
    End With
 

End Sub