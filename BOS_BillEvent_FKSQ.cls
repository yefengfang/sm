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
'
'  功能说明：动态下拉列表框
'


'定义 BillEvent 接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillInterface As Object)
 
    'BillEvent 接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = oBillInterface
 
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = Nothing

End Sub

Private Sub m_BillInterface_HeadCtlClick(ByVal ctl As Object, ByVal dct As KFO.Dictionary)
    On Error GoTo ExitSub1:
    If (dct("FCaption") = "收款账号") Then
        With m_BillInterface
            Dim CMD As String
            sKey = .TableInfo("Map")("FCustomer")
            CMD = "SELECT TOP (1000) [FItemID], [F_101] ,[FName] FROM [AISCJCS].[dbo].[t_Item_3016] WHERE F_101=" + .Data("Page1")(sKey)("FFLD")
            Set rs = m_BillInterface.K3Lib.GetData(CMD)
            If (rs.RecordCount > 0) Then
                '循环拼节字符串
                Dim liststr As String
                liststr = "VALUELIST{"
                If (Not rs.EOF) Then
                    For I = 1 To rs.RecordCount
                        On Error Resume Next
                        liststr = liststr + CStr(rs("FItemID")) + "=" + rs("FName") + ","
                        rs.MoveNext
                    Next
                End If
                liststr = liststr + "}"
    
                '反写回下拉框
                Dim ICtlIndex As Long
                ICtlIndex = .TableInfo(.TableInfo("Map")("FComboBox1"))("CtlIndex")
                '此为隐藏函数，手册中没有
                Call m_BillInterface.BillCtl.SetComboxList(ICtlIndex, liststr)
            End If
        End With
    End If
ExitSub1:
End Sub
Private Sub m_BillInterface_Change(ByVal dct As KFO.IDictionary, ByVal dctFld As KFO.IDictionary, ByVal Col As Long, ByVal Row As Long, Cancel As Boolean)
    On Error GoTo ExitSub2
    If (dct("FCaption") = "收款账号") Then
        With m_BillInterface
             '获取数据
             CMD = "SELECT TOP 1 [F_102] ,[F_103],[F_104] FROM [AISCJCS].[dbo].[t_Item_3016] WHERE FItemID='" + dctFld("FFLD") + "'"
             Set rs = .K3Lib.GetData(CMD)
             '写入字段
             If (rs.RecordCount > 0) Then
                On Error Resume Next
                a = .SetFieldValue("FText6", rs.Fields("F_102"))
                a = .SetFieldValue("FText7", rs.Fields("F_103"))
                a = .SetFieldValue("FText8", rs.Fields("F_104"))
            End If
        End With
    End If
ExitSub2:
End Sub
