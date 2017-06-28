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

'已废弃，使用HeadCtlClick事件可以避免单据打开后无法选择的问题
'Private Sub m_BillInterface_Change(ByVal dct As KFO.IDictionary, ByVal dctFld As KFO.IDictionary, ByVal Col As Long, ByVal Row As Long, Cancel As Boolean)
'
'    'TODO: 请在此处添加代码响应事件
'    If (dct("FCaption") = "核算项目") Then
'        '对供应商名称变动事件处理
'        With m_BillInterface
'            '获取数据
'            Dim CMD As String
'            CMD = "SELECT TOP (1000) [FItemID], [F_105] ,[FName] FROM [AISCJCS].[dbo].[t_Item_3016] WHERE F_105=" + dctFld("FFLD")
'            Set rs = m_BillInterface.K3Lib.GetData(CMD)
'            '循环拼节字符串
'            Dim liststr As String
'            liststr = "VALUELIST{"
'            If (Not rs.EOF) Then
'                For I = 1 To rs.RecordCount
'                    On Error Resume Next
'                    liststr = liststr + CStr(rs("FItemID")) + "=" + rs("FName") + ","
'                    rs.MoveNext
'                Next
'            End If
'            liststr = liststr + "}"
'
'            '反写回下拉框
'            Dim ICtlIndex As Long
'            ICtlIndex = .TableInfo(.TableInfo("Map")("FComboBox1"))("CtlIndex")
'
'            '此为隐藏函数，手册中没有
'            Call m_BillInterface.BillCtl.SetComboxList(ICtlIndex, liststr)
'        End With
'    ElseIf (dct("FCaption") = "收款账号") Then
'        '对下拉框变动事件处理
'        With m_BillInterface
'            '获取数据
'            CMD = "SELECT TOP 1 [F_102] ,[F_103],[F_104] FROM [AISCJCS].[dbo].[t_Item_3016] WHERE FItemID='" + dctFld("FFLD") + "'"
'            Set rs = .K3Lib.GetData(CMD)
'            '写入字段
'            a = .SetFieldValue("FText6", rs.Fields("F_102"))
'            a = .SetFieldValue("FText7", rs.Fields("F_103"))
'            a = .SetFieldValue("FText8", rs.Fields("F_104"))
'        End With
'    End If
'
'
'    '已知问题： 保存后在打开会导致下拉框无法选择，因为保存的是核算项目内码
'    '
'End Sub
Private Sub m_BillInterface_HeadCtlClick(ByVal ctl As Object, ByVal dct As KFO.Dictionary)
    On Error GoTo ExitSub1:
    If (dct("FCaption") = "收款账号") Then
        With m_BillInterface
            Dim CMD As String
            sKey = .TableInfo("Map")("FCustomer")
            CMD = "SELECT TOP (1000) [FItemID], [F_105] ,[FName] FROM [AISCJCS].[dbo].[t_Item_3016] WHERE F_105=" + .Data("Page1")(sKey)("FFLD")
            Set rs = m_BillInterface.K3Lib.GetData(CMD)
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
