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

Private Sub m_BillInterface_Change(ByVal dct As KFO.IDictionary, ByVal dctFld As KFO.IDictionary, ByVal Col As Long, ByVal Row As Long, Cancel As Boolean)
 
    'TODO: 请在此处添加代码响应事件
    
    On Error GoTo ExitFun
    With m_BillInterface
        If (dct("FCaption") = "核算项目") Then
            '获取下拉列表控件的Index
            'Set dctTableInfo = m_BillInterface.TableInfo
'            sKey = .TableInfo("Map")("FComboBox1")
'            ICtlIndex = .TableInfo(sKey)("CtlIndex")
'            Set oBill = m_BillInterface.BillCtl
'            Call oBill.setcomboxlist(ICtlIndex, "VALUELIST{1=a,2=b,3=c}")
            
'            Dim CMD As String
'            CMD = "SELECT TOP (1000) [F_105] ,[FName] FROM [AISCJCS].[dbo].[t_Item_3016] WHERE F_105=" + CStr(dctFld("FFLD"))
'            Set rs = .K3Lib.GetData(CMD)
'            If (rs.RecordCount > 0) Then
'                .Data("page1")("_0064_FComboBox1").AddItem = "aabb"
'                'a = .SetFieldValue("FComboBox1", rs.Fields("FName"))
''                a = .SetFieldValue("FText5", rs.Fields("KHM"))  .data("page1")("_0064_FComboBox1").
''                a = .SetFieldValue("FText2", rs.Fields("KHM_S"))
''                a = .SetFieldValue("FText3", rs.Fields("KHYH_S"))
''                a = .SetFieldValue("FText4", rs.Fields("YHZH_S"))
'            Else
'                MsgBox ("未找到对应核算项目！")
'            End If


        Dim dct As KFO.Dictionary
        Dim dctPage As KFO.Dictionary
        Dim oGrid As Object
        dctPage("FPage") = 1
        Set dctPage = New KFO.Dictionary
        Set oGrid = m_BillInterface.BillCtl.GetObject("CTRL_GRID", dctPage)
        Set dct = m_BillInterface.GetFieldInfoByKey("FComboxList1", "", 0)
        Dim CMD As String
        CMD = "SELECT TOP (1000) [F_105] ,[FName] FROM [AISCJCS].[dbo].[t_Item_3016] WHERE F_105=" + CStr(dctFld("FFLD"))
        Set rs = .K3Lib.GetData(CMD)
        Dim strList
        If Not rs.EOF Then
        For i = 1 To rs.Count
        Str

        

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
 
    'TODO: 请在此处添加代码响应事件 LoadBillEnd
    
End Sub
    
