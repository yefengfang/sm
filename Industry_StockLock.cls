VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_StockLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is Industry Bill Interface Class, made by K3BOSPLUGINSWIZAED"
Public Function MainFunction(ByVal sKey As String, oList As Object, ByRef bCancel As Boolean)
    '菜单响应
    Select Case sKey
        Case "StockLock" '和你在数据库里插入的FNAME一样
            MsgBox ("hello")
        
        '通过Set vectBill = oList.GetSelected 可以获取当前选中序时薄数据
        
        '返回记录集方式
        'Set rs = obj.Execute("select * from t_icitem")
        
        '执行存储过程方式
        ' obj.Execute3 ("exec KY_PlanQty")
        
        
        
        
        
        
            'Dim vectBill As KFO.Vector
    'Dim lmul As Long
    ' Dim rs As ADODB.Recordset
    'Dim InBatch As Form
    'Set InBatch = New InBatch
    
    'Set OBJ = CreateObject("K3Connection.AppConnection")
    

    End Select
End Function
