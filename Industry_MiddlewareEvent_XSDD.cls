VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_MiddlewareEvent_XSDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is IndustryMiddleWareEvents Interface Class, made by K3BOSPLUGINSWIZAED"
 
Public Function HookInvoke(ByVal EventID As Long, dctParams As KFO.Dictionary) As Long
'
'   功能说明：不允许审核单价为0的销售订单
'
 
    '定义插件对象接口. 必须具有的声明, 以此来获得事件
    '注意: 此方法必须存在, 请勿修改
    Select Case EventID
        Case 200001 '审核前事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'InterID                | 单据的内码
            'TranType               | 单据的事务类型
            'OperateDelete          | 默认值为0
            'ROB                    | 是否是红单
            'OperateCode            | 操作码:   1------审核
            '                       |           2------反审核
            '                       |           4------保存
            '                       |           8------删除
            '                       |           16-----反作废
            '                       |           32-----作废
            '--------------------------------------------------------
            
            '请在这里添加处理 '审核前' 代码
            'conn sDsn
            'billinter   InterID
            '
            '
            '
            '
            If (dctParams("OperateCode") = 1 And dctParams("TranType") = 81) Then
            On Error GoTo UnknowError
                Set cn = New ADODB.Connection
                cn.CursorLocation = adUseClient
                connStr = Mid(dctParams("sDsn"), InStr(dctParams("sDsn"), "{") + 1, InStr(dctParams("sDsn"), "}") - (InStr(dctParams("sDsn"), "{") + 1))
                cn.Open connStr
                Dim sql As String
                sql = "SELECT 1 From SEOrderEntry WHERE FPrice=0 AND FInterID=" + CStr(dctParams("InterID"))
                Set rs = New ADODB.Recordset
                rs.Open sql, cn
                If rs.RecordCount > 0 Then
                    h = MsgBox("订单存在价格为0的分录，确定审核？", vbYesNo)
                    If h = 6 Then
                    
                        HookInvoke = 1
                    Else
                        HookInvoke = 0
                    End If
                Else
                    HookInvoke = 1
                End If
            Else
                HookInvoke = 1
            End If
            Err.Clear
            Err.Number = 10000
            Err.Description = "订单存在价格为0的分录，退出审核！"
            GoTo ExitCase
UnknowError:
    Err.Clear
    Err.Number = 10000
    Err.Description = "未知错误，请联系管理员！"

ExitCase:
    Case Else
    End Select
 
End Function
' 已知问题：用0返回后系统会跳出调用失败的对话框
