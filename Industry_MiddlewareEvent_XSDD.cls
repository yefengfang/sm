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
 
    '����������ӿ�. ������е�����, �Դ�������¼�
    'ע��: �˷����������, �����޸�
    Select Case EventID
        Case 200001 '���ǰ�¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'InterID                | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateDelete          | Ĭ��ֵΪ0
            'ROB                    | �Ƿ��Ǻ쵥
            'OperateCode            | ������:   1------���
            '                       |           2------�����
            '                       |           4------����
            '                       |           8------ɾ��
            '                       |           16-----������
            '                       |           32-----����
            '--------------------------------------------------------
            
            '�����������Ӵ��� '���ǰ' ����

 
    Case Else
    End Select
 
End Function