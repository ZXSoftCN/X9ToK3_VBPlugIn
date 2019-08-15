VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CheckBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is IndustryMiddleWareEvents Interface Class, made by K3BOSPLUGINSWIZAED"
 
Public Function HookInvoke(ByVal EventID As Long, dctParams As KFO.Dictionary) As Long
 
    '����������ӿ�. ������е�����, �Դ�������¼�
    'ע��: �˷����������, �����޸�
    Dim lngOperate As Long
    Dim lngTranType As Integer
    Dim lngFInterID As Long
    Dim iEntryID As Integer
    Dim lROB As Integer
    Dim sBillNo As String
    Dim sCurrUser As String
    Dim iCurrUserStart As Integer
    Dim iCurrUserEnd As Integer
    
    
    Dim sDsn    As String
    Dim iPos As Long
    
    Dim HookInvokeInnter As Boolean
    
    
    On Error GoTo err_Handle
    'GetObjectContext.SetComplete
    HookInvokeInnter = True
    HookInvoke = True
    
    If dctParams.Lookup("BillCode") Then
        sBillNo = CStr(dctParams("BillNo"))
    Else
        sBillNo = ""
    End If
    
    If dctParams.Lookup("ROB") Then
        lROB = CStr(dctParams("ROB"))
    Else
        lROB = 0
    End If
    
    If dctParams.Lookup("TranType") Then
        lngTranType = CStr(dctParams("TranType"))
    Else
        lngTranType = 0
    End If
    
    iEntryID = 0 '�����йر�/�з��ر��⣬Ĭ��iEntryIDΪ0
    If EventID = 200001 Or EventID = 200003 Then
        '���ǰ�¼�/������¼�
        lngOperate = CLng(dctParams("OperateCode"))
        lngFInterID = CLng(dctParams("InterID"))
    
    ElseIf EventID = 300007 Then
        '�ر�ǰ�¼�/���ر��¼����Զ���OperateΪ64��128���Ա�ϲ�ͳһ��Operate������
        If CInt(dctParams("CloseSwitch")) = 1 Then
            '���ر�
            lngOperate = 128
        Else
             '�ر�
            lngOperate = 64
        End If
        
        lngFInterID = CLng(dctParams("BillInterID"))
        '�ر��¼�ͨ��TableHeadName�趨lngTranType
        If CStr(dctParams("TableHeadName")) = "POOrder" Then
            lngTranType = 71
        End If
        
    ElseIf EventID = 300015 Then
        '�йر�ǰ�¼�/���йر��¼�
        If CInt(dctParams("CloseSwitch")) = 1 Then
            '���ر�
            lngOperate = 128
        Else
             '�ر�
            lngOperate = 64
        End If
        lngFInterID = CLng(dctParams("BillInterID"))
        iEntryID = CInt(dctParams("BillEntryID"))
        
        '�йر��¼�ͨ��TableHeadName�趨lngTranType
        If CStr(dctParams("TableHeadName")) = "POOrder" Then
            lngTranType = 71
        End If
    Else
        lngOperate = 0
        lngFInterID = 0
    End If
    
    iPos = InStr(1, dctParams("sDsn"), "}", vbTextCompare)

    If iPos > 17 Then
        sDsn = Mid(dctParams("sDsn"), 63, iPos - 63)
    End If
    
    iCurrUserStart = InStr(1, dctParams("sDsn"), "UserName=", vbTextCompare)
    iCurrUserEnd = InStr(iCurrUserStart, dctParams("sDsn"), ";", vbTextCompare)
    If iCurrUserEnd > iCurrUserStart Then
        sCurrUser = Mid(dctParams("sDsn"), iCurrUserStart + Len("UserName="), iCurrUserEnd - Len("UserName=") - iCurrUserStart)
    Else
        sCurrUser = "test"
    End If
    
    Call DoImport(sDsn, lngTranType, lROB, lngOperate, EventID, lngFInterID, iEntryID, sBillNo, sCurrUser, "", HookInvokeInnter)
    HookInvoke = HookInvokeInnter
    
    'Set objVoucherImplent = CreateObject("K3ToX9BillTransfer.InterceptHandler")
    
   'Call objVoucherImplent.Handle(sDsn, lngTranType, lROB, EventID, lngFInterID, sBillNo, "")
    
    Select Case EventID
'        Case 200002 '��˷�д�¼�
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
        Case 200003 '������¼�
    Case Else
    End Select
    If Not HookInvoke Then
        dctParams("ErrDescription") = "X9Ч�鲻ͨ��"
        Err.Raise dctParams.GetValue("ErrCode", 0), dctParams.GetValue("ErrSource", ""), dctParams.GetValue("ErrDescription", "")
    End If
'insert into t_ThirdPartyComponent VALUES (2,200003,2,'','K3MThirdPartyCheck.CheckBill','')
    Exit Function
err_Handle:
    HookInvoke = False
    MsgBox "X9��K3ҵ�������м������:" & Err.Description
    
    dctParams("ErrDescription") = "X9��K3ҵ�������м������"
    
End Function

Public Sub DoImport(ByVal sDsn As String, ByVal lngTranType As Integer, ByVal lROB As Integer, ByVal lngOperate As Long, _
                            ByVal EventID As Long, ByVal lngFInterID As Long, ByVal EntryID As Integer, ByVal sBillNo As String, _
                            ByVal currUser As String, ByVal sData As String, ByRef HookInvokeInnter As Boolean)
    On Error GoTo errhandle
    Dim objVoucherImplent As Object
    
    Set objVoucherImplent = CreateObject("K3ToX9BillTransfer.InterceptHandler")
    
   Call objVoucherImplent.Handle(sDsn, lngTranType, lROB, lngOperate, EventID, lngFInterID, EntryID, sBillNo, currUser, sData, HookInvokeInnter)
    
    
errhandle:
    HookInvokeInnter = HookInvokeInnter
    
End Sub
    
