VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BeforeMultiCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is MiddleWareEvents Interface Class [ BeforeMultiCheck ], made by K3BOSPLUGINSWIZAED"
Public Function BeforeMultiCheck(ByVal sDsn As String, _
                                 ByVal nClassID As Long, _
                                 ByVal nFBillID As Long, _
                                 Optional ByVal nFPage As Long = 1, _
                                 Optional ByVal nFBillEntryID As Long = 0, _
                                 Optional ByVal dctBillCheckRecord As KFO.Dictionary) As Boolean


'/********************************************************************/
'/*����:�༶���ǰ��չ��������
'/*����:
'/*@ nClassID           Long            ҵ�������������
'/*@ nFBillID           Long            ҵ�񵥾�����
'/*@ nFPage             Long            ҵ�񵥾�ҳ��
'/*@ nFBillEntryID      Long            ҵ�񵥾ݷ�¼����
'/*@ dctBillCheckRecord KFO.Dictionary  ���·�����ݰ�
'/*@ sDsn               String          K/3��׼���ݿ����Ӵ�
'/*����ֵ:
'/*����Trueʱ��ʾ��չ���������ɹ�ִ�У�����False���ʾ��չ��������δ�ɹ�ִ��
'/********************************************************************/
        
        'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� BeforeMultiCheck
        
        Dim lngOperate As Long
        Dim lngTranType As Long
        Dim lngFInterID As Long
        Dim iEntryID As Integer
        Dim lROB As Integer
        Dim sBillNo As String
        Dim sCurrUser As String
        Dim iCurrUserStart As Integer
        Dim iCurrUserEnd As Integer
        Dim EventID As Long
        
        Dim iPos As Long
        Dim sConn As String
        Dim HookInvokeInnter As Boolean
        
        On Error GoTo err_Handle
        'GetObjectContext.SetComplete
        
        HookInvokeInnter = True
        BeforeMultiCheck = True
        
        '�� FModelΪ0�����ʱ��OverBusinessLevelΪTrue�����ʾ���������ͨ��
        '�� FModelΪ2������ʱ��OverBusinessLevelΪFalse�����ʾ����������ͨ��
        '���϶�����Ҫ������Ӧ
        
        If CInt(dctBillCheckRecord("FMode")) = 0 And CBool(dctBillCheckRecord("OverBusinessLevel")) Then
            Exit Function
        End If
        If CInt(dctBillCheckRecord("FMode")) = 2 And CBool(dctBillCheckRecord("OverBusinessLevel")) = False Then
            Exit Function
        End If
        
        'CInt(dctBillCheckRecord("FBusinessLevel")) ��ǰ��˼���
        'CInt(dctBillCheckRecord("FMaxLevel")) ���󼶱�
        
        iEntryID = 0
        EventID = 200001 '���ǰ
        lngFInterID = nFBillID
        sBillNo = ""
        lROB = 0
        lngTranType = nClassID
        
        iPos = InStr(1, sDsn, "}", vbTextCompare)

        lngOperate = 0 'Ĭ�� δ֪����
        If CInt(dctBillCheckRecord("FMode")) = 0 Then
            lngOperate = 1 '���
        End If
        
        If CInt(dctBillCheckRecord("FMode")) = 2 Then
            lngOperate = 2 '�����
        End If
        
        
        If iPos > 17 Then
            sConn = Mid(sDsn, 63, iPos - 63)
        End If
    
        iCurrUserStart = InStr(1, sDsn, "UserName=", vbTextCompare)
        iCurrUserEnd = InStr(iCurrUserStart, sDsn, ";", vbTextCompare)
        If iCurrUserEnd > iCurrUserStart Then
            sCurrUser = Mid(sDsn, iCurrUserStart + Len("UserName="), iCurrUserEnd - Len("UserName=") - iCurrUserStart)
        Else
            sCurrUser = "test"
        End If
        
        'iCurrUserStart = InStr(1, dctParams("sDsn"), "UserName=", vbTextCompare)
        'iCurrUserEnd = InStr(iCurrUserStart, dctParams("sDsn"), ";", vbTextCompare)
        'If iCurrUserEnd > iCurrUserStart Then
        '    sCurrUser = Mid(dctParams("sDsn"), iCurrUserEnd - Len("UserName=") - 2, iCurrUserEnd - Len("UserName=") - iCurrUserStart)
        'Else
        '    sCurrUser = "test"
        'End If
        
        Call DoImport(sConn, lngTranType, lROB, lngOperate, EventID, lngFInterID, iEntryID, sBillNo, sCurrUser, "", HookInvokeInnter)
        BeforeMultiCheck = HookInvokeInnter
        
        Exit Function
err_Handle:
    BeforeMultiCheck = False
    MsgBox "X9��K3ҵ�������м������:" & Err.Description
End Function
 
 

Public Sub DoImport(ByVal sDsn As String, ByVal lngTranType As Long, ByVal lROB As Integer, ByVal lngOperate As Long, _
                            ByVal EventID As Long, ByVal lngFInterID As Long, ByVal EntryID As Integer, ByVal sBillNo As String, _
                            ByVal currUser As String, ByVal sData As String, ByRef HookInvokeInnter As Boolean)
    On Error GoTo errhandle
    Dim objVoucherImplent As Object
    
    Set objVoucherImplent = CreateObject("K3ToX9BillTransfer.InterceptHandler")
    
    Call objVoucherImplent.Handle(sDsn, lngTranType, lROB, lngOperate, EventID, lngFInterID, EntryID, sBillNo, currUser, sData, HookInvokeInnter)
    Exit Sub
    
errhandle:
    HookInvokeInnter = False
    
End Sub
