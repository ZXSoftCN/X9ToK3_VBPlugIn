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
'/*描述:多级审核前扩展处理方法
'/*参数:
'/*@ nClassID           Long            业务对象事务类型
'/*@ nFBillID           Long            业务单据内码
'/*@ nFPage             Long            业务单据页码
'/*@ nFBillEntryID      Long            业务单据分录内码
'/*@ dctBillCheckRecord KFO.Dictionary  审核路线数据包
'/*@ sDsn               String          K/3标准数据库连接串
'/*返回值:
'/*返回True时表示扩展处理方法成功执行，返回False则表示扩展处理方法未成功执行
'/********************************************************************/
        
        'TODO: 请在此处添加代码响应事件 BeforeMultiCheck
        
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
        Dim HookErrMsgInner As String
        
        On Error GoTo err_Handle
        'GetObjectContext.SetComplete
        
        HookInvokeInnter = True
        BeforeMultiCheck = True
        HookErrMsgInner = ""
        
        '当 FModel为0，审核时，OverBusinessLevel为True，则表示单据已审核通过
        '当 FModel为2，弃审时，OverBusinessLevel为False，则表示单据已弃审通过
        '以上都不需要拦截响应
        
        If CInt(dctBillCheckRecord("FMode")) = 0 And CBool(dctBillCheckRecord("OverBusinessLevel")) Then
            Exit Function
        End If
        If CInt(dctBillCheckRecord("FMode")) = 2 And CBool(dctBillCheckRecord("OverBusinessLevel")) = False Then
            Exit Function
        End If
        
        'CInt(dctBillCheckRecord("FBusinessLevel")) 当前审核级别
        'CInt(dctBillCheckRecord("FMaxLevel")) 终审级别
        
        iEntryID = 0
        EventID = 200001 '审核前
        lngFInterID = nFBillID
        sBillNo = ""
        lROB = 0
        lngTranType = nClassID
        
        iPos = InStr(1, sDsn, "}", vbTextCompare)

        lngOperate = 0 '默认 未知动作
        If CInt(dctBillCheckRecord("FMode")) = 0 Then
            lngOperate = 1 '审核
        End If
        
        If CInt(dctBillCheckRecord("FMode")) = 2 Then
            lngOperate = 2 '反审核
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
        
        Call DoImport(sConn, lngTranType, lROB, lngOperate, EventID, lngFInterID, iEntryID, sBillNo, sCurrUser, "", HookInvokeInnter, HookErrMsgInner)
        BeforeMultiCheck = HookInvokeInnter
        If Not BeforeMultiCheck Then
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile
        End If
        
        Exit Function
err_Handle:
    BeforeMultiCheck = False
    'MsgBox "X9到K3业务拦截中间件错误:" & Err.Description
    Err.Description = HookErrMsgInner
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile '必须再抛出Err，才能被K3主程序截获抛出自定义的错误信息
End Function
 
 

Public Sub DoImport(ByVal sDsn As String, ByVal lngTranType As Long, ByVal lROB As Integer, ByVal lngOperate As Long, _
                            ByVal EventID As Long, ByVal lngFInterID As Long, ByVal EntryID As Integer, ByVal sBillNo As String, _
                            ByVal currUser As String, ByVal sData As String, ByRef HookInvokeInnter As Boolean, ByRef HookErrorMsg As String)
    On Error GoTo errhandle
    Dim objVoucherImplent As Object
    
    Set objVoucherImplent = CreateObject("K3ToX9BillTransfer.InterceptHandler")
    
    Call objVoucherImplent.Handle(sDsn, lngTranType, lROB, lngOperate, EventID, lngFInterID, EntryID, sBillNo, currUser, sData, HookInvokeInnter, HookErrorMsg)
    Exit Sub
    
errhandle:
    HookInvokeInnter = False
    
End Sub

