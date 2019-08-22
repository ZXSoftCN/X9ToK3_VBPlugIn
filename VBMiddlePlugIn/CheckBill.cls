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
 
    '定义插件对象接口. 必须具有的声明, 以此来获得事件
    '注意: 此方法必须存在, 请勿修改
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
    
    Dim HookInvokeInner As Boolean
    Dim HookErrMsgInner As String
    
    
    On Error GoTo err_Handle
    'GetObjectContext.SetComplete
    HookInvokeInner = True
    HookInvoke = True
    HookErrMsgInner = ""
    
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
    
    iEntryID = 0 '除了行关闭/行反关闭外，默认iEntryID为0
    If EventID = 200001 Or EventID = 200003 Then
        '审核前事件/反审核事件
        lngOperate = CLng(dctParams("OperateCode"))
        lngFInterID = CLng(dctParams("InterID"))
    
    ElseIf EventID = 300007 Or EventID = 300008 Then
        '关闭前事件/反关闭事件：自定义Operate为64，128，以便合并统一用Operate处理。
        If CInt(dctParams("CloseSwitch")) = 1 Then
            '反关闭
            lngOperate = 128
        Else
             '关闭
            lngOperate = 64
        End If
        
        lngFInterID = CLng(dctParams("BillInterID"))
        '关闭事件通过TableHeadName设定lngTranType
        If CStr(dctParams("TableHeadName")) = "POOrder" Then
            lngTranType = 71
        ElseIf CStr(dctParams("TableHeadName")) = "POInstock" Then
            lngTranType = 72
        ElseIf CStr(dctParams("TableHeadName")) = "SEOutStock" Then
            lngTranType = 83
        ElseIf CStr(dctParams("TableHeadName")) = "SEOrder" Then
            lngTranType = 81
        End If
        
    ElseIf EventID = 300015 Or EventID = 300016 Then
        '行关闭前事件/反行关闭事件
        If CInt(dctParams("CloseSwitch")) = 1 Then
            '反关闭
            lngOperate = 128
        Else
             '关闭
            lngOperate = 64
        End If
        lngFInterID = CLng(dctParams("BillInterID"))
        iEntryID = CInt(dctParams("BillEntryID"))
        
        '行关闭事件通过TableHeadName设定lngTranType
        If CStr(dctParams("TableHeadName")) = "POOrder" Then
            lngTranType = 71
        ElseIf CStr(dctParams("TableHeadName")) = "SEOrder" Then
            lngTranType = 81
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
    
    Call DoImport(sDsn, lngTranType, lROB, lngOperate, EventID, lngFInterID, iEntryID, sBillNo, sCurrUser, "", HookInvokeInner, HookErrMsgInner)
    HookInvoke = HookInvokeInner
    
    'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile
    
    'Set objVoucherImplent = CreateObject("K3ToX9BillTransfer.InterceptHandler")
    
   'Call objVoucherImplent.Handle(sDsn, lngTranType, lROB, EventID, lngFInterID, sBillNo, "")
    
    Select Case EventID
'        Case 200002 '审核反写事件
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
        Case 200003 '反审核事件
    Case Else
    End Select
    If Not HookInvoke Then
        'dctParams("ErrDescription") = "X9效验不通过"
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile
    End If
'insert into t_ThirdPartyComponent VALUES (2,200003,2,'','K3MThirdPartyCheck.CheckBill','')
    Exit Function
err_Handle:
    HookInvoke = False
    'MsgBox "X9到K3业务拦截中间件错误:" & Err.Description
    'Err.Description = "X9到K3业务拦截中间件错误123" '定义抛出的错误信息
    Err.Description = HookErrMsgInner
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile '必须再抛出Err，才能被K3主程序截获抛出自定义的错误信息
    'Err.Raise dctParams.GetValue("ErrCode", 0), dctParams.GetValue("ErrSource", ""), dctParams.GetValue("ErrDescription", "")
    
End Function

Public Sub DoImport(ByVal sDsn As String, ByVal lngTranType As Integer, ByVal lROB As Integer, ByVal lngOperate As Long, _
                            ByVal EventID As Long, ByVal lngFInterID As Long, ByVal EntryID As Integer, ByVal sBillNo As String, _
                            ByVal currUser As String, ByVal sData As String, ByRef HookInvokeInner As Boolean, ByRef HookErrorMsg As String)
    On Error GoTo errhandle
    Dim objVoucherImplent As Object
    
    Set objVoucherImplent = CreateObject("K3ToX9BillTransfer.InterceptHandler")
    
   Call objVoucherImplent.Handle(sDsn, lngTranType, lROB, lngOperate, EventID, lngFInterID, EntryID, sBillNo, currUser, sData, HookInvokeInner, HookErrorMsg)
    
    
errhandle:
    HookInvokeInner = HookInvokeInner
    
End Sub
    

