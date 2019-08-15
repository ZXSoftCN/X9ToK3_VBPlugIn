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
    Dim lngTranType As Long
    Dim lngFInterID As Long
    
    Dim sDsn    As String
    Dim iPos As Long
    Dim cn As ADODB.Connection
    Dim strMrpClose As String
    Dim strMrpOpen As String
    Dim strRecordOrg As String
    Dim strMrpUpdate As String
    
'    MsgBox "OK"
    Set cn = New ADODB.Connection
    
      'cn.Execute "select" & rowcount & "=count(*) from icstockbillentry where finterid=" & Para("InterID")

    On Error GoTo err_Handle
    'GetObjectContext.SetComplete

    HookInvoke = True

    lngOperate = CLng(dctParams("OperateCode"))
    lngTranType = CLng(dctParams("TranType"))
    lngFInterID = CLng(dctParams("InterID"))

'    strRecordOrg = "set nocount on " & vbCrLf _
'    & "if not exists(select 1 from sysobjects where xtype = 'U' and name = 't_dgkd_orderMrp')" & vbCrLf _
'    & "select ou.FInterID,ou.FEntryID,ou.FMRPClosed into t_dgkd_orderMrp from seorderentry ou where 1 = 0" & vbCrLf _
'    & "go" & vbCrLf
    strRecordOrg = "delete om from t_dgkd_orderMrp om join icstockbillentry u on u.FOrderInterID = om.FInterID where u.FInterID = " & lngFInterID & vbCrLf _
    & "Insert into t_dgkd_orderMrp(FInterID,FEntryID,FMrpClosed)" & vbCrLf _
    & "select distinct u.FOrderInterID,u.FOrderEntryID,ou.FMRPClosed from seorderentry ou join icstockbillentry u on u.FOrderInterID = ou.FInterID and u.FOrderEntryID = ou.FEntryID where u.FInterID = " & lngFInterID
    
    strMrpUpdate = "update ou " & vbCrLf _
    & "set ou.FMrpClosed = om.FMrpClosed" & vbCrLf _
    & "from seorderentry ou join t_dgkd_orderMrp om on ou.FInterID = om.FInterID and ou.FEntryID = om.FEntryID" & vbCrLf _
    & "join icstockbillentry u on u.FOrderInterID = om.FInterID and u.FOrderEntryID = om.FEntryID" & vbCrLf _
    & "where u.FInterID = " & lngFInterID
    
    
    iPos = InStr(1, dctParams("sDsn"), "}", vbTextCompare)

    If iPos > 17 Then
        sDsn = Mid(dctParams("sDsn"), 16, iPos - 16)
    End If
    
    With cn
        .ConnectionString = sDsn
        .CursorLocation = adUseClient
    End With
    cn.Open
    
    
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
            If dctParams("ROB") = -1 Then
                Select Case lngTranType
                    Case 21 '销售出库单
                        Select Case lngOperate
                            Case 1
                                cn.Execute strRecordOrg
                            Case 2
                                cn.Execute strRecordOrg
                            Case Else
                        End Select
                    Case Else
                End Select
            End If
        Case 200003 '审核退出前事件
            If dctParams("ROB") = -1 Then '红单
                Select Case lngTranType
                        Case 21 '销售出库单
                            Select Case lngOperate
                                Case 1
                                    cn.Execute strMrpUpdate
                                Case 2
                                    cn.Execute strMrpUpdate
                                Case Else
                            End Select
                 
                        Case Else
                End Select
            End If
'            If dctParams("ROB") = 1 Then '蓝单
'                Select Case lngTranType
'                   Case 21 '销售出库单
'                       Select Case lngOperate
'                           Case 1
'                             '审核前需要记录销售订单的状态，有可能需要用到临时表，需要根据你们的情况而定，
'                             cn.Execute strMrpclose & lngFInterID
'
'                             dictparams("FStatus") = rs.Fields("FStatus").Value
'                               MsgBox "调拨单审核" & rs.Fields("FBillNO").Value
'
'                               '审核前
'                           Case 2
'                               MsgBox "调拨单反审核"
'                               cn.Execute strMrpOpen & lngFInterID
'                           Case Else
'                       End Select
'
'                   Case Else
'                End Select
'                cn.Execute "select top 1 * from t_IcItem"
                
'            End If
    Case Else
    End Select
    If Not HookInvoke Then
        Err.Raise dctParams.GetValue("ErrCode", 0), dctParams.GetValue("ErrSource", ""), dctParams.GetValue("ErrDescription", "")
    End If
'insert into t_ThirdPartyComponent VALUES (2,200003,2,'','K3MThirdPartyCheck.CheckBill','')
    Exit Function
err_Handle:
    HookInvoke = False
    MsgBox "调拨单审核失败　"
    dctParams("ErrDescription") = Err.Description
    
End Function
