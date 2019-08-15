VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_AfterMultiCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is MiddleWareEvents Interface Class [ AfterMultiCheck ], made by K3BOSPLUGINSWIZAED"
Public Function AfterMultiCheck(ByVal sDsn As String, _
                                 ByVal nClassID As Long, _
                                 ByVal nFBillID As Long, _
                                 Optional ByVal nFPage As Long = 1, _
                                 Optional ByVal nFBillEntryID As Long = 0, _
                                 Optional ByVal dctBillCheckRecord As KFO.Dictionary) As Boolean

'/********************************************************************/
'/*描述:多级审核后扩展处理方法
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
                            
        'TODO: 请在此处添加代码响应事件 AfterMultiCheck
        AfterMultiCheck = True

End Function
 
 
 

