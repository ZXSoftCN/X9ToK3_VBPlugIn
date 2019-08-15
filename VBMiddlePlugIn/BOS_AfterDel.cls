VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_AfterDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is MiddleWareEvents Interface Class [ AfterDel ], made by K3BOSPLUGINSWIZAED"
Public Function AfterDel(ByVal sDsn As String, _
                            ByVal nClassID As Long, _
                            ByVal dctClassType As KFO.Dictionary, _
                            ByVal nInterID As Long)

'/********************************************************************/
'/*描述:单据删除后扩展处理方法
'/*参数:
'/*@ dctClassType       KFO.Dictionary  单据整体信息模板
'/*@ nClassID           Long            对应的单据模板的分录主键值
'/*@ nInterID           Long            具体单据分录的ID键值
'/*@ sDsn               String          K/3标准数据库连接串
'/********************************************************************/
                            
        'TODO: 请在此处添加代码响应事件 AfterDel

End Function
 
 
 

