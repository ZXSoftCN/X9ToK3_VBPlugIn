VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BeforeSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is MiddleWareEvents Interface Class [ BeforeSave ], made by K3BOSPLUGINSWIZAED"
Public Function BeforeSave(ByVal sDsn As String, _
                           ByVal dctClassType As KFO.Dictionary, _
                           ByVal vctClassTypeEntry As KFO.Vector, _
                           ByVal dctTableInfo As KFO.Dictionary, _
                           ByVal dctData As KFO.Dictionary, _
                           ByVal dctLink As KFO.Dictionary)

'/********************************************************************/
'/*描述:单据保存前扩展处理方法
'/*参数:
'/*@ dctClassType       KFO.Dictionary  单据整体信息模板
'/*@ dctData            KFO.Dictionary  单据数据包 dctData("FID"). 为保存后的单据内码
'/*@ dctLink            KFO.Dictionary  选单关系模板，无选单为Nothing
'/*@ dctTableInfo       KFO.Dictionary  单据字段模板信息
'/*@ vctClassTypeEntry  KFO.Vector      单据体模板信息
'/*@ sDsn               String          K/3标准数据库连接串
'/*备注:可以将该单据触发器处理的过程在此函数中处理
'/********************************************************************/
                           
        'TODO: 请在此处添加代码响应事件 BeforeSave
        Dim iEntryID As Integer
        Dim lInterID As Long
        
        iEntryID = 0

End Function
 
 
 

