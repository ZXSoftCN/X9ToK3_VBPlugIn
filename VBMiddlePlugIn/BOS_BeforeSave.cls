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
'/*����:���ݱ���ǰ��չ��������
'/*����:
'/*@ dctClassType       KFO.Dictionary  ����������Ϣģ��
'/*@ dctData            KFO.Dictionary  �������ݰ� dctData("FID"). Ϊ�����ĵ�������
'/*@ dctLink            KFO.Dictionary  ѡ����ϵģ�壬��ѡ��ΪNothing
'/*@ dctTableInfo       KFO.Dictionary  �����ֶ�ģ����Ϣ
'/*@ vctClassTypeEntry  KFO.Vector      ������ģ����Ϣ
'/*@ sDsn               String          K/3��׼���ݿ����Ӵ�
'/*��ע:���Խ��õ��ݴ����������Ĺ����ڴ˺����д���
'/********************************************************************/
                           
        'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� BeforeSave
        Dim iEntryID As Integer
        Dim lInterID As Long
        
        iEntryID = 0

End Function
 
 
 
