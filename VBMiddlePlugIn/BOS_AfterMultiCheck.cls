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
'/*����:�༶��˺���չ��������
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
                            
        'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� AfterMultiCheck
        AfterMultiCheck = True

End Function
 
 
 
