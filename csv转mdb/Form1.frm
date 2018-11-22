VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdImport 
      Caption         =   "cmdImport"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'������ǰVB�������ַ����ıȽϷ�ʽΪ�ֽ���
Option Compare Binary
'Ҫ���Ա������ļ����������½�һ�����壬���������һ����Ϊ"cmdImport"�İ�ť��
'������ִ������ķ���������ļ���ȡ��д�����ݿ�Ĳ���
'������ʼ--------------------------------------------------
'��������ִ�еĿ�ʼ
Private Sub cmdImport_Click()
    Dim cn As New ADODB.Connection
    'ָ�������ַ������������ӵ� D:\db1.mdb ���ݿ�
    cn.ConnectionString = _
           "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\db1.mdb"
    '�������ӣ�׼�������ݿ����
    cn.Open
   '����洢�ļ����Ķ���
    Dim fileName As String
    'ָ��Ҫ��ȡ��CSV�ļ������Ʊ���Ϊ D:\test.csv
    fileName = "D:\test.csv"
    'ִ��ImportFile������������²�����
    '1 ��ȡ D:\test.csv �ļ�������
    '2 ����ȡ���ļ�����д�� D:\db1.mdb ���ݿ�� test ����
    ImportFile cn, "test", fileName
    '��ɲ�����ر�����
    cn.Close
    '����������Ϊ Nothing ����ʱ�ͷ�����ռ�õ��ڴ�ռ�
    Set cn = Nothing
End Sub
'��������--------------------------------------------------


'������ʼ--------------------------------------------------
'����------------------------------------------------------------------
'1. cn                          ʹ�õ����ݿ����ӣ��������ӵ� D:\db1.mdb
'2. tblName                 ʹ�õ����ݿ�����ƣ�����Ϊ test
'3. FileFullPath           ��ȡ��CSV�ļ�Ŀ¼�������ļ�����������Ϊ D:\test.csv
'4. FieldDelimiter        ָ��CSV�ļ�ͬһ�������������ַ��ָ� Ĭ��Ϊ ,����Ƕ��ţ�
'5. RecordDelimiter    ָ��CSV�ļ���ͬ�������������ַ��ָ� Ĭ��Ϊ vbCrLf���س����з���
'����------------------------------------------------------------------
Public Sub ImportFile(cn As Object, _
    ByVal tblName As String, FileFullPath As String, _
    Optional FieldDelimiter As String = ",", _
    Optional RecordDelimiter As String = vbCrLf)
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    Dim iFileNum As Integer
    Dim sFileContents As String
    Dim sTableSplit() As String
    Dim sRecordSplit() As String
    
    Dim lCtr As Integer
    Dim iCtr As Integer
    Dim lRecordCount As Long
    Dim iFieldsToImport As Integer
    
    Dim asFieldNames() As String
    Dim abFieldIsString() As Boolean
    Dim iFieldCount As Integer
    Dim sSQL As String
    
    
    '�жϴ����cn�����Ƿ������ݿ����Ӷ���������ǣ��˳�����
    If Not TypeOf cn Is ADODB.Connection Then Exit Sub
    '�ж�ָ��Ŀ¼�µ�CSV�ļ��Ƿ���ʵ���ڣ�������ǣ��˳�����
    If Dir(FileFullPath) = "" Then Exit Sub
    
    '�ж����ݿ������Ƿ��Ѵ򿪣����û���������
    If cn.State = 0 Then cn.Open
    'ʹ�����ݿ����� cn ��һ�����ݼ� rs
    'rs ��������Դ��ָ������ tblName �е�ȫ�����ݣ��򿪷�ʽ�� adOpenKeyset
    rs.Open "select * from " & tblName, cn, adOpenKeyset
    'ȡ�����ݼ��ֶ���Ŀ
    iFieldCount = rs.Fields.Count
    
    '���¶���asFieldNames�Ĵ�С������������ݱ�����Ϊ test�����ֶ���
    ReDim asFieldNames(iFieldCount - 1) As String
    '�������asFieldNames�д洢���ֶ���һһ��Ӧ����¼���ݱ��ÿһ���ֶ��Ƿ����ַ�������
    ReDim abFieldIsString(iFieldCount - 1) As Boolean
   '�������ݱ�������ֶΣ�
    For iCtr = 0 To iFieldCount - 1
            '��¼���ݱ���ֶ���
            asFieldNames(iCtr) = "[" & rs.Fields(iCtr).Name & "]"
            '��¼���ݱ���ֶ��Ƿ����ַ�������
            abFieldIsString(iCtr) = FieldIsString(rs.Fields(iCtr))
    Next
             
    'ʹ��FreeFile����һ���ļ���
    iFileNum = FreeFile
    '��ָ��·���µ��ļ��������м� D:\test.csv �ļ���
    'ע�⣺����ʹ���ֽڣ�Binary����ʽ���ļ���
    'ԭ���������ȡCSV�ļ����ݵĲ�����ʹ����LOF����ȡ���ļ�����
    '         ���ʹ���ı���ʽ���ļ�����ô�������ļ��д��������ַ���
    '         ���ȼ���ͻ᲻��ȷ����ˣ�����Binary��ʽ���ļ�
    Open FileFullPath For Binary As #iFileNum
    'ʹ��LOF�����ļ����ȣ�����������ڴ�ռ䣬��ʹ sFileContents ����ָ����һ�ռ�
    sFileContents = Space(LOF(iFileNum))
    '���ļ�����ȫ����ȡ�� sFileContents ������
    Get #iFileNum, , sFileContents
    '��ȡ��ɺ󣬹ر��ļ�
    Close #iFileNum
    '��vbCrLf���س����з�����Ϊ���б�־����CSV�ļ������ݷֽ�Ϊһ���������ַ�������ɵ�����
    '�����������ݴ洢�� sTableSplit ������
    sTableSplit = Split(sFileContents, RecordDelimiter)
    '��� sTableSplit �����Ͻ磨��CSV�ļ����������洢�� lRecordCount ��
    lRecordCount = UBound(sTableSplit)
    
    '��ʼ������
    cn.BeginTrans
    ' lRecordCount ��¼��CSV�ļ�������������ͨ��ѭ�����δ���CSV�ļ���ÿһ��
    For lCtr = 1 To lRecordCount - 1
            '�� ,����Ƕ��ţ���CSV�ļ�ͬһ���е����ݻ���Ϊһ���ַ������飬
            '�����������ݴ洢�� sRecordSplit ������
            sRecordSplit = Split(sTableSplit(lCtr), FieldDelimiter)
            'CSV�ļ���һ������ ,����Ƕ��ţ����ֳ������鳤�Ⱥ����ݱ���ֶ������ܲ�һ��
            '��ˣ�ȡ����֮�г�����С����Ϊ����Ļ�׼
            '���磬�ãӣ��ļ���ĳһ������Ϊ����ţ����
            '��һ�����ݻ��ֺ��õ� sRecordSplit ��������Ϊ sRecordSplit = {��,ţ,��}������Ϊ 3
            '�����ݱ��� userId �� userName �����ֶΣ��ֶ���Ϊ 2����ô���ս� iFieldsToImport = 2
            iFieldsToImport = IIf(UBound(sRecordSplit) + 1 < _
                    iFieldCount, UBound(sRecordSplit) + 1, iFieldCount)
            '���棬Ҫ��CSV�ļ���ǰ�е����ݲ��뵽���ݱ��У�Ϊ�˿�ʼ׼��SQL�ı�
            sSQL = "INSERT   INTO   " & tblName & "   ("
            'ʹ��ѭ���� asFieldNames �����д洢���ֶ�������׷�ӵ�SQL����
            For iCtr = 0 To iFieldsToImport - 1
                    sSQL = sSQL & asFieldNames(iCtr)
                    If iCtr < iFieldsToImport - 1 Then sSQL = sSQL & ","
            Next iCtr

            sSQL = sSQL & ")   VALUES   ("
            ' sRecordSplit ��������Ĵ����У���CSV�ļ�ͬһ���е�������ת���ɵ�����
            '���ڣ�����������������ӵ�SQL�ı���VALUES������
            For iCtr = 0 To iFieldsToImport - 1
                    If abFieldIsString(iCtr) Then
                    '���ֵΪ�ı�����Ҫ���ر�����Ҫ�����CSV�ļ��еİ�ǵ����źͰ��˫���ţ�
                             sSQL = sSQL & prepStringForSQL(sRecordSplit(iCtr))
                    Else
                    '��������ı�����ֱ����ӵ�SQL�ı���VALUES������
                            sSQL = sSQL & sRecordSplit(iCtr)
                    End If
                    '����������һ��ֵ����ô����׷�ӵ�SQL�ı���VALUES����ʱ����Ҫ����
                    '�����׷��һ����Ƕ���
                    If iCtr < iFieldsToImport - 1 Then sSQL = sSQL & ","
            Next iCtr
            '������ţ����SQL�ı�
            sSQL = sSQL & ")"
            'ִ�в������
            cn.Execute sSQL
    '�����������CSV�ļ�����һ�����ݣ�
    Next lCtr
    '�ύ�����������������в����еĲ�����ʽ�ύ��
    cn.CommitTrans
    '�ر����ݼ�
    rs.Close
    '����Ϊ�գ��ͷ���Դ
    Set rs = Nothing

End Sub
'��������--------------------------------------------------

'������ʼ--------------------------------------------------
'�ж����ݱ���ֶ��Ƿ����ַ�������
Private Function FieldIsString(FieldObject As ADODB.Field) _
        As Boolean
           
    Select Case FieldObject.Type
        '������ݱ���ֶ������������е�һ��������Ϊ���ֶ����ַ�������
        Case adBSTR, adChar, adVarChar, adWChar, adVarWChar, _
                     adLongVarChar, adLongVarWChar
                     FieldIsString = True
        '���򣬲����ַ�������
        Case Else
                     FieldIsString = False
    End Select
                   
End Function
'��������--------------------------------------------------

'������ʼ--------------------------------------------------
'���ܣ�������CSV�ļ����ı�����Ԥ����
'        ���������Excel�ļ���������ı��к��а��˫���ţ����ձ�������������ģ�
'        ��������CSV�ĸ�ʽ�����ļ�ʱ��Excel��������������ı��ж������һЩ���˫���ţ�
'        ��ִ��INSERT��������Щ����İ��˫����Ҳ�ᱻ���뵽���ݱ��У��������Ľ����������
'        ��Ҫ�ģ����磬Excel������� ��""�� �ᱻ�Զ�ת���� "��"""��" �ַ������˱�׷���˰��˫���ţ�
'        �м��һ�����˫����Ҳ��Ϊ������������������Ҫ�������ݱ���� ��""�� ������ "��"""��"
'        ���ԣ�������Ԥ����ȥ������İ��˫���ţ�Ȼ���ټ��뵽SQL�����

'        ���⣬���������Excel�ļ���������ı��к��а�ǵ����ţ���ȻExcel�������ر���
'        ���ǣ�����SQL�﷨�й涨ʹ�õ������������ַ��������ԣ����INSERT����VALUES������
'        ���е����ţ��ͻ�������⣬���磺������Excel������ ��'�������浽CSV�ļ�����Ϊ��'��
'        �ڰ�����INSERT����оͻ��Ϊ VALUES(����������������'��'��') �Ӷ�����SQL�﷨����
'
Private Function prepStringForSQL(ByVal sValue As String) _
      As String

    Dim sAns As String
    
    sAns = sValue
    
    'ȥ���ַ������˿��ܱ��Զ�׷�ӵİ��˫����   ���磺 "��"""��" -> ��"""��
    If Len(sAns) <> 0 Then
      If Left(sAns, 1) = Chr(34) Then
        sAns = Right(sAns, Len(sAns) - 1)
        If Len(sAns) <> 0 Then
          If Right(sAns, 1) = Chr(34) Then
              sAns = Left(sAns, Len(sAns) - 1)
          End If
        End If
      End If
    End If
    
    'ȥ���ַ����п��ܱ��Զ�׷�ӵİ��˫����   ���磺 ��"""�� -> ��""��
    sAns = Replace(sAns, Chr(34) & Chr(34), Chr(34))
    '��SQL�ı��У���ǵ����ŵ�ת���ַ�������������ǵ����ţ����ִ�����²�����
    '���ַ����п��ܴ��ڵ�ÿ����ǵ������滻��������ǵ����ţ����磺 ��'�� -> ��''��
    sAns = Replace(sAns, Chr(39), "''")
    '����SQL���ķ��淶�����ַ��������ð�ǵ�������������
    '���磺 VALUES(1,2,3,�й�) -> VALUES(1,2,3,'�й�')
    sAns = "'" & sAns & "'"
    'Ԥ������ϣ����ش���õ��ַ���
    prepStringForSQL = sAns

End Function

