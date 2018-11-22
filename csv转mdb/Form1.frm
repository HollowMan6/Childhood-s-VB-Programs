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
   StartUpPosition =   3  '窗口缺省
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
'声明当前VB环境中字符串的比较方式为字节码
Option Compare Binary
'要测试本例的文件操作，就新建一个窗体，在上面添加一个名为"cmdImport"的按钮，
'单击后即执行下面的方法，完成文件读取和写入数据库的操作
'函数开始--------------------------------------------------
'主函数，执行的开始
Private Sub cmdImport_Click()
    Dim cn As New ADODB.Connection
    '指定连接字符串，本例连接到 D:\db1.mdb 数据库
    cn.ConnectionString = _
           "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\db1.mdb"
    '开启连接，准备对数据库操作
    cn.Open
   '定义存储文件名的对象
    Dim fileName As String
    '指定要读取的CSV文件的名称本例为 D:\test.csv
    fileName = "D:\test.csv"
    '执行ImportFile函数，完成以下操作：
    '1 读取 D:\test.csv 文件的内容
    '2 将读取的文件内容写入 D:\db1.mdb 数据库的 test 表中
    ImportFile cn, "test", fileName
    '完成操作后关闭连接
    cn.Close
    '将连接设置为 Nothing ，及时释放其所占用的内存空间
    Set cn = Nothing
End Sub
'函数结束--------------------------------------------------


'函数开始--------------------------------------------------
'参数------------------------------------------------------------------
'1. cn                          使用的数据库联接，本例连接到 D:\db1.mdb
'2. tblName                 使用的数据库表名称，本例为 test
'3. FileFullPath           读取的CSV文件目录（包括文件名），本例为 D:\test.csv
'4. FieldDelimiter        指定CSV文件同一行内容以哪种字符分隔 默认为 ,（半角逗号）
'5. RecordDelimiter    指定CSV文件不同行内容以哪种字符分隔 默认为 vbCrLf（回车换行符）
'参数------------------------------------------------------------------
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
    
    
    '判断传入的cn参数是否是数据库连接对象，如果不是，退出程序
    If Not TypeOf cn Is ADODB.Connection Then Exit Sub
    '判断指定目录下的CSV文件是否真实存在，如果不是，退出程序
    If Dir(FileFullPath) = "" Then Exit Sub
    
    '判断数据库联接是否已打开，如果没有则打开连接
    If cn.State = 0 Then cn.Open
    '使用数据库联接 cn 打开一个数据集 rs
    'rs 的数据来源是指定表名 tblName 中的全部数据，打开方式是 adOpenKeyset
    rs.Open "select * from " & tblName, cn, adOpenKeyset
    '取得数据集字段数目
    iFieldCount = rs.Fields.Count
    
    '重新定义asFieldNames的大小，用来存放数据表（本例为 test）的字段名
    ReDim asFieldNames(iFieldCount - 1) As String
    '与上面的asFieldNames中存储的字段名一一对应，记录数据表的每一个字段是否是字符串类型
    ReDim abFieldIsString(iFieldCount - 1) As Boolean
   '遍历数据表的所有字段：
    For iCtr = 0 To iFieldCount - 1
            '记录数据表的字段名
            asFieldNames(iCtr) = "[" & rs.Fields(iCtr).Name & "]"
            '记录数据表的字段是否是字符串类型
            abFieldIsString(iCtr) = FieldIsString(rs.Fields(iCtr))
    Next
             
    '使用FreeFile分配一个文件号
    iFileNum = FreeFile
    '打开指定路径下的文件（本例中即 D:\test.csv 文件）
    '注意：这里使用字节（Binary）方式打开文件，
    '原因：在下面读取CSV文件内容的操作中使用了LOF函数取得文件长度
    '         如果使用文本方式打开文件，那么，由于文件中存在中文字符，
    '         长度计算就会不正确，因此，采用Binary方式打开文件
    Open FileFullPath For Binary As #iFileNum
    '使用LOF计算文件长度，分配所需的内存空间，并使 sFileContents 变量指向这一空间
    sFileContents = Space(LOF(iFileNum))
    '将文件内容全部读取到 sFileContents 变量中
    Get #iFileNum, , sFileContents
    '读取完成后，关闭文件
    Close #iFileNum
    '以vbCrLf（回车换行符）作为换行标志，将CSV文件的内容分解为一个由若干字符串行组成的数组
    '并把数组内容存储到 sTableSplit 数组中
    sTableSplit = Split(sFileContents, RecordDelimiter)
    '获得 sTableSplit 数组上界（即CSV文件行数），存储到 lRecordCount 中
    lRecordCount = UBound(sTableSplit)
    
    '开始事务处理
    cn.BeginTrans
    ' lRecordCount 记录了CSV文件的行数，下面通过循环依次处理CSV文件的每一行
    For lCtr = 1 To lRecordCount - 1
            '以 ,（半角逗号）将CSV文件同一行中的内容划分为一个字符串数组，
            '并把数组内容存储到 sRecordSplit 数组中
            sRecordSplit = Split(sTableSplit(lCtr), FieldDelimiter)
            'CSV文件的一行中以 ,（半角逗号）划分出的数组长度和数据表的字段数可能不一致
            '因此，取两者之中长度最小的作为处理的基准
            '例如，ＣＳＶ文件的某一行数据为　马，牛，羊
            '这一行数据划分后获得的 sRecordSplit 数组内容为 sRecordSplit = {马,牛,羊}，长度为 3
            '而数据表有 userId 和 userName 两个字段，字段数为 2，那么最终将 iFieldsToImport = 2
            iFieldsToImport = IIf(UBound(sRecordSplit) + 1 < _
                    iFieldCount, UBound(sRecordSplit) + 1, iFieldCount)
            '下面，要把CSV文件当前行的内容插入到数据表中，为此开始准备SQL文本
            sSQL = "INSERT   INTO   " & tblName & "   ("
            '使用循环将 asFieldNames 数组中存储的字段名依次追加到SQL文中
            For iCtr = 0 To iFieldsToImport - 1
                    sSQL = sSQL & asFieldNames(iCtr)
                    If iCtr < iFieldsToImport - 1 Then sSQL = sSQL & ","
            Next iCtr

            sSQL = sSQL & ")   VALUES   ("
            ' sRecordSplit 是在上面的处理中，由CSV文件同一行中的内容所转化成的数组
            '现在，将这个数组的内容添加到SQL文本的VALUES集合中
            For iCtr = 0 To iFieldsToImport - 1
                    If abFieldIsString(iCtr) Then
                    '如果值为文本，需要做特别处理（主要是针对CSV文件中的半角单引号和半角双引号）
                             sSQL = sSQL & prepStringForSQL(sRecordSplit(iCtr))
                    Else
                    '如果不是文本，就直接添加到SQL文本的VALUES集合中
                            sSQL = sSQL & sRecordSplit(iCtr)
                    End If
                    '如果不是最后一个值，那么，在追加到SQL文本的VALUES集合时，不要忘记
                    '在其后追加一个半角逗号
                    If iCtr < iFieldsToImport - 1 Then sSQL = sSQL & ","
            Next iCtr
            '添加括号，完成SQL文本
            sSQL = sSQL & ")"
            '执行插入操作
            cn.Execute sSQL
    '下面继续处理CSV文件的下一行数据：
    Next lCtr
    '提交事务处理（即将上面所有插入行的操作正式提交）
    cn.CommitTrans
    '关闭数据集
    rs.Close
    '设置为空，释放资源
    Set rs = Nothing

End Sub
'函数结束--------------------------------------------------

'函数开始--------------------------------------------------
'判断数据表的字段是否是字符串类型
Private Function FieldIsString(FieldObject As ADODB.Field) _
        As Boolean
           
    Select Case FieldObject.Type
        '如果数据表的字段是如下类型中的一个，则认为该字段是字符串类型
        Case adBSTR, adChar, adVarChar, adWChar, adVarWChar, _
                     adLongVarChar, adLongVarWChar
                     FieldIsString = True
        '否则，不是字符串类型
        Case Else
                     FieldIsString = False
    End Select
                   
End Function
'函数结束--------------------------------------------------

'函数开始--------------------------------------------------
'功能：对来自CSV文件的文本进行预处理
'        如果我们在Excel文件中输入的文本中含有半角双引号，按照本文最上面分析的，
'        当我们以CSV的格式保存文件时，Excel会在我们输入的文本中额外插入一些半角双引号，
'        在执行INSERT操作后，这些多余的半角双引号也会被插入到数据表中，而这样的结果不是我们
'        想要的，例如，Excel中输入的 李""四 会被自动转换成 "李"""四" 字符串两端被追加了半角双引号，
'        中间的一个半角双引号也变为了两个，我们真正想要插入数据表的是 李""四 而不是 "李"""四"
'        所以，必须做预处理去掉多余的半角双引号，然后再加入到SQL语句中

'        另外，如果我们在Excel文件中输入的文本中含有半角单引号，虽然Excel不会做特别处理，
'        但是，由于SQL语法中规定使用单引号来引用字符串，所以，如果INSERT语句的VALUES集合中
'        含有单引号，就会出现问题，例如：我们在Excel中输入 张'三，保存到CSV文件中仍为张'三
'        在包含到INSERT语句中就会成为 VALUES(．．．，．．．，'张'三') 从而导致SQL语法错误
'
Private Function prepStringForSQL(ByVal sValue As String) _
      As String

    Dim sAns As String
    
    sAns = sValue
    
    '去掉字符串两端可能被自动追加的半角双引号   例如： "李"""四" -> 李"""四
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
    
    '去掉字符串中可能被自动追加的半角双引号   例如： 李"""四 -> 李""四
    sAns = Replace(sAns, Chr(34) & Chr(34), Chr(34))
    '在SQL文本中，半角单引号的转义字符是两个连续半角单引号，因此执行如下操作，
    '将字符串中可能存在的每个半角单引号替换成两个半角单引号，例如： 张'三 -> 张''三
    sAns = Replace(sAns, Chr(39), "''")
    '按照SQL的文法规范，把字符串两端用半角单引号引起来，
    '例如： VALUES(1,2,3,中国) -> VALUES(1,2,3,'中国')
    sAns = "'" & sAns & "'"
    '预处理完毕，返回处理好的字符串
    prepStringForSQL = sAns

End Function

