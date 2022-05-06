Attribute VB_Name = "Module1"
'************************************************************************************
' ＜サンプルデータ作成の自動生成＞
'
' 変更履歴
' №   日付       変更者　変更内容
'-----------------------------------------------------------------------------------
' v1.0 22022/04/27 梶原    新規作成
'************************************************************************************
Option Explicit
Const StartIndex As Integer = 10 'StartIndex

Const FieldColumn As Integer = 5 'Field Column Number
Const VarinantColumn As Integer = 6 'Variant Column Number
Const MandatoryColumn As Integer = 7 'Mandatory Column Number
Const LengthColumn As Integer = 8 'Length Column Number
Const TypeColumn As Integer = 9 'Type Column Number

Dim ProcessSheet As Worksheet 'ActiveSheet
Dim HeaderFileName As String 'HeaderFileName
Dim DetailFileName As String 'DetailFileName
Dim LastRecord As Integer 'LastRecord

Dim ColumnLength_Header As Integer 'Column Length Header
Dim ColumnLength_Detail As Integer 'Column Length Detail
Dim FileObject As Object 'FileSystemObject
Dim APIName As String    'APIName
Dim Path As String       'Path for the Output File

'|for json
Const jsonObjectName As String = "filed_id" 'json ObjectName
Dim jsonArray As Variant 'Insert json Statement by 1column

'|for XML
Const ParentTAG As String = "REQ" 'XML BranchName
Const HeaderTAG As String = "IN_PARM" 'XML BranchName
Const DetailTAG As String = "IN_DETAIL" 'XML BranchName

Dim XMLFileName As String 'XMLFileName

Const Indent As String = "      " '　IndentBlank

'Create json Data
Sub CreateJsonData()
         
   '|*Inital Process
    InitialProcess ("")
   
   '|*CreateSampleData
   CreateSampleData_json
   
    '|*ResultMessage & OutPut jsonData
     If VarType(jsonArray) = 0 Then
        MsgBox "出力はありません"
     Else
        OutputJson (jsonArray)
        MsgBox "シート「json出力」にデータを出力しました。"
     End If
    
End Sub

'Create TEXT Data
Sub CreateTextData()
    
   '|*Inital Process
    InitialProcess ("Flat")
   
   '|*CreateSampleData
    CreateSampleData_Flat
   
   '|*ResultMessage
    If HeaderFileName = Null Then
         MsgBox "ヘッダーファイル出力時にエラーが発生しました"
         Exit Sub
     End If
     If DetailFileName = Null Then
         MsgBox "詳細ファイル出力時にエラーが発生しました"
        Exit Sub
     End If
     If HeaderFileName = "" And DetailFileName = "" Then
         MsgBox "出力ファイルはありません"
        Exit Sub
     End If
   
    Dim Message As New StringBuilder
    Message.Append ("以下のファイルを作成しました。" & vbCrLf & vbCrLf)
    If HeaderFileName <> "" Then
        Message.Append ("ヘッダーファイル：" & HeaderFileName)
    End If
    If DetailFileName <> "" Then
        Message.Append (vbCrLf & vbCrLf & "詳細ファイル：" & DetailFileName)
    End If
    
    MsgBox Message.ToString
    
End Sub

'Create XML Data
Sub CreateXMLData()
   
   '|*Inital Process
    InitialProcess ("XML")
     
   '|*CreateSampleData
    CreateSampleData_XML
   
   '|*Result Message
    If XMLFileName = Null Then
         MsgBox "ヘッダーファイル出力時にエラーが発生しました"
         Exit Sub
     End If
     If XMLFileName = "" Then
         MsgBox "出力ファイルはありません"
        Exit Sub
     Else
        Dim Message As New StringBuilder
        Message.Append ("以下のファイルを作成しました。" & vbCrLf & vbCrLf)
        If XMLFileName <> "" Then
            Message.Append ("XMLファイル：" & XMLFileName)
        End If
    End If
    
    MsgBox Message.ToString

End Sub

'Initial Process
Sub InitialProcess(Format As String)

   Set ProcessSheet = ActiveSheet 'SET ActiveSheet
   APIName = GetAPIName  'GET APIName
   Set FileObject = SetFileObject 'SET FileSystemObject
   
   'Set Folder Path
   If Format = "" Then
        Path = ""
   Else
        Path = ThisWorkbook.Path & "\" & APIName
   End If
   
   'Create File Directory
   If Path <> "" Then
        Call CreateDirectory(Path, Format)     'Create Path for the Output File
   End If
   Path = Path & "\" & Format
   
   'ColumnSize for Header & Detail
    Dim col As Integer
    Dim row As Integer
   col = 3
   row = 3
   ColumnLength_Header = 0
   ColumnLength_Detail = 0
   Do While True
        'Check Length
        If ProcessSheet.Cells(col, row).Value <> "" Then
            'Header Information
            If ProcessSheet.Cells(3, row).Interior.ColorIndex = xlNone Then
                ColumnLength_Header = ColumnLength_Header + 1
            Else
            'Detail Information
                ColumnLength_Detail = ColumnLength_Detail + 1
            End If
        Else
            Exit Do
        End If
        row = row + 1
    Loop
    LastRecord = ProcessSheet.Range("B1048576").End(xlUp).row
End Sub

'Create Directory for OutPut
Sub CreateDirectory(Path As String, Format As String)
    
    If Dir(Path, vbDirectory) = "" Then
        MkDir Path      'Create Pre-Directory
    End If
    
    If Dir(Path & "\" & Format, vbDirectory) = "" Then
        MkDir Path & "\" & Format      'Create Directory
    End If
    
End Sub

'Get API Name
Function GetAPIName() As String
    GetAPIName = ProcessSheet.Range("D2").Value
End Function

'SET File System Object
Function SetFileObject() As Object
    Set SetFileObject = CreateObject("Scripting.FileSystemObject")
End Function
'CreateSampleData for json
Sub CreateSampleData_json()
    
   Dim tmpStr As New StringBuilder  'Record

   Dim BranchNUM As Integer   'BranchNumber(for SalesInput)
   Dim OutputFlg As Boolean: OutputFlg = False 'Flg for output json data  or not
   
   Dim tmp As Integer
   Dim col As Integer
   Dim row As Integer
   col = StartIndex
   row = 2
   
   Do While True
        'Process Target
        If ProcessSheet.Cells(col, row).Value = 1 Then
            row = row + 1
            'Check First Row for Record
            If ProcessSheet.Cells(col, row).Value <> "" Then
                
                'Initial json Format
                OutputFlg = True
                tmp = 0
                ReDim jsonArray(tmp)
                jsonArray(tmp) = "{"
                
                tmp = UBound(jsonArray) + 1
                ReDim Preserve jsonArray(tmp)
                jsonArray(tmp) = Indent & """" & jsonObjectName & """" & ": {"
                
                'Set Each Header Column
                Dim i As Integer
                 For i = 1 To ColumnLength_Header
                    tmpStr.Append (Indent)
                    tmpStr.Append (Indent)
                    tmpStr.Append ("""")
                    tmpStr.Append (ProcessSheet.Cells(FieldColumn, row).Value)
                    tmpStr.Append ("""")
                    tmpStr.Append (" : ")
                    If UCase(ProcessSheet.Cells(TypeColumn, row).Value) = "NUMBER" Or UCase(ProcessSheet.Cells(LengthColumn, row).Value) = "INTEGER" Or UCase(ProcessSheet.Cells(LengthColumn, row).Value) = "LONG" _
                    Or UCase(ProcessSheet.Cells(TypeColumn, row).Value) = "NULL" Then
                        tmpStr.Append (LCase(ProcessSheet.Cells(col, row).Value))
                    Else
                        tmpStr.Append ("""")
                        tmpStr.Append (ProcessSheet.Cells(col, row).Value)
                        tmpStr.Append ("""")
                    End If
                    If i <> ColumnLength_Header Then
                         tmpStr.Append (",")
                    ElseIf ColumnLength_Detail > 0 Then
                         tmpStr.Append (",")
                    End If
                    
                    tmp = UBound(jsonArray) + 1
                    ReDim Preserve jsonArray(tmp)
                    jsonArray(tmp) = tmpStr.ToString
                    
                    Set tmpStr = New StringBuilder
                    row = row + 1
                 Next
                       
            'Process Detail
                If ColumnLength_Detail = 0 Then
                     Exit Do
                Else
                    col = col + 1
                End If
                    
                 Do While True
                    If ProcessSheet.Cells(col, row + 1).Value <> "" Then
                        
                        'Initial Detail Process
                       If BranchNUM = 0 Then
                            
                            'Prefix for Array
                            tmpStr.Append (Indent)
                            tmpStr.Append (Indent)
                            tmpStr.Append ("""")
                            tmpStr.Append (ProcessSheet.Cells(FieldColumn, row).Value)
                            tmpStr.Append ("""")
                            tmpStr.Append (" : ")
                            tmpStr.Append ("[")
                            tmp = UBound(jsonArray) + 1
                            ReDim Preserve jsonArray(tmp)
                            jsonArray(tmp) = tmpStr.ToString
                            Set tmpStr = New StringBuilder
                                   
                         
                       End If
                        BranchNUM = BranchNUM + 1
                        row = row + 1
                        tmpStr.Append (Indent)
                        tmpStr.Append (Indent)
                        tmpStr.Append (Indent)
                        tmpStr.Append ("{")
                        tmp = UBound(jsonArray) + 1
                        ReDim Preserve jsonArray(tmp)
                        jsonArray(tmp) = tmpStr.ToString
                        Set tmpStr = New StringBuilder

                      'Set EaceColumn(Detail)
                       For i = 1 To ColumnLength_Detail - 1
                            tmpStr.Append (Indent)
                            tmpStr.Append (Indent)
                            tmpStr.Append (Indent)
                            tmpStr.Append (Indent)
                            tmpStr.Append ("""")
                            tmpStr.Append (ProcessSheet.Cells(FieldColumn, row).Value)
                            tmpStr.Append ("""")
                            tmpStr.Append (" : ")
                            If UCase(ProcessSheet.Cells(TypeColumn, row).Value) = "NUMBER" Or UCase(ProcessSheet.Cells(LengthColumn, row).Value) = "INTEGER" Or UCase(ProcessSheet.Cells(LengthColumn, row).Value) = "LONG" _
                              Or UCase(ProcessSheet.Cells(TypeColumn, row).Value) = "NULL" Then
                                  tmpStr.Append (LCase(ProcessSheet.Cells(col, row).Value))
                              Else
                                  tmpStr.Append ("""")
                                  tmpStr.Append (ProcessSheet.Cells(col, row).Value)
                                  tmpStr.Append ("""")
                              End If
                              If i <> ColumnLength_Detail - 1 Then
                                 tmpStr.Append (",")
                            End If
                            
                            tmp = UBound(jsonArray) + 1
                            ReDim Preserve jsonArray(tmp)
                            jsonArray(tmp) = tmpStr.ToString
                            
                            Set tmpStr = New StringBuilder
                            row = row + 1
                        Next
                       'Suffix for 1 record
                        tmpStr.Append (Indent)
                        tmpStr.Append (Indent)
                        tmpStr.Append (Indent)
                        tmpStr.Append ("}")
                        
                        'Check Whether next Record Exist or not
                        If ProcessSheet.Cells(col + 1, row - ColumnLength_Detail + 1).Value <> "" _
                        Or ProcessSheet.Cells(col + 1, row - ColumnLength_Detail + 2).Value <> "" Then
                            tmpStr.Append (",")
                        End If
                        
                         tmp = UBound(jsonArray) + 1
                        ReDim Preserve jsonArray(tmp)
                        jsonArray(tmp) = tmpStr.ToString
                        Set tmpStr = New StringBuilder
                        
                    Else
                        Exit Do
                    End If
                    col = col + 1
                    row = row - ColumnLength_Detail
                 Loop
                 
                 'Suffix for Array
                  tmpStr.Append (Indent)
                  tmpStr.Append (Indent)
                  tmpStr.Append ("]")
                  
                  tmp = UBound(jsonArray) + 1
                  ReDim Preserve jsonArray(tmp)
                  jsonArray(tmp) = tmpStr.ToString
                  Set tmpStr = New StringBuilder
                  
                  Exit Do
            End If
        Else
            If LastRecord < col Then
                Exit Do
            Else
                col = col + 1
            End If
        End If
   Loop
   
   If OutputFlg Then
      'Suffix for filed_id
        tmpStr.Append (Indent)
        tmpStr.Append ("}")
        
         tmp = UBound(jsonArray) + 1
        ReDim Preserve jsonArray(tmp)
        jsonArray(tmp) = tmpStr.ToString
        Set tmpStr = New StringBuilder
        
        'Suffix for json
        tmpStr.Append ("}")
        
         tmp = UBound(jsonArray) + 1
        ReDim Preserve jsonArray(tmp)
        jsonArray(tmp) = tmpStr.ToString
        Set tmpStr = New StringBuilder
    End If
End Sub
'CreateSampleData for XML
Sub CreateSampleData_XML()
    
   Dim tmpStr As New StringBuilder  'Record

   Dim BranchNUM As Integer   'BranchNumber(for SalesInput)
   Dim OutputFlg As Boolean: OutputFlg = False 'Flg for output json data  or not
   
   Dim retXMLTAG 'Return XML TAG Name
   Dim col As Integer
   Dim row As Integer
   col = StartIndex
   row = 2
   
   If Not CheckXML Then
        MsgBox "XML変換対象外のAPIです。"
        Exit Sub
   End If
   
   Do While True
        'Process Target
        If ProcessSheet.Cells(col, row).Value = 1 Then
            row = row + 1
            'Check First Row for Record
            If ProcessSheet.Cells(col, row).Value <> "" Then

                'Create XML File
                
                'Branch Parent TAG (Prefix)
                OutputFlg = True
                XMLFileName = CreateFile("", Path, "xml", "<" & ParentTAG & ">")
                 If XMLFileName = Null Then
                    Exit Sub
                End If
                
                'Header Parent TAG (Prefix)
                tmpStr.Append (Indent)
                tmpStr.Append ("<")
                tmpStr.Append (HeaderTAG)
                tmpStr.Append (">")
                Call AddFile(XMLFileName, tmpStr.ToString)
                Set tmpStr = New StringBuilder

                'Set Each Header Column
                Dim i As Integer
                 For i = 1 To ColumnLength_Header
                    tmpStr.Append (Indent)
                    tmpStr.Append (Indent)
                    
                    retXMLTAG = GetXMLTAG(ProcessSheet.Cells(FieldColumn, row).Value)
                    If retXMLTAG = Null Then
                         MsgBox "XMLのTAG取得に失敗しました" _
                         & vbCrLf & vbCrLf & "FieldId : " & ProcessSheet.Cells(FieldColumn, row).Value
                         Exit Sub
                    End If
                    tmpStr.Append ("<")
                    tmpStr.Append (retXMLTAG)
                    tmpStr.Append (">")
                    tmpStr.Append (ProcessSheet.Cells(col, row).Value)
                    tmpStr.Append ("</")
                    tmpStr.Append (retXMLTAG)
                    tmpStr.Append (">")
                    
                    Call AddFile(XMLFileName, tmpStr.ToString)
                    Set tmpStr = New StringBuilder
                    row = row + 1
                 Next
                
                'Header Parent TAG (suffix)
                tmpStr.Append (Indent)
                tmpStr.Append ("</")
                tmpStr.Append (HeaderTAG)
                tmpStr.Append (">")
                Call AddFile(XMLFileName, tmpStr.ToString)
                Set tmpStr = New StringBuilder
                       
            'Process Detail
                If ColumnLength_Detail = 0 Then
                     Exit Do
                Else
                    col = col + 1
                End If
                    
                 Do While True
                    If ProcessSheet.Cells(col, row + 1).Value <> "" Then
                        
                        'Detail Parent TAG (Prefix)
                         tmpStr.Append (Indent)
                         tmpStr.Append ("<")
                         tmpStr.Append (DetailTAG)
                         tmpStr.Append (">")
                         Call AddFile(XMLFileName, tmpStr.ToString)
                         Set tmpStr = New StringBuilder
                         
                         
                         BranchNUM = BranchNUM + 1
                         row = row + 1
                         tmpStr.Append (Indent)
                         tmpStr.Append (Indent)
                         tmpStr.Append ("<")
                         tmpStr.Append ("BRCH_NUM")
                         tmpStr.Append (">")
                         tmpStr.Append (Format(BranchNUM, "00000"))
                         tmpStr.Append ("</")
                         tmpStr.Append ("BRCH_NUM")
                         tmpStr.Append (">")
                      
                         Call AddFile(XMLFileName, tmpStr.ToString)
                         Set tmpStr = New StringBuilder
                         
                      'Set EaceColumn(Detail)
                       For i = 1 To ColumnLength_Detail - 1
                            tmpStr.Append (Indent)
                            tmpStr.Append (Indent)
                            retXMLTAG = GetXMLTAG(ProcessSheet.Cells(FieldColumn, row).Value)
                            If retXMLTAG = Null Then
                                 MsgBox "XMLのTAG取得に失敗しました" _
                                 & vbCrLf & vbCrLf & "FieldId : " & ProcessSheet.Cells(FieldColumn, row).Value
                                 Exit Sub
                            End If
                            tmpStr.Append ("<")
                            tmpStr.Append (retXMLTAG)
                            tmpStr.Append (">")
                            tmpStr.Append (ProcessSheet.Cells(col, row).Value)
                            tmpStr.Append ("</")
                            tmpStr.Append (retXMLTAG)
                            tmpStr.Append (">")
                            
                            Call AddFile(XMLFileName, tmpStr.ToString)
                            Set tmpStr = New StringBuilder
                            row = row + 1
                        Next
                       
                       'Detail Parent TAG (Suffix)
                         tmpStr.Append (Indent)
                         tmpStr.Append ("</")
                         tmpStr.Append (DetailTAG)
                         tmpStr.Append (">")
                         Call AddFile(XMLFileName, tmpStr.ToString)
                         Set tmpStr = New StringBuilder
                    Else
                        Exit Do
                    End If
                    col = col + 1
                    row = row - ColumnLength_Detail
                 Loop
                  
                  Exit Do
            End If
        Else
            If LastRecord < col Then
                Exit Do
            Else
                col = col + 1
            End If
        End If
   Loop
   
   If OutputFlg Then
        'Branch Parent TAG (Suffix)
        Call AddFile(XMLFileName, "</" & ParentTAG & ">")
    End If
End Sub
'CreateSampleData for TextFile
Sub CreateSampleData_Flat()
    
   Dim TMPHeader As New StringBuilder  'Header Record
   Dim TMPDetail As New StringBuilder  'Detail Record

   Dim BranchNUM As Integer   'BranchNumber(for SalesInput)
   Dim SEQ As Integer              'SEQ(for BATCH Connection)
   BranchNUM = 0
   SEQ = 0
   
   Dim col As Integer
   Dim row As Integer
   col = StartIndex
   row = 2
   
   'OutPut Header
   HeaderFileName = ""
   DetailFileName = ""
   
   Do While True
        'Process Target
        If ProcessSheet.Cells(col, row).Value = 1 Then
            row = row + 1
            'Check First Row for Record
            If ProcessSheet.Cells(col, row).Value <> "" Then
                    
            'Process Header
                 If SEQ > 0 Then
                    Set TMPHeader = New StringBuilder
                    BranchNUM = 0
                 End If
                'Set SEQ
                SEQ = SEQ + 1
                TMPHeader.Append (Format(SEQ, "00000"))
         
                                  
                'Set EaceColumn(Header)
                Dim i As Integer
                 For i = 1 To ColumnLength_Header
                    TMPHeader.Append (PaddingProcess(row, col))
                    row = row + 1
                 Next
                 'Output Header
                  If SEQ = 1 Then
                    HeaderFileName = CreateFile("H", Path, "txt", TMPHeader.ToString)
                  Else
                    Call AddFile(HeaderFileName, TMPHeader.ToString)
                  End If
                  If HeaderFileName = Null Then
                    Exit Sub
                 End If
                col = col + 1
                       
            'Process Detail
                If ColumnLength_Detail Then
                     row = row + 1
                End If
                    
                 Do While True
                    If ProcessSheet.Cells(col, row).Value <> "" Then
                        If SEQ > 0 Then
                            Set TMPDetail = New StringBuilder
                        End If
                       'Set SEQ & Branch Number
                       TMPDetail.Append (Format(SEQ, "00000"))
                       BranchNUM = BranchNUM + 1
                       TMPDetail.Append (Format(BranchNUM, "00000"))
                       
                      'Set EaceColumn(Detail)
                       For i = 1 To ColumnLength_Detail - 1
                           TMPDetail.Append (PaddingProcess(row, col))
                           row = row + 1
                        Next
                    Else
                        Exit Do
                    End If
                    'OutPut Detail
                    If SEQ = 1 And BranchNUM = 1 Then
                        DetailFileName = CreateFile("D", Path, "txt", TMPDetail.ToString)
                        If DetailFileName = Null Then
                             Exit Sub
                         End If
                    Else
                          Call AddFile(DetailFileName, TMPDetail.ToString)
                    End If
                    
                    col = col + 1
                    row = row - ColumnLength_Detail + 1
                 Loop
                 
                'Set Initial HeaderRow
                row = 2
            End If
        Else
            If LastRecord < col Then
                Exit Do
            Else
                col = col + 1
            End If
        End If
   Loop
     
End Sub
'PaddingProcess for Batch
Function PaddingProcess(row As Integer, col As Integer) As String

        Dim tmpString As New StringBuilder
        Dim j As Integer
        'Byte Length for Value
        Dim StringLength As Integer
        StringLength = LenB(StrConv(ProcessSheet.Cells(col, row).Value, vbFromUnicode))
      
      'IN case Number Type
        If UCase(ProcessSheet.Cells(TypeColumn, row).Value) = "NUMBER" Or UCase(ProcessSheet.Cells(LengthColumn, row).Value) = "INTEGER" Or UCase(ProcessSheet.Cells(LengthColumn, row).Value) = "LONG" Then
            'Process 0 Padding
            For j = 1 To ProcessSheet.Cells(LengthColumn, row).Value - StringLength
                tmpString.Append ("0")
            Next
            tmpString.Append (ProcessSheet.Cells(col, row).Value)
        
        'IN case String Type
        Else
            tmpString.Append (ProcessSheet.Cells(col, row).Value)
            
            'Process Blank Padding
            For j = 1 To ProcessSheet.Cells(LengthColumn, row).Value - StringLength
                tmpString.Append (" ")
            Next
        End If
        PaddingProcess = tmpString.ToString
End Function
'CreateFile
'@Param TypeId: H　→　Header , D　→　Detail
'@Param FilePath:Output FilePath
'@Param FileFormat:"txt","xml","json"
'@Param FileStatement:Target Statement for Writing Process
'@Return : NOT NULL →　Success, NULL →　Failure
Function CreateFile(TypeId As String, FilePath As String, FileFormat As String, Filestr As String) As String
    
    'CreateFile
    Dim SysTime As Date
    SysTime = Now
    
    Dim Filefullname As String
    Filefullname = FilePath & "\" & APIName & "_" & TypeId & "_" _
                          & Format(Year(SysTime), "0000") & Format(Month(SysTime), "00") & Format(Day(SysTime), "00") & Format(Hour(SysTime), "00") & Format(Minute(SysTime), "00") & Format(Second(SysTime), "00") _
                          & "." & FileFormat

On Error GoTo ErrHandl
        Open Filefullname For Output As #1
        Print #1, Filestr
        Close #1
        CreateFile = Filefullname
        Exit Function
ErrHandl:
        CreateFile = Null
        Exit Function

End Function
'AddFile
'@Param FileStatement:Target Statement for Writing Process
Sub AddFile(Filefullname As String, Filestr As String)

    Open Filefullname For Append As #1
    Print #1, Filestr
    Close #1

End Sub
'Output json Data
Sub OutputJson(jsonArray As Variant)
    
    'CreateSheets
    Dim ws As Worksheet
    Dim OutSheetsName As String: OutSheetsName = "json出力"
    Dim isExist As Boolean: isExist = False
     
    'Check Sheets "json出力"
    For Each ws In Worksheets
        If ws.Name = OutSheetsName Then
            isExist = True
            Exit For
        End If
    Next ws
    
    If isExist = False Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = OutSheetsName
    End If
    
    'Clear Cell
    Sheets(OutSheetsName).Cells.Clear
    
    'Output json Data
    Dim i As Integer
    Dim OutPutCol As Integer: OutPutCol = 2
    
    For i = 0 To UBound(jsonArray)
        Sheets(OutSheetsName).Cells(OutPutCol, 2).Value = jsonArray(i)
        OutPutCol = OutPutCol + 1
    Next
End Sub

'Check XML is available
Function CheckXML() As Boolean

    Dim SerchRange As Range   'range for serch
    Dim Output As String 'Output
    Dim ws As Worksheet
    Set ws = Worksheets("参考_XML→json変換")
    
    Set SerchRange = ws.Range("A1:A100001")
    
    On Error GoTo ErrHandl
        Output = WorksheetFunction.VLookup(APIName, SerchRange, 1, False)
        CheckXML = True
        Exit Function
        
ErrHandl:
        CheckXML = False
        Exit Function
    
End Function

'GET XML TAG is available
Function GetXMLTAG(Value As String) As String

    Dim SerchRange As Range   'range for serch
    Dim Output As String 'Output
    Dim ws As Worksheet
    Set ws = Worksheets("参考_XML→json変換")
    
    Set SerchRange = ws.Range("C1:D100001")
    
    On Error GoTo ErrHandl
        GetXMLTAG = WorksheetFunction.VLookup(Value, SerchRange, 2, False)
        Exit Function

ErrHandl:
        GetXMLTAG = Null
        Exit Function
    
End Function

