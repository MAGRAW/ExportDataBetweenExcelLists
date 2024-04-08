' Импортируемое пространство имён для процесса сериализации
Imports System.IO
'Imports System.Runtime.Serialization
'Imports System.Runtime.Serialization.Formatters.Binary


Module Module1
    'ReadFile2015(FileName,MassData)
    'WriteFile2015(FileName,MassData)
    'AppendText2015(FileName,MassData)
    'ReplaceText2015(FileName,NumStr,NewText)
    'DeleteText2015(FileName,NumStr)

    'Чтение текстового файла с построчной записью содержимого в массив MassData()
    Public Sub ReadFile2015(ByVal FileName As String, ByRef MassData() As String)
        MassData = IO.File.ReadAllLines(FileName, System.Text.Encoding.UTF8)
        'If MassData.Length = Nothing Then 'проверка пустого файла
        'MassData(0) = "0"
        'End If
    End Sub
    'Запись массива строк MassData() в файл
    Public Sub WriteFile2015(ByVal FileName As String, ByVal MassData() As String)
        IO.File.WriteAllLines(FileName, MassData, System.Text.Encoding.UTF8)
        File.WriteAllLines(FileName, MassData, System.Text.Encoding.UTF8)
        'Using fstream As New FileStream(FileName, FileMode.OpenOrCreate)
        'For i = 0 To MassData.Length - 1
        'Dim array As Byte() = System.Text.Encoding.Default.GetBytes(MassData(i))
        'fstream.Write(array, 0, array.Length)
        'Next
        'End Using
    End Sub
    'Последовательное добавление строк в файл
    Public Sub AppendText2015(ByVal FileName As String, ByVal stroka As String)
        'Dim i As Integer
        'For i = 1 To MassData.Length
        IO.File.AppendAllText(FileName, stroka)
        'Next
    End Sub
    'Замена в файле строки номер NumStr на текст NewText
    Public Sub ReplaceText2015(ByVal FileName As String, ByVal NumStr As Integer, ByVal NewText As String)
        Dim MassData() As String
        ReadFile2015(FileName, MassData)
        MassData(NumStr) = NewText
        IO.File.Delete(FileName)
        WriteFile2015(FileName, MassData)
        'For i = 1 To MassData.Length
        'AppendText2015(FileName, MassData(i - 1))
        'Next
    End Sub
    'Удаление в файле строки номер NumStr
    Public Sub DeleteText2015(ByVal FileName As String, ByVal NumStr As Integer)
        Dim MassData() As String
        Dim newMassData() As String
        Dim i, j As Integer
        j = 0
        ReDim newMassData(MassData.Length - 1)
        ReadFile2015(FileName, MassData)
        IO.File.Delete(FileName)
        For i = 1 To MassData.Length
            If i <> NumStr Then
                newMassData(i - 1 - j) = MassData(i - 1)
            Else
                j = 1
            End If
        Next
        WriteFile2015(FileName, newMassData)
    End Sub

    Public Sub n_var()
        Dim MyPath As String
        Dim adres As String

        MyPath = CurDir() 'Возвращает путь к данному приложению
        'Формирование файла исходных данных
        adres = MyPath + "\59_4.dat"
        FileOpen(2, adres, OpenMode.Input)

        FileClose()
    End Sub

    '1)
    'ReplaceLine имя_файла, номер_строки, на_что_заменить
    'Заменяет в файле имя_файла строчку под номером номер_строки на текст, содержащийся в переменной на_что_заменить 
    Public Sub ReplaceLine(ByVal sFileName As String, ByVal lLineNum As Long, ByVal sReplaceWith As String)
        Dim hFile As Integer
        Dim sBuffer As String
        Dim sLines() As String

        hFile = FreeFile()
        FileOpen(hFile, sFileName, OpenMode.Binary)
        sBuffer = Space(LOF(hFile))
        FileGet(hFile, sBuffer, 1)
        sLines = Split(sBuffer, vbNewLine)
        sLines(lLineNum) = sReplaceWith
        sBuffer = Join(sLines, vbNewLine)
        FilePut(hFile, sBuffer, 1)
        FileClose(hFile)
    End Sub

    Public Sub ReplaceStrInFile(ByVal strPathToFile As String, ByVal lngNumOfReStr As String, ByVal strToReplace As String)
        'strToReplace - путь к файлу
        'lngNumOfReStr - номер заменяемой строки в файле, номер первой строки: 1
        'strToReplace - новая строка
        Dim strStrings() As String
        Dim lngFree As Long
        Dim strTmpInp As String

        On Error GoTo erCh
        lngFree = FreeFile()
        'читаем из файла
        'Open strPathToFile For Input As #lngFree
        FileOpen(lngFree, strPathToFile, OpenMode.Input)

        ReDim strStrings(0)
        Do While Not (EOF(lngFree))
            'Line Input #lngFree, strTmpInp
            ReadLine(lngFree, strTmpInp)
            ReDim Preserve strStrings(UBound(strStrings) + 1)
            strStrings(UBound(strStrings)) = strTmpInp
        Loop
        'Close #lngFree
        FileClose(lngFree)
        'заменяем строку
        strStrings(lngNumOfReStr) = strToReplace
        'Переписываем файл
        lngFree = FreeFile()
        'Open strPathToFile For Output As #lngFree
        FileOpen(lngFree, strPathToFile, OpenMode.Output)
        Dim cnt As Long
        For cnt = 1 To UBound(strStrings)
            'Print #lngFree, strStrings(cnt)
            PrintLine(lngFree, strStrings(cnt))
        Next cnt
        'Close #lngFree
        FileClose(lngFree)
        Exit Sub
erCh:
        MsgBox("Ошибка в процедуре ReplaceStrInFile", vbCritical)
    End Sub

    '2) 
    'DeleteLine имя_файла, номер_строки 
    'Удалёет из файла имя_файла строку под номером номер_строки 
    Public Sub DeleteLine(ByVal sFileName As String, ByVal lLineNum As Long)
        'Dim hFile As Long
        Dim hFile As Integer
        Dim sBuffer As String
        Dim sLines() As String
        Dim i, j As Long

        hFile = FreeFile()
        FileOpen(hFile, sFileName, OpenMode.Binary)
        sBuffer = Space(LOF(hFile))
        FileGet(hFile, sBuffer, 1)
        sLines = Split(sBuffer, vbNewLine)
        For i = lLineNum To UBound(sLines) - 1
            sLines(i) = sLines(i + 1)
        Next i
        ReDim Preserve sLines(UBound(sLines) - 1)
        sBuffer = Join(sLines, vbNewLine)
        'FilePut(hFile, sBuffer, 1)' некорректно пишет
        FileClose(hFile)

        IO.File.WriteAllText(sFileName, sBuffer)
        'Microsoft.VisualBasic.FileIO.FileSystem.WriteAllText(sFileName, sBuffer, False)
        'Microsoft.VisualBasic.FileIO.FileSystem.WriteAllText(sFileName, sBuffer, 1)
    End Sub

    '3) 
    'LineText = ReadLine(имя_файла, номер_строки)
    'Переменная LineText будет содержать строчку под номером номер_строки из файла имя_файла 
    Public Function ReadLine(ByVal sFileName As String, ByVal lLineNum As Long) As String
        'Dim hFile As Long

        Dim hFile As Integer
        Dim sBuffer As String
        Dim sLines() As String

        On Error Resume Next
        hFile = FreeFile()
        'Open sFileName For Binary As #hFile
        FileOpen(hFile, sFileName, OpenMode.Binary)
        sBuffer = Space(LOF(hFile))
        'Get #hFile, 1, sBuffer
        FileGet(hFile, sBuffer, 1)
        sLines = Split(sBuffer, vbNewLine)
        'Close #hFile
        FileClose(hFile)
        ReadLine = sLines(lLineNum - 1)
    End Function

    '4) 
    'AddLine имя_файла, номер_строки, что_добавить
    'Добавляет в файл имя_файла вместо строчки с номером номер_строки текст из переменной что_добавить
    Public Sub AddLine(ByVal sFileName As String, ByVal lLineNum As Long, ByVal sTextToAdd As String)
        'Dim hFile As Long
        Dim hFile As Integer
        Dim sBuffer As String
        Dim sLines() As String

        hFile = FreeFile()
        'Open sFileName For Binary As #hFile
        FileOpen(hFile, sFileName, OpenMode.Binary)
        sBuffer = Space(LOF(hFile))
        'Get #hFile, 1, sBuffer
        FileGet(hFile, sBuffer, 1)
        sLines = Split(sBuffer, vbNewLine)
        ReDim Preserve sLines(UBound(sLines) + 1)
        For i = UBound(sLines) To lLineNum + 1
            sLines(i) = sLines(i - 1)
        Next i
        sLines(lLineNum) = sTextToAdd

        sBuffer = Join(sLines, vbNewLine)
        'Put #hFile, 1, sBuffer
        FilePut(hFile, sBuffer, 1)
        'Close #hFile
        FileClose(hFile)
    End Sub

    Public Function Chislo_Strok(ByVal sFileName As String) As Integer
        Dim stroka As String
        Dim hFile As Object

        hFile = FreeFile()
        FileOpen(hFile, sFileName, OpenMode.Input)
        Chislo_Strok = 0
        Do Until EOF(1)
            Chislo_Strok = Chislo_Strok + 1
            stroka = LineInput(hFile)
        Loop
        FileClose(hFile)
    End Function


    Public Sub Save_IN(ByVal Z() As Double,
                         ByVal R() As Double,
                         ByVal b() As Double,
                         ByVal S() As Double,
                         ByVal a0() As Double,
                         ByVal a1() As Double,
                         ByVal a1ef() As Double,
                         ByVal Cmax() As Double,
                         ByVal w1() As Double,
                         ByVal w2() As Double,
                         ByVal r1() As Double,
                         ByVal r2() As Double,
                         ByVal tX() As Double,
                         ByVal Sprof() As Double,
                         ByVal g1() As Double,
                         ByVal g2() As Double,
                         ByVal tg() As Double,
                         ByVal k1() As Double,
                         ByVal k2() As Double, ByRef i As Integer)

        Dim CellVal As String
        Dim MyPath As String
        Dim adres As String
        Dim str0, str1, str2, str3, str4, str5, str6, str7, str8, str9 As String

        MyPath = CurDir() 'Возвращает путь к данному приложению
        'MyPath = "D:\Работа\Прочее\Профилирование\Rabota\Программа профилирования лопаток\"
        'Формирование файла исходных данных
        adres = MyPath + "\IN.DAT"
        FileOpen(2, adres, OpenMode.Output)
        str0 = "0" + " " + "0" + " " + "0" + " " + "1" + " " + "0"
        PrintLine(2, str0)
        str1 = "1" + " " + "2" + " " + "0" + " " + "0"
        PrintLine(2, str1)
        str2 = "0" + " " + "0" + " " + "-1" + " " + "0" + " " + "0"
        PrintLine(2, str2)
        str3 = Trim(Str(Z(i))) + " " + Trim(Str(R(i)))
        PrintLine(2, str3)
        str4 = Trim(Str(g1(i))) + " " + Trim(Str(tg(i))) + " " + Trim(Str(g2(i)))
        PrintLine(2, str4)
        str5 = Trim(Str(a0(i))) + " " + Trim(Str(a1(i))) + " " + Trim(Str(a1ef(i))) + " " + Trim(Str(w1(i))) + " " + Trim(Str(w2(i)))
        PrintLine(2, str5)
        str6 = Trim(Str(S(i))) + " " + Trim(Str(b(i))) + " " + Trim(Str(Cmax(i))) + " " + Trim(Str(Sprof(i)))
        PrintLine(2, str6)
        str7 = Trim(Str(r1(i))) + " " + Trim(Str(r2(i)))
        PrintLine(2, str7)
        str8 = Trim(Str(k1(i))) + " " + Trim(Str(k2(i)))
        PrintLine(2, str8)
        str9 = Trim(Str(tX(i))) + " " + "0" + " " + "0" + " " + "0" + " " + "0"
        PrintLine(2, str9)
        FileClose()
    End Sub

    Public Sub Read_File_OUT(ByVal MyPath As String, ByVal NewName As String, ByVal i_ As Integer,
            ByRef Nomer() As Integer,
            ByRef SPIN(,) As Double,
            ByRef KORblT(,) As Double,
            ByRef X_sp(,) As Double,
            ByRef Y_sp(,) As Double,
            ByRef X_kor(,) As Double,
            ByRef Y_kor(,) As Double)
        Dim adres As String
        Dim stroka As String
        Dim i, j As Integer
        Dim ident1, ident2 As String

        adres = MyPath + "\" + NewName
        'FileSystem.FileOpen(2, adres, OpenMode.Input)
        'Microsoft.VisualBasic.FileOpen(2, adres, OpenMode.Input)
        FileOpen(2, adres, OpenMode.Input)
        'ident1 = ""
        'Do While ident1 <> "СПИН"
        '  Line Input #2, stroka
        '  ident1 = Trim(Mid(stroka, 1, 10))
        'Loop
        For i = 1 To 45
            'stroka = FileSystem.LineInput(2)
            'stroka = Microsoft.VisualBasic.LineInput(2)
            'Input(2, stroka)
            stroka = LineInput(2)
        Next i

        ident2 = ""
        j = 0
        Do While ident2 <> "R"
            j = j + 1
            'stroka = FileSystem.LineInput(2)
            'stroka = Microsoft.VisualBasic.LineInput(2)
            'Input(2, stroka)
            stroka = LineInput(2)
            ident2 = Trim(Mid(stroka, 1, 35))
            If ident2 <> "R" Then
                Nomer(i_) = j
                'stroka = Trim(Mid(stroka, 1, 11)) : SPIN(j) = System.Convert.ToDouble(stroka)
                'SPIN(j) = Trim(Mid(stroka, 1, 11))
                'stroka = Trim(Mid(stroka, 1, 11)) : SPIN(j) = CType(stroka, Double)
                'stroka = Trim(Mid(stroka, 1, 11)) : SPIN(j) = Val(stroka)
                SPIN(i_, j) = Val(Trim(Mid(stroka, 1, 11)))
                KORblT(i_, j) = Val(Trim(Mid(stroka, 11, 10)))
                X_sp(i_, j) = Val(Trim(Mid(stroka, 28, 13)))
                Y_sp(i_, j) = Val(Trim(Mid(stroka, 41, 13)))
                X_kor(i_, j) = Val(Trim(Mid(stroka, 54, 13)))
                Y_kor(i_, j) = Val(Trim(Mid(stroka, 67, 13)))
            End If
        Loop
        FileClose()
    End Sub

    Public Sub NameSeriesFile(ByRef NameFile1 As String, ByRef NameFile2 As String)

        Dim hFile As Integer
        Dim MyPath As String
        Dim sFileName As String

        'hFile = FreeFile()
        MyPath = CurDir()
        sFileName = MyPath + "\Acontr.ctl"
        'FileOpen(hFile, sFileName, OpenMode.Binary)
        NameFile1 = ReadLine(sFileName, 3)
        NameFile2 = ReadLine(sFileName, 4)
        'FileClose(hFile)

    End Sub










    ' Замена имени ячеек из формата A1 в формат R1C1
    Public Sub Convert_Range_from_cell(ByVal sFileName1 As String, ByRef sell_X As Long, ByRef sell_Y As Long)
        Dim i As Integer
        Dim Len_i As Integer
        Dim stroka As String

        i = 0
        Do While i <= Len(sFileName1)
            i = i + 1
            stroka = Mid(sFileName1, i, 1)
            If Logica(stroka) Then
                Len_i = i
            Else
                Exit Do
            End If
        Loop
        i = 0 : sell_X = 0
        If Len_i > 1 Then
            Do While i <= Len_i
                i = i + 1
                sell_X = sell_X + 26 * Range_from_cell(Mid(sFileName1, i, 1))
            Loop
        ElseIf Len_i = 1 Then
            sell_X = Range_from_cell(Mid(sFileName1, 1, 1))
        End If
        sell_Y = Val(Mid(sFileName1, Len_i + 1, Len(sFileName1) - Len_i + 1))
    End Sub
    Public Function Range_from_cell(ByVal sFileName As String) As Integer
        Dim i As Integer
        Select Case sFileName
            Case "A" : i = 1
            Case "B" : i = 2
            Case "C" : i = 3
            Case "D" : i = 4
            Case "E" : i = 5
            Case "F" : i = 6
            Case "G" : i = 7
            Case "H" : i = 8
            Case "I" : i = 9
            Case "J" : i = 10
            Case "K" : i = 11
            Case "L" : i = 12
            Case "M" : i = 13
            Case "N" : i = 14
            Case "O" : i = 15
            Case "P" : i = 16
            Case "Q" : i = 17
            Case "R" : i = 18
            Case "S" : i = 19
            Case "I" : i = 20
            Case "U" : i = 21
            Case "V" : i = 22
            Case "W" : i = 23
            Case "X" : i = 24
            Case "Y" : i = 25
            Case "Z" : i = 26
        End Select
        Range_from_cell = i
    End Function
    Public Function Logica(ByVal sFileName As String) As Boolean
        Dim i As Integer
        Logica = False
        Select Case sFileName
            Case "A" : i = True
            Case "B" : i = True
            Case "C" : i = True
            Case "D" : i = True
            Case "E" : i = True
            Case "F" : i = True
            Case "G" : i = True
            Case "H" : i = True
            Case "I" : i = True
            Case "J" : i = True
            Case "K" : i = True
            Case "L" : i = True
            Case "M" : i = True
            Case "N" : i = True
            Case "O" : i = True
            Case "P" : i = True
            Case "Q" : i = True
            Case "R" : i = True
            Case "S" : i = True
            Case "I" : i = True
            Case "U" : i = True
            Case "V" : i = True
            Case "W" : i = True
            Case "X" : i = True
            Case "Y" : i = True
            Case "Z" : i = True
        End Select
        Logica = i
    End Function


    ''Определение имени файла
    'Public Function FileName(ByVal adres As String) As String
    'Dim i1, i2 As Integer
    '    i = Len(adres)
    '    Do While stroka <> "."
    '        i = i - 1
    '        stroka = Mid(adres, i, 1)
    '        i2 = i
    '    Loop
    '    Do While stroka <> "\"
    '        i = i - 1
    '        stroka = Mid(adres, i, 1)
    '        i1 = i + 1
    '    Loop
    '    FileName = Mid(adres, i1, i2 - i1)
    'End Function

    'Расшифровка текста из цифрового вила
    Public Function ConvertStroka(ByVal stroka As String) As String
        Dim Stroka1, Stroka2 As String
        Dim j, k As Integer
        Dim fsf, i1, i2, i3 As Integer

        fsf = 0 'счётчик для определения слова/транскрипция/перевод
        Stroka2 = "|" : j = 1
        j = j + 2
        Do While j < Len(stroka)
            k = j
            Do While Stroka1 <> "_" And Stroka1 <> "|"
                Stroka1 = Mid(stroka, j, 1)
                j = j + 1
            Loop
            If Stroka1 = "|" Then
                Stroka2 = Stroka2 + ChrW(Val(Mid(stroka, k, j - k))) + "|"
                Stroka1 = ""
                j = j + 1
                fsf = fsf + 1
                Select Case fsf
                    Case 1 : i1 = Len(Stroka2)
                    Case 2 : i2 = Len(Stroka2)
                    Case 3 : i3 = Len(Stroka2)
                End Select
            Else
                Stroka2 = Stroka2 + ChrW(Val(Mid(stroka, k, j - k)))
                Stroka1 = ""
            End If
        Loop
        ConvertStroka = Stroka2
    End Function

    'Sub SerializeBase(ByVal adres As String)
    'Dim hFile As Integer
    'Dim z0 As New MyClass0
    '    hFile = FreeFile()
    '    FileOpen(hFile, adres, OpenMode.Output)
    '
    '    z0.a = 22
    '    z0.b = False
    'Dim w As New MyClass1
    ''w.a = New Integer() {2, -3}
    ''Write(hFile, stroka1)
    '    z0.c = w
    '    FileClose(hFile)
    '
    '' Сериализация.
    'Dim fs As New FileStream(adres, FileMode.Create)
    'Dim bf As New BinaryFormatter
    '   bf.Serialize(fs, z0)
    '   fs.Close()
    '
    '
    '' Десериализация.
    'Dim z1 As New MyClass0
    '    fs = New FileStream(adres, FileMode.Open)
    '    z1 = Convert.ChangeType(bf.Deserialize(fs), z0.GetType())
    '    fs.Close()
    'End Sub

    Public Function SearchNumberStr(ByVal massiv() As String, ByVal h As Integer, ByVal stroka As String) As Integer
        Dim i As Integer
        For i = 0 To h - 1
            If stroka = massiv(i) Then Exit For
        Next i
        SearchNumberStr = i + 1
    End Function

    Dim stroka As String
    Dim str_IN As String
    Dim str_OUT As String
    Dim str_Prefix As String
    Dim str_Sound As String
    Dim NumberPoz As Integer
    Dim Syllable() As Object
    Dim StartPozSyllable() As Integer
    Dim LenSyllable() As Integer
    Dim Approval As Boolean
    Dim OpenСlosedSyllable() As Object
    Dim NumberSound As Integer
    Dim NumberOfTranscriptionSymbols As String
    Dim AccentSyllable() As Boolean
    Dim PhoneticTransliteration As String

    'вставка символа "_" в строке между графическими слогами
    Public Sub Inserf_Underscore(ByVal str_IN As String, ByVal NumberPoz As Integer, ByRef str_OUT As String)
        Dim stroka_left As String
        Dim stroka_right As String
        stroka_left = Mid(str_IN, 1, Len(str_IN) - NumberPoz + 1)
        stroka_right = Mid(str_IN, NumberPoz, Len(str_IN) - NumberPoz + 1)
        str_OUT = stroka_left + "_" + stroka_right
    End Sub

    'Получение списка всех листов из книги Excel
    'Sub GetAllListExcel(ByVal WbFrom As Workbook, ByRef AllList() As String)
    'Dim I As Integer
    'Dim QuantityList As Integer
    '    WbFrom.Activate()
    '    QuantityList = ActiveWorkbook.Worksheets.Count
    '    ReDim AllList(0 To QuantityList)
    '    AllList(0) = QuantityList
    '    For I = 1 To AllList(0)
    '        AllList(I) = ActiveWorkbook.Worksheets(I).Name
    '    Next I
    'End Sub

    'функция для просмотра содержания массива
    Public Sub ViewArray(ByVal iArray())
    End Sub

    'меняем текущий статус ячеек на противоположный (ОБЪЕДИНЕНИЕ)
    Public Sub ReversMergeCells(ByVal iSheet As Microsoft.Office.Interop.Excel.Worksheet,
                                ByVal stroka As String)

        If iSheet.Range(stroka).MergeCells = True Then
            iSheet.Range(stroka).MergeCells = False
            GoTo ex
        ElseIf iSheet.Range(stroka).MergeCells = False Then
            iSheet.Range(stroka).MergeCells = True
            GoTo ex
        End If
ex:
    End Sub
End Module