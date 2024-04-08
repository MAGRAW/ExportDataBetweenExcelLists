Imports Excel = Microsoft.Office.Interop.Excel
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports System.IO

Public Class Form1
    Inherits System.Windows.Forms.Form

    Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
    Private Const SYNCHRONIZE As Long = &H100000
    Private Const INFINITE As Long = &HFFFFFFFF
    Public Gindex As Long

    Private buttonPanel As New Panel ' ХЗ что это
    Private WithEvents songsDataGridView As New DataGridView ' НОВЫЙ элемент DataGridView
    Private WithEvents addNewRowButton As New System.Windows.Forms.Button 'скорее всего не нужно
    Private WithEvents deleteRowButton As New System.Windows.Forms.Button 'скорее всего не нужно

    'переменные хранящие размеры крупных элементов управления
    Public Form1_Width As Object = 600
    Public Form1_Height As Object = 887
    Public GroupBox4_Width As Object = 560
    Public GroupBox4_Height As Object = 80
    Public GroupBox4_Left As Object = 12
    Public GroupBox4_Top As Object = 36
    Public GroupBox2_Width As Object = 560
    Public GroupBox2_Height As Object = 108
    Public GroupBox2_Left As Object = 12
    Public GroupBox2_Top As Object = 132
    Public GroupBox1_Width As Object = 560
    Public GroupBox1_Height As Object = 234
    Public GroupBox1_Left As Object = 12
    Public GroupBox1_Top As Object = 256
    Public songsDataGridView_Width As Object = 560
    Public songsDataGridView_Height As Object = 311
    Public songsDataGridView_Left As Object = 12
    Public songsDataGridView_Top As Object = 506
    Public RichTextBox3_Width As Object
    Public RichTextBox3_Height As Object
    Public RichTextBox3_Left As Object
    Public RichTextBox3_Top As Object
    Public RichTextBox4_Width As Object = 560
    Public RichTextBox4_Height As Object = 80
    Public RichTextBox4_Left As Object = 12
    Public RichTextBox4_Top As Object = 36
    Public RichTextBox5_Width As Object
    Public RichTextBox5_Height As Object
    Public RichTextBox5_Left As Object
    Public RichTextBox5_Top As Object
    Public RichTextBox6_Width As Object
    Public RichTextBox6_Height As Object
    Public RichTextBox6_Left As Object
    Public RichTextBox6_Top As Object
    Public RichTextBox7_Width As Object
    Public RichTextBox7_Height As Object
    Public RichTextBox7_Left As Object
    Public RichTextBox7_Top As Object
    Public RichTextBox8_Width As Object
    Public RichTextBox8_Height As Object
    Public RichTextBox8_Left As Object
    Public RichTextBox8_Top As Object

    Dim MyPath As String
    Dim Name As String
    Dim adres As String
    Dim adres1 As String
    Dim adres2 As String
    Dim adres3 As String
    Dim adres4 As String
    Dim Stroka, Stroka1, Stroka2 As String
    Dim file_name As String
    Dim razdelitel As String
    Dim hFile1 As Integer
    Dim hFile2 As Integer
    Dim hFile3 As Integer
    Dim hFile4 As Integer

    Public Word1(2) As String
    Public PhoneticNotation1(2) As String
    Public Version1(2) As String
    Public Word2(2) As String
    Public PhoneticNotation2(2) As String
    Public Version2(2) As String
    Public Value1(2) As String
    Public Value2(2) As String
    Public Value3() As String
    Public Value3_() As String
    Public Value4(2) As String
    Public Value5(2) As String
    Public Value6(2) As String

    Dim i, j, k As Integer
    Dim str_IN, str_OUT As String

    Dim QuantitySyllable As Integer
    Dim Syllable() As Object
    Dim StartPozSyllable() As Integer
    Dim LenSyllable() As Integer

    Public UniqueAddres1()
    Public UniqueBook1()
    Public UniqueAddres2()
    Public UniqueBook2()

    Public SelectIndex

    'Размеры полей GroupBox
    Public GB4_H As Integer = 80
    Public GB2_H As Integer = 108
    Public GB1_H As Integer = 234
    Public sDGC_H As Integer = 311

    'форматирование ячеек таблицы
    Private Sub songsDataGridView_CellFormatting(ByVal sender As Object,
     ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles songsDataGridView.CellFormatting
        If e IsNot Nothing Then
            If Me.songsDataGridView.Columns(e.ColumnIndex).Name = "Release Date" Then
                If e.Value IsNot Nothing Then
                    Try
                        e.Value = DateTime.Parse(e.Value.ToString()).ToLongDateString()
                        e.FormattingApplied = True
                    Catch ex As FormatException
                        Console.WriteLine("{0} is not a valid date.", e.Value.ToString())
                    End Try
                End If
            End If
        End If
    End Sub

    'добавляем DataGridView со всеми настройками
    Private Sub SetupDataGridView()
        Me.Controls.Add(songsDataGridView)
        songsDataGridView.ColumnCount = 13 'задаём число столбцов
        With songsDataGridView.ColumnHeadersDefaultCellStyle
            .BackColor = Color.Navy
            .ForeColor = Color.White
            .Font = New System.Drawing.Font(songsDataGridView.Font, FontStyle.Bold)
        End With
        With songsDataGridView
            .Name = "songsDataGridView"
            .Location = New System.Drawing.Point(12, 506)
            .Size = New Size(560, 311)
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .CellBorderStyle = DataGridViewCellBorderStyle.Single
            .GridColor = Color.Black
            .RowHeadersVisible = False
            .Columns(0).Name = "Index"
            .Columns(1).Name = "Sheet (From)"
            .Columns(1).DefaultCellStyle.ForeColor = Color.Black
            .Columns(1).DefaultCellStyle.Font = New System.Drawing.Font("Arial", 12, FontStyle.Bold)
            .Columns(2).Name = "Cells (From)"
            .Columns(2).DefaultCellStyle.ForeColor = Color.Green
            .Columns(2).DefaultCellStyle.Font = New System.Drawing.Font("Lucida Sans Unicode", 12, FontStyle.Regular)
            .Columns(3).Name = "help-text (From)"
            .Columns(3).DefaultCellStyle.ForeColor = Color.DimGray
            .Columns(3).DefaultCellStyle.Font = New System.Drawing.Font("Times New Roman", 12, FontStyle.Italic)
            .Columns(4).Name = "Sheet (To)"
            .Columns(4).DefaultCellStyle.ForeColor = Color.Black
            .Columns(4).DefaultCellStyle.Font = New System.Drawing.Font("Arial", 12, FontStyle.Bold)
            .Columns(5).Name = "Cells (To)"
            .Columns(5).DefaultCellStyle.ForeColor = Color.Green
            .Columns(5).DefaultCellStyle.Font = New System.Drawing.Font("Lucida Sans Unicode", 12, FontStyle.Regular)
            .Columns(6).Name = "help-text (To)"
            .Columns(6).DefaultCellStyle.ForeColor = Color.DimGray
            .Columns(6).DefaultCellStyle.Font = New System.Drawing.Font("Times New Roman", 12, FontStyle.Italic)
            .Columns(7).Name = "Variant"
            .Columns(7).DefaultCellStyle.ForeColor = Color.Green
            .Columns(7).DefaultCellStyle.Font = New System.Drawing.Font("Lucida Sans Unicode", 12, FontStyle.Regular)
            .Columns(8).Name = "Variable (From)"
            .Columns(8).DefaultCellStyle.ForeColor = Color.Maroon
            .Columns(8).DefaultCellStyle.Font = New System.Drawing.Font("Verdana", 12, FontStyle.Bold)
            .Columns(9).Name = "Formula (To)"
            .Columns(9).DefaultCellStyle.ForeColor = Color.RoyalBlue
            .Columns(9).DefaultCellStyle.Font = New System.Drawing.Font("Trebuchet MS", 9.75, FontStyle.Bold)
            .Columns(10).Name = "Adres (From)"
            .Columns(10).DefaultCellStyle.ForeColor = Color.Black
            .Columns(10).DefaultCellStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
            .Columns(11).Name = "Adres (To)"
            .Columns(11).DefaultCellStyle.ForeColor = Color.Black
            .Columns(11).DefaultCellStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
            .Columns(12).Name = "## ## ##"
            .Columns(12).DefaultCellStyle.ForeColor = Color.Black
            .Columns(12).DefaultCellStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            '.Dock = DockStyle.Fill
        End With

        'изменение размеров и локализации
        Dim a, b As Double
        Form1_Width = Me.Size.Width
        a = Form1_Width

        Form1_Height = Me.Size.Height
        b = Form1_Height

        'songsDataGridView_Width = 560 + (a - 600) * 560 / 600
        songsDataGridView_Width = 560 + (a - 600)
        songsDataGridView.Width = songsDataGridView_Width

        songsDataGridView_Height = 311 + (b - 887) * 311 / 887
        songsDataGridView.Height = songsDataGridView_Height

        'songsDataGridView_Left = 12 + (a - 600) * 12 / 600
        'songsDataGridView.Left = songsDataGridView_Left

        songsDataGridView_Top = 506 + (b - 887) * 506 / 887
        songsDataGridView.Top = songsDataGridView_Top
    End Sub
    'Заполнение DataGridView данными
    Private Sub PopulateDataGridView()
        Dim i, j, g As Integer
        Me.songsDataGridView.Rows.Clear()

        For i = 0 To Gindex - 1
            Dim row_i As String() = {i + 1, Word1(i), PhoneticNotation1(i), Version1(i), Word2(i), PhoneticNotation2(i), Version2(i), Value1(i), Value2(i), Value3(i), Value4(i), Value5(i), Value6(i)}
            With Me.songsDataGridView.Rows
                .Add(row_i)
            End With
            With songsDataGridView

                If j = 1 Then
                    j = 0
                    GoTo g1
                ElseIf j = 0 Then
                    j = 1
                End If
g1:
                If j = 1 Then .Rows(i).DefaultCellStyle.BackColor = Color.FromKnownColor(92 - 1) '3
                If j = 0 Then .Rows(i).DefaultCellStyle.BackColor = Color.FromKnownColor(94 - 1) '18
            End With
        Next i
        With Me.songsDataGridView 'определяем порядок столбцов
            .Columns(0).DisplayIndex = 0
            .Columns(1).DisplayIndex = 1
            .Columns(2).DisplayIndex = 2
            .Columns(3).DisplayIndex = 3
            .Columns(4).DisplayIndex = 6 '5 '4
            .Columns(5).DisplayIndex = 7 '6 '5
            .Columns(6).DisplayIndex = 8 '7 '6
            .Columns(7).DisplayIndex = 9 '8 '7
            .Columns(8).DisplayIndex = 4 '8
            .Columns(9).DisplayIndex = 10 '9
            .Columns(10).DisplayIndex = 5 '10
            .Columns(11).DisplayIndex = 11
            .Columns(12).DisplayIndex = 12
        End With

        For i = LBound(Value2) To UBound(Value2)
            For j = i To UBound(Value2)
                If i <> j Then
                    If Value2(i) = Value2(j) And Value2(i) <> "Nothing" Then _
                       MsgBox("Переменная с именем " & Value2(i) &
                              " первый раз встречается в строке " & i + 1 &
                              ", второй раз встречается в строке " & j + 1 &
                              ". Имена переменных должны быть уникальными." &
                              " Повторение имён переменных приведёт к неверному исполнению задуманного алгоритма.")
                End If
            Next j
        Next i

        'обработчик этики использования разных имён для исходной книги которая будет скопирована как копия
        For i = LBound(Value6) To UBound(Value6)
            For j = i To UBound(Value6)
                If i <> j Then
                    If Value6(i) = Value6(j) And
                        Value6(i) <> "Nothing" And Value6(j) <> "Nothing" And
                        Value6(i) <> Nothing And Value6(j) <> Nothing Then
                        MsgBox("Упрощённое имя копии книги " & Value6(i) &
                              " первый раз встречается в строке " & i + 1 &
                              ", второй раз встречается в строке " & j + 1)
                    ElseIf Value6(i) <> Value6(j) And
                        Value6(i) <> "Nothing" And Value6(j) <> "Nothing" And
                        Value6(i) <> Nothing And Value6(j) <> Nothing Then
                        MsgBox("Упрощённое имя копии книги " & Value6(i) &
                              " первый раз встречается в строке " & i + 1 &
                              ", Другое имя копии книги встречается в строке " & j + 1 &
                              ". Достаточно указать только одно имя. Не предполагается использование больше одного упрощённого имени.")
                    End If
                End If
            Next j
        Next i
    End Sub
    <STAThreadAttribute()>
    Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.Run(New Form1())
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim adres As String
        OpenFileDialog1.Filter = "Все файлы (*.*)|*.*|Книга Microsoft Office Excel 2003 (*.xls)|*.xls"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = False
        OpenFileDialog1.ShowDialog()
        adres = OpenFileDialog1.FileName
        TextBox1.Text = adres
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim adres As String
        OpenFileDialog1.Filter = "Все файлы (*.*)|*.*|Книга Microsoft Office Excel 2003 (*.xls)|*.xls"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = False
        OpenFileDialog1.ShowDialog()
        adres = OpenFileDialog1.FileName
        TextBox2.Text = adres
    End Sub

    Private Sub Button22_Click(sender As System.Object, e As System.EventArgs) Handles Button22.Click
        Dim adres As String
        OpenFileDialog1.Filter = "Все файлы (*.*)|*.*|Книга Microsoft Office Excel 2003 (*.xls)|*.xls"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = False
        OpenFileDialog1.ShowDialog()
        adres = OpenFileDialog1.FileName
        RichTextBox9.Text = adres
    End Sub

    Private Sub Button23_Click(sender As System.Object, e As System.EventArgs) Handles Button23.Click
        Dim adres As String
        OpenFileDialog1.Filter = "Все файлы (*.*)|*.*|Книга Microsoft Office Excel 2003 (*.xls)|*.xls"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = False
        OpenFileDialog1.ShowDialog()
        adres = OpenFileDialog1.FileName
        RichTextBox10.Text = adres
    End Sub

    Private Sub Button10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        ChangePattern()
    End Sub

    'Изменяем шаблон
    Private Sub ChangePattern()
        Dim adres As String
        Dim Name As Object
        Dim NameLR As String
        OpenFileDialog1.InitialDirectory = CurDir()
        OpenFileDialog1.Filter = "Файл шаблона (*.LR)|*.LR|Все файлы (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = False
        OpenFileDialog1.ShowDialog()
        adres = OpenFileDialog1.FileName
        Name = My.Computer.FileSystem.GetName(adres)
        NameLR = Mid(Name, 1, Len(Name) - 4)
        Me.TextBox4.Text = NameLR

        MyPath = CurDir()
        hFile1 = FreeFile()
        FileOpen(hFile1, MyPath & "\" & "range.tmp", OpenMode.Output)
        NameLR = Me.TextBox4.Text 'печать имени шаблона 
        PrintLine(hFile1, NameLR)
        adres1 = MyPath & "\" & NameLR & "1.LR"
        adres2 = MyPath & "\" & NameLR & "2.LR"
        adres3 = MyPath & "\" & NameLR & "3.LR"
        adres4 = MyPath & "\" & NameLR & "4.LR"
        PrintLine(hFile1, adres1)
        PrintLine(hFile1, adres2)
        If My.Computer.FileSystem.FileExists(adres3) = True Then 'СОВМЕСТИМОСТЬ: создаём отсутствующий файл "...3.LR"
            PrintLine(hFile1, adres3)
        Else
            PrintLine(hFile1, "") 'не уверен в этом решении
        End If
        If My.Computer.FileSystem.FileExists(adres4) = True Then 'СОВМЕСТИМОСТЬ: создаём отсутствующий файл "...4.LR"
            PrintLine(hFile1, adres4)
        Else
            PrintLine(hFile1, "") 'не уверен в этом решении
        End If

        FileClose(hFile1)

        LoadDataRichTextBox()
        SetupDataGridView()
        PopulateDataGridView()

        Me.Button6.Enabled = True
        Me.Button9.Enabled = False
        Me.Button7.Enabled = False
        Me.Button8.Enabled = True
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Open_Close()
        Button24.Enabled = True
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer
        Dim Name As Object
        Dim NameLR As String
        Dim NameR As String

        Dim adres1 As String
        Dim adres2 As String
        adres1 = TextBox1.Text
        adres2 = TextBox2.Text

        ' Раннее связывание
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application With {.Visible = True}
        Dim xlsBook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsBook2 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsSheet As Microsoft.Office.Interop.Excel.Worksheet



        '' Позднее связывание
        'Dim xlsApp As Object
        'Dim xlsBook1 As Object
        'Dim xlsBook2 As Object
        'Dim xlsSheet As Object
        'xlsApp = CreateObject("Excel.Application")

        'обработчик этики использования разных имён для исходной книги которая будет скопирована как копия
        For i = LBound(Value6) To UBound(Value6)
            If Value6(i) <> "Nothing" Then
                Name = My.Computer.FileSystem.GetName(adres1)
                NameLR = Mid(Name, 1, Len(Name) - 4)
                NameR = Mid(Name, Len(Name) - 4)
            End If
        Next i

        ' Поздно связать экземпляр книги Excel.
        'xlsBook = xlsApp.Workbooks.add
        xlsBook1 = xlsApp.Workbooks.Open(adres1) 'Открываем существующий Excel файл
        xlsBook2 = xlsApp.Workbooks.Open(adres2)

        ' Поздно связать экземпляр листа Excel.
        'xlsSheet = xlsBook.Worksheets(1)
        xlsSheet = xlsBook1.Sheets(1) 'Получаем лист по его имени (или порядковому номеру, если передаём Integer, т.е. xlsBook.Sheets(0))
        xlsSheet.Activate()

        ' Показать приложение.
        xlsSheet.Application.Visible = True
    End Sub

    'ПЕРЕНОС ДАННЫХ МЕЖДУ ЛИСТАМИ
    Public Sub Open_Close()
        ' Раннее связывание
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application With {.Visible = True}
        Dim xlsBook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim oBook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlsBook2 As Microsoft.Office.Interop.Excel.Workbook
        Dim oBook2 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsSheet2 As Microsoft.Office.Interop.Excel.Worksheet
        Dim var As Integer
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim iLink As Object = ""
        Dim iFormula As Object = ""
        Dim Ceels_from As Object
        Dim Ceels_to As Object

        Dim iRow1 As Long
        Dim iRow2 As Long
        Dim iClm1 As Integer
        Dim iClm2 As Integer
        Dim reiteration As Boolean

        Dim adres1 As String
        Dim adres2 As String
        adres1 = TextBox1.Text
        adres2 = TextBox2.Text
        Dim Name As String
        Dim NameLR As String
        Dim NameR As String
        Dim adres_new As String
        adres_new = ""
        'обработчик этики использования разных имён для исходной книги которая будет скопирована как копия
        For i = LBound(Value6) To UBound(Value6)
            If Value6(i) <> "Nothing" And Value6(i) <> Nothing Then
                Name = My.Computer.FileSystem.GetName(adres1)
                'NameLR = Mid(Name, 1, Len(Name) - 5)
                'NameR = Mid(Name, Len(Name) - 4, 4)
                NameR = Path.GetExtension(adres1)
                NameLR = Replace(Name, NameR, "")
                If NameLR <> Value6(i) Then
                    adres_new = Replace(adres1, Name, Value6(i) & NameR)
                    My.Computer.FileSystem.CopyFile(adres1, adres_new)
                    adres1 = adres_new
                End If
                Exit For
            End If
        Next i

        '' Позднее связывание
        'Dim xlsApp As Object
        'Dim xlsBook1 As Object
        'Dim xlsSheet1 As Object
        'Dim xlsBook2 As Object
        'Dim xlsSheet2 As Object
        'xlsApp = CreateObject("Excel.Application")
        '
        '' Поздно связать экземпляр книги Excel.
        ''xlsBook = xlsApp.Workbooks.add
        'xlsBook = xlsApp.Workbooks.Open(adres) 'Открываем существующий Excel файл
        '
        '' Поздно связать экземпляр листа Excel.
        ''xlsSheet = xlsBook.Worksheets(1)
        'xlsSheet = xlsBook.Sheets(1) 'Получаем лист по его имени (или порядковому номеру, если передаём Integer, т.е. xlsBook.Sheets(0))
        'xlsSheet.Activate()
        '
        '' Показать приложение.
        'xlsSheet.Application.Visible = True

        If adres1 <> adres2 Then
            'If Trim(TextBox1.Text) <> "" Then
            'xlsBook1 = xlsApp.Workbooks.Open(TextBox1.Text) 'Открываем существующий Excel файл
            xlsBook1 = xlsApp.Workbooks.Open(adres1,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, True)
            'If Trim(TextBox2.Text) <> "" Then 
            'xlsBook2 = xlsApp.Workbooks.Open(TextBox2.Text) 'Открываем существующий Excel файл
            xlsBook2 = xlsApp.Workbooks.Open(adres2,
                                             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                             Type.Missing, Type.Missing, Type.Missing, True)
        ElseIf adres1 = adres2 And adres1 <> "" And adres2 <> "" Then
            'If Trim(TextBox1.Text) <> "" Then 
            'xlsBook1 = xlsApp.Workbooks.Open(TextBox1.Text) 'Открываем существующий Excel файл
            xlsBook1 = xlsApp.Workbooks.Open(adres1,
                                             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                             Type.Missing, Type.Missing, Type.Missing, True)
            xlsBook2 = xlsBook1
        End If

        k = 0
        ReDim UniqueAddres1(0)
        ReDim UniqueBook1(0)
        ReDim UniqueAddres2(0)
        ReDim UniqueBook2(0)
        ReDim Value3_(0)
        For i = LBound(Value3) To UBound(Value3)
            ReDim Preserve Value3_(i)
            Value3_(i) = Value3(i)
        Next
        For i = 0 To UBound(Word1) - 1
            If Value4(i) <> "Nothing" And Value4(i) <> "" And Trim(CStr(Value4(i))) <> "0" Then
                reiteration = False
                For j = LBound(UniqueAddres1) To UBound(UniqueAddres1)
                    If Trim(UniqueAddres1(j)) = Trim(Value4(i)) Then
                        reiteration = True
                        oBook1 = UniqueBook1(j)
                    End If
                Next j
                If reiteration = False Then
                    If i = 0 Then
                        oBook1 = xlsApp.Workbooks.Open(Value4(i))
                    Else
                        If Value4(i) <> Value5(i) Then
                            oBook1 = xlsApp.Workbooks.Open(Value4(i))
                        Else
                            oBook1 = oBook2
                        End If
                    End If
                    ReDim Preserve UniqueAddres1(k)
                    ReDim Preserve UniqueBook1(k)
                    UniqueAddres1(k) = Value4(i)
                    UniqueBook1(k) = oBook1
                    k = k + 1
                End If
            ElseIf Value4(i) <> "GLOBAL" Then 'Ссылка на глобальный источник книги из TextBox1 (Бахтизин попросил, б...н)
                oBook1 = xlsBook1
            Else
                oBook1 = xlsBook1
            End If
            xlsSheet1 = CType(oBook1.Sheets.Item(Word1(i)), Excel.Worksheet) 'Получаем лист по его имени (или порядковому номеру, если передаём Integer, т.е. xlsBook.Sheets(0))

            If Value5(i) <> "Nothing" And Value5(i) <> "" And Trim(CStr(Value5(i))) <> "0" Then
                reiteration = False
                For j = LBound(UniqueAddres2) To UBound(UniqueAddres2)
                    If Trim(UniqueAddres2(j)) = Trim(Value5(i)) Then
                        reiteration = True
                        oBook2 = UniqueBook2(j)
                    End If
                Next j
                If reiteration = False Then
                    'oBook2 = xlsApp.Workbooks.Open(Value5(I))
                    If Value5(i) <> Value4(i) Then
                        oBook2 = xlsApp.Workbooks.Open(Value5(i))
                    Else
                        oBook2 = oBook1
                    End If
                    ReDim Preserve UniqueAddres2(k)
                    ReDim Preserve UniqueBook2(k)
                    UniqueAddres2(k) = Value5(i)
                    UniqueBook2(k) = oBook2
                    k = k + 1
                End If
            ElseIf Value5(i) <> "GLOBAL" Then 'Ссылка на глобальный источник книги из TextBox2 (Бахтизин попросил, б...н)
                oBook2 = xlsBook2
            Else
                oBook2 = xlsBook2
            End If
            xlsSheet2 = CType(oBook2.Sheets.Item(Word2(i)), Excel.Worksheet)

            'xlsApp.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual   'Отключение втоматического 
            'xlsSheet1.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual 'режима пересчёта ячеек
            'xlsSheet2.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual

            xlsSheet1.Range(PhoneticNotation1(i)).Copy() '

            If Value1(i) <> "Nothing" And Value1(i) <> "" And Trim(CStr(Value1(i))) <> "0" Then
                var = Val(Value1(i))
            Else
                var = 2
            End If

            If Value3_(i) <> "Nothing" And Value3_(i) <> "" And Trim(CStr(Value3_(i))) <> "0" Then
                For j = LBound(Value2) To UBound(Value2) 'перебираем весь список переменных
                    If Value2(j) <> "Nothing" And Value2(j) <> "" And Trim(CStr(Value2(j))) <> "0" Then 'если переменая в данной строке определена
                        iFormula = Value3_(i)
                        If InStr(Value3_(i), Value2(j)) <> 0 Then 'если переменная встречается в формуле
                            GetLinkOfVariable(j, iLink)
                            iFormula = Replace(CStr(Value3_(i)), CStr(Value2(j)), CStr(iLink))
                            Value3_(i) = iFormula
                        End If
                    End If
                Next j

                If Value1(i) = 3 Or Value1(i) = 2 Then 'вставка формулы
                    If InStrRev(PhoneticNotation2(i), ":") <> 0 Then
                        'Ceels_from = Mid(PhoneticNotation2(I), 1, InStrRev(PhoneticNotation2(I), ":") - 1)
                        iClm1 = xlsSheet2.Range(PhoneticNotation2(i)).Cells(1).Column()
                        iRow1 = xlsSheet2.Range(PhoneticNotation2(i)).Cells(1).Row()
                        Ceels_from = БукваСтолбца(xlsApp, iClm1) & iRow1
                    Else
                        Ceels_from = PhoneticNotation2(i)
                    End If
                    xlsSheet2.Range(Ceels_from).FormulaLocal = "=" & iFormula
                    xlsSheet2.Range(Ceels_from).Copy() 'ToDo: программа в скомпилированном виде выполняет это
                    ExcelPasteSettingFunction(xlsSheet2, PhoneticNotation2(i), var)
                ElseIf Value1(i) = 2 Then 'вставка значения
                    xlsSheet2.Range(PhoneticNotation2(i)).Copy()
                    ExcelPasteSettingFunction(xlsSheet2, PhoneticNotation2(i), var)
                End If
            Else
                ExcelPasteSettingFunction(xlsSheet2, PhoneticNotation2(i), var) 'отработка обычных сценариев копирования вставки
            End If

        Next i

        'удаляем временную копию файла с коротким именем
        If adres_new <> "" And adres1 <> adres2 Then
            xlsBook1.Close()
            My.Computer.FileSystem.DeleteFile(adres_new)
        End If

    End Sub

    Public Sub ExcelPasteSettingFunction(ByVal iSheet As Microsoft.Office.Interop.Excel.Worksheet,
                                         ByVal stroka As String,
                                         ByVal PasteVariant As Integer)
        Select Case PasteVariant
            Case 1 'Вставить (А)
                'ReversMergeCells(iSheet, stroka)
                iSheet.Range(stroka).PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                                                                           Operation:=Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                                                           SkipBlanks:=False, Transpose:=False)
                'ReversMergeCells(iSheet, stroka)
            Case 2 'Значения (З)
                iSheet.Range(stroka).PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues,
                                                                           Operation:=Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                                                           SkipBlanks:=False, Transpose:=False)
            Case 3 'Формулы (Ф)
                iSheet.Range(stroka).PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas,
                                                                           Operation:=Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                                                           SkipBlanks:=False, Transpose:=False)
            Case 4 'Транспонировать (А)
                iSheet.Range(stroka).PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                                                                           Operation:=Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                                                           SkipBlanks:=False, Transpose:=True)
            Case 5 'Форматирование (Ф)
                iSheet.Range(stroka).PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats,
                                                                           Operation:=Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                                                           SkipBlanks:=False, Transpose:=False)
            Case 6 'Вставить связь (Ь)
                iSheet.Range(stroka).PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValidation,
                                                                           Operation:=Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                                                           SkipBlanks:=False, Transpose:=False)
        End Select
    End Sub

    'Получение ссылки на переменную
    Public Sub GetLinkOfVariable(ByVal j As Integer, ByRef iLink As Object)
        Dim Sheet_from As Object
        Dim Ceels_from As Object
        Dim AdresFrom As String

        AdresFrom = Trim(TextBox1.Text)
        If AdresFrom <> "Nothing" And AdresFrom <> "" And AdresFrom <> "0" Then
        Else
            AdresFrom = Value4(j)
        End If
        Sheet_from = Word1(j)
        Ceels_from = PhoneticNotation1(j)
        iLink = "'" &
                Mid(AdresFrom, 1, InStrRev(AdresFrom, "\") - 1) & Replace(AdresFrom, "\", "\[", InStrRev(AdresFrom, "\")) &
                "]" &
                Sheet_from &
                "'!" &
        Ceels_from
    End Sub

    Function БукваСтолбца(ByRef xlsApp As Object, ByVal col As Long) As String
        On Error Resume Next
        БукваСтолбца = xlsApp.ConvertFormula("r1c" & col, XlReferenceStyle.xlR1C1, XlReferenceStyle.xlA1)
        БукваСтолбца = Replace(Replace(Mid(БукваСтолбца, 2), "$", ""), "1", "")
    End Function

    ''ПОСТРОЕНИЕ ДВУМЕРНОЙ ДИАГРАММЫ
    'Public Sub BildChart(ByVal xlsSheet As Microsoft.Office.Interop.Excel.Worksheet, _
    '                     ByVal xlsBook As Microsoft.Office.Interop.Excel.Workbook, _
    '                     ByVal xlsApp As Microsoft.Office.Interop.Excel.Application, _
    '                     ByVal Nomer As Integer, _
    '                     ByVal ChisloSech As Integer, _
    '                     ByVal N As Integer, _
    '                     ByVal ColymsX As Integer, _
    '                     ByVal ColymsY As Integer, _
    '                     ByVal Chart_name As String, _
    '                     ByVal Comment_X As String, _
    '                     ByVal Comment_Y As String, _
    '                     ByVal Chart_Location As Boolean, _
    '                     ByVal Colyms_Location As Integer, _
    '                     ByVal Rows_Location As Integer)
    'Dim xlsRange As Microsoft.Office.Interop.Excel.Range
    'Dim oChart As Microsoft.Office.Interop.Excel.Chart
    'Dim oSeries As Microsoft.Office.Interop.Excel.Series
    'Dim SeriesCol As Microsoft.Office.Interop.Excel.SeriesCollection
    'Dim SK As String
    'Dim OsbX As Microsoft.Office.Interop.Excel.Range
    'Dim OsbY As Microsoft.Office.Interop.Excel.Range
    'Dim PodpisbX As String
    'Dim PodpisbY As String
    'Dim a3, a4, a5, a6 As Integer
    'Dim i, i_1, i_2 As Integer
    '
    '    xlsApp.ReferenceStyle = Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1 'Стиль ссылок "R1C1"
    '    SK = Chart_name 'Заголовок диаграммы
    ''xlsApp.ReferenceStyle = Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1
    '    PodpisbX = Comment_X
    '    PodpisbY = Comment_Y
    ''Application.DisplayAlerts = False ' Это свойство позволяет отключить показ различных предупреждений
    '
    ''oChart = xlsBook.Charts.Add ' Создание диаграммы
    '    oChart = xlsSheet.Parent.Charts.Add()
    '    oChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterSmooth ' Тип диаграммы
    '    oChart.Name = SK ' Имя диаграммы
    '    oChart.SizeWithWindow = True 'Размер диаграммы будет подогнан таким образом, чтобы точно соответство-вать размеру листа
    '    oChart.Tab.ColorIndex = 35 ' Настройка внешнего вида вкладки диаграммы в книге
    '    oChart.HasLegend = True
    '
    ''Добавление на диаграмму первого сечения
    '    OsbX = xlsSheet.Range(xlsSheet.Cells(21, ColymsX), xlsSheet.Cells(N + 21, ColymsX))
    '    OsbY = xlsSheet.Range(xlsSheet.Cells(21, ColymsY), xlsSheet.Cells(N + 21, ColymsY))
    '    xlsRange = xlsApp.Union(OsbX, OsbY)
    '    oChart.SetSourceData(Source:=xlsRange, PlotBy:=Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)
    '    oChart.SeriesCollection(1).Name = "n=" + xlsSheet.Range("B" & 21).Text '.CharSeriesName
    ''oChart.SeriesCollection(1).CharSeriesName = "первый" '.CharSeriesName
    '    i_1 = 20 + N + 3 : i_2 = (20 + N + 2) + N
    '    For i = 2 To ChisloSech
    '        With oChart
    '            .SeriesCollection.NewSeries()
    '            .SeriesCollection(i).Name = "n=" + xlsSheet.Range("B" & i_1).Text
    '            .SeriesCollection(i).XValues = "=Лист1!R" + Convert.ToString(i_1) + "C" + Convert.ToString(ColymsX) + ":R" + Convert.ToString(i_2) + "C" + Convert.ToString(ColymsX)
    '            .SeriesCollection(i).Values = "=Лист1!R" + Convert.ToString(i_1) + "C" + Convert.ToString(ColymsY) + ":R" + Convert.ToString(i_2) + "C" + Convert.ToString(ColymsY)
    '        End With
    '        i_1 = i_1 + N + 2
    '        i_2 = i_2 + N + 2
    '    Next i
    '
    '''Добавление на диаграмму последующих сечений
    ''For i = 2 To ChisloSech
    ''a3 = 3 + (6 * i - 6)
    ''a4 = 4 + (6 * i - 6)
    ''a5 = 5 + (6 * i - 6)
    ''a6 = 6 + (6 * i - 6)
    ''With oChart
    '' .SeriesCollection.NewSeries()
    '' .SeriesCollection(2 * i - 1).Name = "sech" + Convert.ToString(i) + "_spin"
    '' .SeriesCollection(2 * i - 1).XValues = "=AprResults!R" + Convert.ToString(4) + "C" + Convert.ToString(a3) + ":R" + Convert.ToString(Nomer(i) + 3) + "C" + Convert.ToString(a3)
    '' .SeriesCollection(2 * i - 1).Values = "=AprResults!R" + Convert.ToString(4) + "C" + Convert.ToString(a4) + ":R" + Convert.ToString(Nomer(i) + 3) + "C" + Convert.ToString(a4)
    '' End With
    '' With oChart
    '' .SeriesCollection.NewSeries()
    '' .SeriesCollection(2 * i).Name = "sech" + Convert.ToString(i) + "_kor"
    '' .SeriesCollection(2 * i).XValues = "=AprResults!R" + Convert.ToString(4) + "C" + Convert.ToString(a5) + ":R" + Convert.ToString(Nomer(i) + 3) + "C" + Convert.ToString(a5)
    '' .SeriesCollection(2 * i).Values = "=AprResults!R" + Convert.ToString(4) + "C" + Convert.ToString(a6) + ":R" + Convert.ToString(Nomer(i) + 3) + "C" + Convert.ToString(a6)
    '' End With
    '' Next i
    '
    '    With oChart
    '        .HasTitle = True 'Наличие заголовка диаграммы
    '        .ChartTitle.Text = SK 'Текст заголовка диавграммы
    '    End With
    '    With oChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue) 'форматирование по оси Y
    '        .HasTitle = True
    '        With .AxisTitle
    '            .Caption = PodpisbY
    '            .Font.Name = "Arial Cyr"
    '            .Font.Size = 10
    ''.Characters(2, 2).Font.Italic = True
    ''.Characters(2, 2).Font.Size = 8
    '        End With
    '    End With
    '    With oChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory) 'форматирование по оси X
    '        .HasTitle = True
    '        With .AxisTitle
    '            .Caption = PodpisbX
    '            .Font.Name = "Arial Cyr"
    '            .Font.Size = 10
    '        End With
    '        .HasMajorGridlines = True
    '        .MajorGridlines.Border.Color = RGB(0, 0, 0)
    '        .MajorGridlines.Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '    End With
    '    oChart.Tab.ColorIndex = 35 'Окраска вкладки диаграммы
    ''Application.DisplayAlerts = True ' Это свойство позволяет отключить показ различных предупреждений
    '    xlsApp.ReferenceStyle = Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1 'Стиль ссылок "A1"
    '
    '    If Chart_Location = True Then
    '        oChart.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsObject, xlsSheet.Name)
    '' Переместить диаграмму, чтобы не покрыть ваши данные.
    ''With xlsSheet.Shapes.Item("Chart 1")
    '        With xlsSheet.Shapes.Item(Nomer)
    '            .Top = xlsSheet.Rows(Rows_Location).Top
    '            .Left = xlsSheet.Columns(Colyms_Location).Left
    '        End With
    '    End If
    '
    'End Sub

    'ограницение использования по годам
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ''TextBox1.Text = Year(DateTime.Now)
        ''If Year(DateTime.Now) > 2016 Then End
        'If (DateTime.Now < "13.09.2017") Or (DateTime.Now > "13.09.2018") Then                                           ' Workbooks(Bookname).Worksheets("Ëèñò3").Cells(1, 23) > Date - 3 Or
        '    End
        'End If
    End Sub

    Private Sub Form1_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Me.Width = My.Settings.x_del
        'Me.Height = My.Settings.y_del
        LoadDataRichTextBox()
        SetupDataGridView()
        PopulateDataGridView()
    End Sub

    'СЧИТЫВАНИЕ И ЗАГРУЗКА ДАННЫХ в поле просмотра
    Public Sub LoadDataRichTextBox()
        Dim i, fsf, i1, i2, i3 As Integer
        Dim Position As Integer
        Dim MyPath As Object
        Dim Name As Object
        Dim NameLR As String
        Dim stroka As String

        MyPath = CurDir() '"C:\Documents and Settings"
        Name = "" : NameLR = ""
        If My.Computer.FileSystem.FileExists(MyPath & "\" & "range.tmp") = True Then
            adres1 = Trim(ReadLine(MyPath & "\" & "range.tmp", 2))
            adres2 = Trim(ReadLine(MyPath & "\" & "range.tmp", 3))
            adres3 = Trim(ReadLine(MyPath & "\" & "range.tmp", 4))
            adres4 = Trim(ReadLine(MyPath & "\" & "range.tmp", 5))
            If adres1 <> "" And adres2 <> "" Then
                Name = My.Computer.FileSystem.GetName(adres2)
                NameLR = Mid(Name, 1, Len(Name) - 4)
            End If
        End If

        If My.Computer.FileSystem.FileExists(adres1) = True And
        My.Computer.FileSystem.FileExists(adres2) = True Then
            If NameLR <> "" Then TextBox4.Text = NameLR 'печать имени шаблона 
        ElseIf My.Computer.FileSystem.FileExists(adres1) = False And
        My.Computer.FileSystem.FileExists(adres2) = False Then
            hFile1 = FreeFile()
            FileOpen(hFile1, MyPath & "\" & "range.tmp", OpenMode.Output)
            PrintLine(hFile1, Name)
            NameLR = InputBox("Введите имя для шаблона")
            adres1 = MyPath & "\" & NameLR & "1.LR"
            adres2 = MyPath & "\" & NameLR & "2.LR"
            adres3 = MyPath & "\" & NameLR & "3.LR"
            adres4 = MyPath & "\" & NameLR & "4.LR"
            PrintLine(hFile1, adres1)
            PrintLine(hFile1, adres2)
            PrintLine(hFile1, adres3)
            PrintLine(hFile1, adres4)
            FileClose(hFile1)

            'создаём файл MyBase1.LR
            hFile1 = FreeFile()
            FileOpen(hFile1, adres1, OpenMode.Output)
            FileClose(hFile1)

            'создаём файл MyBase2.LR
            hFile1 = FreeFile()
            FileOpen(hFile1, adres2, OpenMode.Output)
            FileClose(hFile1)

            'создаём файл MyBase3.LR
            hFile1 = FreeFile()
            FileOpen(hFile1, adres3, OpenMode.Output)
            FileClose(hFile1)

            'создаём файл MyBase4.LR
            hFile1 = FreeFile()
            FileOpen(hFile1, adres4, OpenMode.Output)
            FileClose(hFile1)
        End If

        'Загрузка базы данных MyBase1 в оперативную память
        ReadMyBase(adres1, Word1, PhoneticNotation1, Version1, Gindex)

        'Загрузка базы данных MyBase2 в оперативную память
        ReadMyBase(adres2, Word2, PhoneticNotation2, Version2, Gindex)

        If adres3 = "" Then 'СОВМЕСТИМОСТЬ: создаём отсутствующий файл "...3.LR"
            adres1 = MyPath & "\" & NameLR & "1.LR"
            adres2 = MyPath & "\" & NameLR & "2.LR"
            adres3 = MyPath & "\" & NameLR & "3.LR"
            'adres4 = MyPath & "\" & NameLR & "4.LR"

            hFile1 = FreeFile()
            FileOpen(hFile1, MyPath & "\" & "range.tmp", OpenMode.Output)
            PrintLine(hFile1, Name)
            PrintLine(hFile1, adres1)
            PrintLine(hFile1, adres2)
            PrintLine(hFile1, adres3)
            'PrintLine(hFile1, adres4)
            FileClose(hFile1)
            'создаём файл MyBase3.LR
            hFile1 = FreeFile()
            FileOpen(hFile1, adres3, OpenMode.Output)
            For i = 1 To Gindex
                stroka = "|_78_111_116_104_105_110_103|_78_111_116_104_105_110_103|_78_111_116_104_105_110_103|"
                PrintLine(hFile1, stroka)
            Next i
            PrintLine(hFile1, "")
            FileClose(hFile1)
            If i >= Gindex + 1 Then DeleteLine(adres3, Gindex + 1)
        End If
        ReadMyBase(adres3, Value1, Value2, Value3, Gindex)

        If adres4 = "" Then 'СОВМЕСТИМОСТЬ: создаём отсутствующий файл "...4.LR"
            adres1 = MyPath & "\" & NameLR & "1.LR"
            adres2 = MyPath & "\" & NameLR & "2.LR"
            adres3 = MyPath & "\" & NameLR & "3.LR"
            adres4 = MyPath & "\" & NameLR & "4.LR"

            hFile1 = FreeFile()
            FileOpen(hFile1, MyPath & "\" & "range.tmp", OpenMode.Output)
            PrintLine(hFile1, Name)
            PrintLine(hFile1, adres1)
            PrintLine(hFile1, adres2)
            PrintLine(hFile1, adres3)
            PrintLine(hFile1, adres4)
            FileClose(hFile1)
            'создаём файл MyBase4.LR
            hFile1 = FreeFile()
            FileOpen(hFile1, adres4, OpenMode.Output)
            For i = 1 To Gindex
                stroka = "|_78_111_116_104_105_110_103|_78_111_116_104_105_110_103|_78_111_116_104_105_110_103|"
                PrintLine(hFile1, stroka)
            Next i
            PrintLine(hFile1, "")
            FileClose(hFile1)
            If i >= Gindex + 1 Then DeleteLine(adres4, Gindex + 1)
        End If
        ReadMyBase(adres4, Value4, Value5, Value6, Gindex)

    End Sub

    'Загрузка и форматирование данных в RichTextBox1
    Public Sub ReadMyBase(ByVal adres As String,
                          ByRef txt1() As String,
                          ByRef txt2() As String,
                          ByRef txt3() As String,
                          ByRef Gindex As Integer)
        Dim i As Integer
        Dim MassData() As String

        ReadFile2015(adres, MassData)
        ReDim txt1(MassData.Length)
        ReDim txt2(MassData.Length)
        ReDim txt3(MassData.Length)
        For i = 0 To MassData.Length - 1
            j = 1 : Stroka1 = ""
            Stroka = MassData(i)
            If InStrRev(Stroka, "п»ї") <> 0 Then
                Stroka = Replace(Stroka, "п»ї", "") 'заебали эти каракули в начале файла
                ReplaceText2015(adres, i, Stroka)
            End If
            Stroka = ConvertStroka(Stroka)
            HL(j, k, Stroka) : txt1(i) = Mid(Stroka, k, j - k)
            HL(j, k, Stroka) : txt2(i) = Mid(Stroka, k, j - k)
            HL(j, k, Stroka) : txt3(i) = Mid(Stroka, k, j - k)
        Next i
        Gindex = i
    End Sub
    Private Sub HL(ByRef j, ByRef k, ByRef Stroka)
        j = j + 1 : Stroka1 = "" : k = j
        Do While Stroka1 <> "|"
            Stroka1 = Mid(Stroka, j, 1)
            j = j + 1
        Loop
        j = j - 1
    End Sub

    'Добавить слово
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim stroka3 As String
        Dim i As Long
        Dim var As Integer
        Dim str_var As String

        If RichTextBox3.Text <> "" And RichTextBox4.Text <> "" And RichTextBox5.Text <> "" And
        RichTextBox6.Text <> "" And RichTextBox7.Text <> "" And RichTextBox8.Text <> "" Then
            hFile1 = FreeFile()
            FileOpen(hFile1, adres1, OpenMode.Append)

            hFile2 = FreeFile()
            FileOpen(hFile2, adres2, OpenMode.Append)

            If My.Computer.FileSystem.FileExists(adres3) = True Then 'Если файл ...3.LR есть
                hFile3 = FreeFile()
                FileOpen(hFile3, adres3, OpenMode.Append)
            End If

            If My.Computer.FileSystem.FileExists(adres4) = True Then 'Если файл ...4.LR есть
                hFile4 = FreeFile()
                FileOpen(hFile4, adres4, OpenMode.Append)
            End If

            Stroka = ""
            For i = 1 To Len(RichTextBox3.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox3.Text), i, 1))), "13", "32")
                Stroka = Stroka + "_" + stroka3
            Next i
            Stroka1 = ""
            For i = 1 To Len(RichTextBox4.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox4.Text), i, 1))), "13", "32")
                Stroka1 = Stroka1 + "_" + stroka3
            Next i
            Stroka2 = ""
            For i = 1 To Len(RichTextBox5.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox5.Text), i, 1))), "13", "32")
                Stroka2 = Stroka2 + "_" + stroka3
            Next i
            Stroka = "|" & Stroka & "|" & Stroka1 & "|" & Stroka2 & "|"
            PrintLine(hFile1, Stroka)
            FileClose(hFile1)

            RichTextBox3.Text = "" : RichTextBox4.Text = "" : RichTextBox5.Text = ""

            Stroka = ""
            For i = 1 To Len(RichTextBox6.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox6.Text), i, 1))), "13", "32")
                Stroka = Stroka + "_" + stroka3
            Next i
            Stroka1 = ""
            For i = 1 To Len(RichTextBox7.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox7.Text), i, 1))), "13", "32")
                Stroka1 = Stroka1 + "_" + stroka3
            Next i
            Stroka2 = ""
            For i = 1 To Len(RichTextBox8.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox8.Text), i, 1))), "13", "32")
                Stroka2 = Stroka2 + "_" + stroka3
            Next i
            Stroka = "|" & Stroka & "|" & Stroka1 & "|" & Stroka2 & "|"
            PrintLine(hFile2, Stroka)
            FileClose(hFile2)

            RichTextBox6.Text = "" : RichTextBox7.Text = "" : RichTextBox8.Text = ""

            If Me.Button16.BackColor = Color.Red Then var = "1"
            If Me.Button17.BackColor = Color.Red Then var = "2"
            If Me.Button18.BackColor = Color.Red Then var = "3"
            If Me.Button19.BackColor = Color.Red Then var = "4"
            If Me.Button20.BackColor = Color.Red Then var = "5"
            If Me.Button21.BackColor = Color.Red Then var = "6"

            If var = 0 Then
                str_var = "2"
            Else
                str_var = Trim(Str(var))
            End If
            Stroka = ""
            For i = 1 To Len(str_var)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(str_var), i, 1))), "13", "32")
                Stroka = Stroka + "_" + stroka3
            Next i
            Stroka1 = ""
            For i = 1 To Len(RichTextBox1.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox1.Text), i, 1))), "13", "32")
                Stroka1 = Stroka1 + "_" + stroka3
            Next i
            Stroka2 = RichTextBox2.Text
            'If InStrRev(Stroka2, "=") <> 0 Then
            'Stroka2 = Replace(Stroka2, "=", "") 'ликвидируем знак "="
            'End If
            RichTextBox2.Text = Stroka2
            Stroka2 = Trim(RichTextBox2.Text)
            If Stroka2 = "" Then Stroka2 = "Nothing"
            RichTextBox2.Text = Stroka2
            Stroka2 = ""
            For i = 1 To Len(RichTextBox2.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox2.Text), i, 1))), "13", "32")
                Stroka2 = Stroka2 + "_" + stroka3
            Next i
            Stroka = "|" & Stroka & "|" & Stroka1 & "|" & Stroka2 & "|"
            PrintLine(hFile3, Stroka)
            FileClose(hFile3)
            RichTextBox1.Text = "" : RichTextBox2.Text = ""

            Stroka2 = Trim(RichTextBox9.Text)
            If Stroka2 = "" Then Stroka2 = "Nothing"
            RichTextBox9.Text = Stroka2
            Stroka = ""
            For i = 1 To Len(RichTextBox9.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox9.Text), i, 1))), "13", "32")
                Stroka = Stroka + "_" + stroka3
            Next i
            Stroka2 = Trim(RichTextBox10.Text)
            If Stroka2 = "" Then Stroka2 = "Nothing"
            RichTextBox10.Text = Stroka2
            Stroka1 = ""
            For i = 1 To Len(RichTextBox10.Text)
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(RichTextBox10.Text), i, 1))), "13", "32")
                Stroka1 = Stroka1 + "_" + stroka3
            Next i

            Stroka2 = Trim(TextBox5.Text)
            If Stroka2 = "" Then Stroka2 = "Nothing"
            TextBox5.Text = Stroka2
            Stroka2 = ""
            'For i = 1 To Len("Nothing")
            For i = 1 To Len(TextBox5.Text)
                'stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim("Nothing"), i, 1))), "13", "32")
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(TextBox5.Text), i, 1))), "13", "32")
                Stroka2 = Stroka2 + "_" + stroka3
            Next i
            Stroka = "|" & Stroka & "|" & Stroka1 & "|" & Stroka2 & "|"
            PrintLine(hFile4, Stroka)
            FileClose(hFile4)
            RichTextBox9.Text = "" : RichTextBox10.Text = "" : TextBox5.Text = ""

        Else
            If RichTextBox3.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите ИМЯ листа Excel книги для КОПИРОВАНИЯ")
            ElseIf RichTextBox4.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите наименование ССЫЛКИ на листе Excel книги для КОПИРОВАНИЯ")
            ElseIf RichTextBox5.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите комментарий для пояснения данной ссылки")
            ElseIf RichTextBox6.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите ИМЯ листа Excel книги для ВСТАВКИ")
            ElseIf RichTextBox7.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите наименование ССЫЛКИ на листе Excel книги для ВСТАВКИ")
            ElseIf RichTextBox8.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите комментарий для пояснения данной ссылки")
                'ElseIf TextBox5.Text = "" Then
                '    MsgBox("Недостаточно информации для ввода. Введите ...")
            End If
        End If

        FileClose(hFile1)
        FileClose(hFile2)
        If My.Computer.FileSystem.FileExists(adres3) = True Then 'Если файл ...3.LR есть
            FileClose(hFile3)
        End If
        If My.Computer.FileSystem.FileExists(adres4) = True Then 'Если файл ...4.LR есть
            FileClose(hFile4)
        End If
        EnabledEditButtons(False)
        LoadDataRichTextBox()
        PopulateDataGridView()
    End Sub

    'Изменить строку
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim var As Integer
        Dim str_var As String
        Dim str1, str2, str3 As String

        If Trim(RichTextBox3.Text) <> "" And Trim(RichTextBox4.Text) <> "" And Trim(RichTextBox5.Text) <> "" And
           Trim(RichTextBox6.Text) <> "" And Trim(RichTextBox7.Text) <> "" And Trim(RichTextBox8.Text) <> "" Then

            'ВХОДНЫЕ ССЫЛКИ
            iBaseReplace(adres1,
                         RichTextBox3.Text,
                         RichTextBox4.Text,
                         RichTextBox5.Text,
                         Val(Trim(Label10.Text)),
                         Val(Trim(TextBox3.Text)))
            RichTextBox3.Text = "" : RichTextBox4.Text = "" : RichTextBox5.Text = ""

            'ВЫХОДНЫЕ ССЫЛКИ
            iBaseReplace(adres2,
                         RichTextBox6.Text,
                         RichTextBox7.Text,
                         RichTextBox8.Text,
                         Val(Trim(Label10.Text)),
                         Val(Trim(TextBox3.Text)))
            RichTextBox6.Text = "" : RichTextBox7.Text = "" : RichTextBox8.Text = ""

            'НАСТРОЕЧНЫЕ ССЫЛКИ
            If Me.Button16.BackColor = Color.Red Then var = "1"
            If Me.Button17.BackColor = Color.Red Then var = "2"
            If Me.Button18.BackColor = Color.Red Then var = "3"
            If Me.Button19.BackColor = Color.Red Then var = "4"
            If Me.Button20.BackColor = Color.Red Then var = "5"
            If Me.Button21.BackColor = Color.Red Then var = "6"
            If var = 0 Then
                str_var = "2"
            Else
                str_var = Trim(Str(var))
            End If
            Stroka2 = RichTextBox2.Text
            'If InStrRev(Stroka2, "=") <> 0 Then
            'Stroka2 = Replace(Stroka2, "=", "") 'ликвидируем знак "="
            'End If
            RichTextBox2.Text = Stroka2
            str1 = Trim(RichTextBox1.Text)
            If str1 = "" Then str1 = "Nothing"
            str2 = Trim(RichTextBox2.Text)
            If str2 = "" Then str2 = "Nothing"
            iBaseReplace(adres3,
                         str_var,
                         str1,
                         str2,
                         Val(Trim(Label10.Text)),
                         Val(Trim(TextBox3.Text)))
            RichTextBox1.Text = "" : RichTextBox2.Text = ""

            'Ссылки на документы копирования и вставки
            str1 = Trim(RichTextBox9.Text)
            If str1 = "" Then str1 = "Nothing"
            str2 = Trim(RichTextBox10.Text)
            If str2 = "" Then str2 = "Nothing"
            str3 = Trim(TextBox5.Text)
            If str3 = "" Then str3 = "Nothing"
            iBaseReplace(adres4,
                         str1,
                         str2,
                         str3,
                         Val(Trim(Label10.Text)),
                         Val(Trim(TextBox3.Text)))
            RichTextBox9.Text = "" : RichTextBox10.Text = "" : TextBox5.Text = ""

            'Обновление списка слов в главном окне
            TextBox3.Text = ""
            Label10.Text = ""
            EnabledEditButtons(False)
            LoadDataRichTextBox()
            PopulateDataGridView()
            songsDataGridView.CurrentCell = songsDataGridView.Item(0, SelectIndex) 'благодаря этому при большом количестве строк еще и прокрутка таблицы произойдет
            songsDataGridView.Rows(SelectIndex).Selected = True ' ну и выделяем строку
        Else
            If RichTextBox3.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите ИМЯ листа Excel книги для КОПИРОВАНИЯ")
            ElseIf RichTextBox4.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите наименование ССЫЛКИ на листе Excel книги для КОПИРОВАНИЯ")
            ElseIf RichTextBox5.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите комментарий для пояснения данной ссылки")
            ElseIf RichTextBox6.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите ИМЯ листа Excel книги для ВСТАВКИ")
            ElseIf RichTextBox7.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите наименование ССЫЛКИ на листе Excel книги для ВСТАВКИ")
            ElseIf RichTextBox8.Text = "" Then
                MsgBox("Недостаточно информации для ввода. Введите комментарий для пояснения данной ссылки")
                'ElseIf TextBox5.Text = "" Then
                '    MsgBox("Недостаточно информации для ввода. Введите ...")
            End If
        End If

    End Sub

    'iBaseReplace(adres, Text1, Text2, Text3, iGindex, FutureGindex)
    Public Sub iBaseReplace(ByVal adres As String,
                            ByVal Text1 As String,
                            ByVal Text2 As String,
                            ByVal Text3 As String,
                            ByVal iGindex As Integer,
                            ByVal FutureGindex As Integer)
        Dim NomerStr As Integer
        Dim Stroka, Stroka1, Stroka2, stroka3, str As String
        Dim LenStrOld As Integer
        Dim LenStrNew As Integer
        Dim i As Integer

        If FutureGindex <> 0 Then
            NomerStr = FutureGindex
            If Gindex < NomerStr Then
                Exit Sub
            ElseIf 1 > NomerStr Then
                Exit Sub
            End If
        Else
            NomerStr = iGindex
        End If
        LenStrOld = Chislo_Strok(adres)
        If FutureGindex = 0 Or FutureGindex = iGindex Then 'если мы не собирваемся менять индекс записи
            Stroka = ""
            For i = 1 To Len(Text1)
                stroka3 = System.Convert.ToString(AscW(Mid(Trim(Text1), i, 1)))
                Stroka = Stroka + "_" + stroka3
                Stroka = Replace(Stroka, "п»ї", "") 'заебали эти каракули в начале файла
            Next i
            Stroka1 = ""
            For i = 1 To Len(Text2)
                stroka3 = System.Convert.ToString(AscW(Mid(Trim(Text2), i, 1)))
                Stroka1 = Stroka1 + "_" + stroka3
                Stroka1 = Replace(Stroka1, "п»ї", "") 'заебали эти каракули в начале файла
            Next i
            Stroka2 = ""
            For i = 1 To Len(Text3)
                stroka3 = System.Convert.ToString(AscW(Mid(Trim(Text3), i, 1)))
                Stroka2 = Stroka2 + "_" + stroka3
                Stroka2 = Replace(Stroka2, "п»ї", "") 'заебали эти каракули в начале файла
            Next i
            Stroka = "|" & Stroka & "|" & Stroka1 & "|" & Stroka2 & "|"
            ReplaceText2015(adres, NomerStr - 1, Stroka)
        ElseIf FutureGindex <> 0 And FutureGindex <> iGindex Then 'если меняем индекс записи
            If FutureGindex < iGindex Then 'в начало списка
                str = System.Convert.ToString(ReadLine(adres, iGindex))
                str = Replace(str, "п»ї", "") 'заебали эти каракули в начале файла
                For i = iGindex To FutureGindex + 1 Step -1
                    Stroka = ReadLine(adres, i - 1)
                    Stroka = Replace(Stroka, "п»ї", "") 'заебали эти каракули в начале файла
                    ReplaceText2015(adres, i - 1, Stroka)
                Next i
                ReplaceText2015(adres, FutureGindex - 1, str)
            ElseIf FutureGindex > iGindex Then 'в конец списка
                str = System.Convert.ToString(ReadLine(adres, iGindex))
                str = Replace(str, "п»ї", "") 'заебали эти каракули в начале файла
                For i = iGindex To FutureGindex - 1
                    Stroka = ReadLine(adres, i + 1)
                    Stroka = Replace(Stroka, "п»ї", "") 'заебали эти каракули в начале файла
                    ReplaceText2015(adres, i - 1, Stroka)
                Next i
                ReplaceText2015(adres, FutureGindex - 1, str)
            End If
        End If
        LenStrNew = Chislo_Strok(adres)                            'плавающая неисправность
        If LenStrNew > LenStrOld Then DeleteLine(adres, LenStrNew - 1) 'иногда добавляется лишняя строка
    End Sub

    'Удаление записи
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim NomerStr As Long

        NomerStr = Convert.ToDecimal(Label10.Text)
        DeleteLine(adres1, NomerStr - 1)
        DeleteLine(adres2, NomerStr - 1)
        If My.Computer.FileSystem.FileExists(adres3) = True Then DeleteLine(adres3, NomerStr - 1)
        If My.Computer.FileSystem.FileExists(adres4) = True Then DeleteLine(adres4, NomerStr - 1)

        RichTextBox1.Text = "" : RichTextBox2.Text = "" : RichTextBox3.Text = "" : RichTextBox4.Text = "" : RichTextBox5.Text = ""
        RichTextBox6.Text = "" : RichTextBox7.Text = "" : RichTextBox8.Text = "" : RichTextBox9.Text = "" : RichTextBox10.Text = ""
        TextBox5.Text = ""

        TextBox3.Text = ""
        Label10.Text = ""
        EnabledEditButtons(False)
        LoadDataRichTextBox()
        PopulateDataGridView()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        RichTextBox1.Text = "" : RichTextBox2.Text = "" : RichTextBox3.Text = "" : RichTextBox4.Text = "" : RichTextBox5.Text = ""
        RichTextBox6.Text = "" : RichTextBox7.Text = "" : RichTextBox8.Text = "" : RichTextBox9.Text = "" : RichTextBox10.Text = ""
        TextBox5.Text = ""

        TextBox3.Text = ""
        Label10.Text = ""
        EnabledEditButtons(False)
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadDataRichTextBox()
        PopulateDataGridView()
    End Sub
    Private Sub ИзменитьШаблонToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ИзменитьШаблонToolStripMenuItem.Click
        ChangePattern()
    End Sub
    'определение активности кнопок редактирования
    Sub EnabledEditButtons(ByVal stroka As Boolean)
        Me.TextBox3.Enabled = stroka
        Me.Button6.Enabled = stroka
        Me.Button9.Enabled = stroka
        Me.Button7.Enabled = stroka
        Me.Button8.Enabled = stroka
        EnabledEditButtons4(stroka)
    End Sub
    'определение активности кнопок редактирования
    Sub EnabledEditButtons2(ByVal stroka As Boolean)
        Me.Button11.Enabled = stroka
        Me.Button15.Enabled = stroka
        Me.Button12.Enabled = stroka
        Me.Button13.Enabled = stroka
        Me.Button14.Enabled = stroka
    End Sub
    Sub EnabledEditButtons3(ByVal stroka As Boolean)
        Me.Button6.Enabled = stroka
        Me.Button7.Enabled = stroka
        EnabledEditButtons4(stroka)
    End Sub
    Sub EnabledEditButtons4(ByVal stroka As Boolean)
        Me.RichTextBox1.Enabled = stroka
        Me.RichTextBox2.Enabled = stroka
        Me.RichTextBox4.Enabled = stroka
        Me.RichTextBox5.Enabled = stroka
        Me.RichTextBox7.Enabled = stroka
        Me.RichTextBox8.Enabled = stroka
        Me.RichTextBox9.Enabled = stroka
        Me.RichTextBox10.Enabled = stroka
        Me.TextBox5.Enabled = stroka
        Me.Button16.Enabled = stroka : Button16.BackColor = Me.GroupBox1.BackColor
        Me.Button17.Enabled = stroka : Button17.BackColor = Me.GroupBox1.BackColor
        Me.Button18.Enabled = stroka : Button18.BackColor = Me.GroupBox1.BackColor
        Me.Button19.Enabled = stroka : Button19.BackColor = Me.GroupBox1.BackColor
        Me.Button20.Enabled = stroka : Button20.BackColor = Me.GroupBox1.BackColor
        Me.Button21.Enabled = stroka : Button21.BackColor = Me.GroupBox1.BackColor
        Me.Button22.Enabled = stroka 'Button22.BackColor = Me.GroupBox1.BackColor
        Me.Button23.Enabled = stroka 'Button23.BackColor = Me.GroupBox1.BackColor
    End Sub
    Sub VisibleGB4(ByVal stroka As Boolean)
        Me.Label8.Visible = stroka
        Me.TextBox4.Visible = stroka
        Me.Button10.Visible = stroka
        Me.Button11.Visible = stroka
        Me.Button12.Visible = stroka
        Me.Button13.Visible = stroka
        Me.Button14.Visible = stroka
        Me.Button15.Visible = stroka
    End Sub
    Sub VisibleGB2(ByVal stroka As Boolean)
        Me.Label1.Visible = stroka
        Me.TextBox1.Visible = stroka
        Me.Label2.Visible = stroka
        Me.TextBox2.Visible = stroka
        Me.Button1.Visible = stroka
        Me.Button2.Visible = stroka
    End Sub
    Sub VisibleGB1(ByVal stroka As Boolean)
        Me.Label3.Visible = stroka
        Me.Label4.Visible = stroka
        Me.Label5.Visible = stroka
        Me.Label6.Visible = stroka
        Me.Label7.Visible = stroka
        Me.RichTextBox1.Visible = stroka
        Me.RichTextBox2.Visible = stroka
        Me.RichTextBox3.Visible = stroka
        Me.RichTextBox4.Visible = stroka
        Me.RichTextBox5.Visible = stroka
        Me.RichTextBox6.Visible = stroka
        Me.RichTextBox7.Visible = stroka
        Me.RichTextBox8.Visible = stroka
        Me.RichTextBox9.Visible = stroka
        Me.RichTextBox10.Visible = stroka
        Me.TextBox5.Visible = stroka
        Me.Label9.Visible = stroka
        Me.Label11.Visible = stroka
        Me.TextBox3.Visible = stroka
        Me.Button6.Visible = stroka
        Me.Button9.Visible = stroka
        Me.Button7.Visible = stroka
        Me.Button8.Visible = stroka
        Me.Button16.Visible = stroka
        Me.Button17.Visible = stroka
        Me.Button18.Visible = stroka
        Me.Button19.Visible = stroka
        Me.Button20.Visible = stroka
        Me.Button21.Visible = stroka
        Me.Button22.Visible = stroka
        Me.Button23.Visible = stroka
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        End
    End Sub
    Private Sub RichTextBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox3.Click
        Me.Button6.Enabled = True : Me.Button8.Enabled = True
        EnabledEditButtons4(True)
    End Sub
    Private Sub RichTextBox4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox4.Click
        Me.Button6.Enabled = True : Me.Button8.Enabled = True
    End Sub
    Private Sub RichTextBox5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox5.Click
        Me.Button6.Enabled = True : Me.Button8.Enabled = True
    End Sub
    Private Sub RichTextBox6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox6.Click
        Me.Button6.Enabled = True : Me.Button8.Enabled = True
        EnabledEditButtons4(True)
    End Sub
    Private Sub RichTextBox7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox7.Click
        Me.Button6.Enabled = True : Me.Button8.Enabled = True
    End Sub
    Private Sub RichTextBox8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox8.Click
        Me.Button6.Enabled = True : Me.Button8.Enabled = True
    End Sub
    Private Sub ОПрограммеToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ОПрограммеToolStripMenuItem.Click
        'Form2.Show()
    End Sub
    Private Sub TextBox4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.Click
        EnabledEditButtons2(True)
    End Sub
    'Отменить
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim Name As Object
        Dim NameLR, NameLR2 As String

        MyPath = CurDir()
        adres = Trim(ReadLine(MyPath & "\" & "range.tmp", 2))
        Name = My.Computer.FileSystem.GetName(adres)
        NameLR = Mid(Name, 1, Len(Name) - 4)
        Me.TextBox4.Text = NameLR

        EnabledEditButtons2(False)
    End Sub
    'Удалить шаблон
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim NameLR As String

        NameLR = Me.TextBox4.Text
        MyPath = CurDir()
        If My.Computer.FileSystem.FileExists(MyPath & "\" & NameLR & "1.LR") = True Then
            hFile1 = FreeFile()
            FileOpen(hFile1, MyPath & "\" & "range.tmp", OpenMode.Output)
            PrintLine(hFile1, NameLR)
            adres1 = MyPath & "\" & NameLR & "1.LR"
            My.Computer.FileSystem.DeleteFile(adres1)
            adres2 = MyPath & "\" & NameLR & "2.LR"
            My.Computer.FileSystem.DeleteFile(adres2)
            adres3 = MyPath & "\" & NameLR & "3.LR"
            My.Computer.FileSystem.DeleteFile(adres3)
            adres4 = MyPath & "\" & NameLR & "4.LR"
            My.Computer.FileSystem.DeleteFile(adres4)
            PrintLine(hFile1, "")
            PrintLine(hFile1, "")
            FileClose(hFile1)
        Else
            MsgBox("Не найдено шаблона по введеному имени.")
        End If

        Me.TextBox4.Text = ""

        LoadDataRichTextBox()
        PopulateDataGridView()

        RichTextBox1.Text = "" : RichTextBox2.Text = "" : RichTextBox3.Text = "" : RichTextBox4.Text = "" : RichTextBox5.Text = ""
        RichTextBox6.Text = "" : RichTextBox7.Text = "" : RichTextBox8.Text = ""
        EnabledEditButtons(False)
        EnabledEditButtons2(False)
    End Sub
    'Переименовать шаблон
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim Name As Object
        Dim NameLR, NameLR2 As String

        MyPath = CurDir()
        adres = Trim(ReadLine(MyPath & "\" & "range.tmp", 2))
        Name = My.Computer.FileSystem.GetName(adres)
        NameLR = Mid(Name, 1, Len(Name) - 4)

        NameLR2 = Me.TextBox4.Text

        If NameLR <> NameLR2 Then
            hFile1 = FreeFile()
            FileOpen(hFile1, MyPath & "\" & "range.tmp", OpenMode.Output)
            PrintLine(hFile1, NameLR2)
            adres1 = MyPath & "\" & NameLR & "1.LR"
            My.Computer.FileSystem.RenameFile(adres1, NameLR2 & "1.LR")
            adres2 = MyPath & "\" & NameLR & "2.LR"
            My.Computer.FileSystem.RenameFile(adres2, NameLR2 & "2.LR")
            adres3 = MyPath & "\" & NameLR & "3.LR"
            My.Computer.FileSystem.RenameFile(adres3, NameLR2 & "3.LR")
            adres4 = MyPath & "\" & NameLR & "4.LR"
            My.Computer.FileSystem.RenameFile(adres4, NameLR2 & "4.LR")
            PrintLine(hFile1, MyPath & "\" & NameLR2 & "1.LR")
            PrintLine(hFile1, MyPath & "\" & NameLR2 & "2.LR")
            PrintLine(hFile1, MyPath & "\" & NameLR2 & "3.LR")
            PrintLine(hFile1, MyPath & "\" & NameLR2 & "4.LR")
            FileClose(hFile1)
        Else
            MsgBox("Вы не задали новое имя.")
        End If

        Me.TextBox4.Text = NameLR2

        LoadDataRichTextBox()
        PopulateDataGridView()

        RichTextBox1.Text = "" : RichTextBox2.Text = "" : RichTextBox3.Text = "" : RichTextBox4.Text = "" : RichTextBox5.Text = ""
        RichTextBox6.Text = "" : RichTextBox7.Text = "" : RichTextBox8.Text = ""
        EnabledEditButtons(False)
        EnabledEditButtons2(False)
    End Sub
    'клонировать шаблон
    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim NameLR As String

        NameLR = Me.TextBox4.Text
        MyPath = CurDir()
        If My.Computer.FileSystem.FileExists(MyPath & "\" & NameLR & "1.LR") = True Then
            hFile1 = FreeFile()
            FileOpen(hFile1, MyPath & "\" & "range.tmp", OpenMode.Output)
            PrintLine(hFile1, NameLR & "-копия")
            adres1 = MyPath & "\" & NameLR & "1.LR"
            My.Computer.FileSystem.CopyFile(adres1, MyPath & "\" & NameLR & "-копия1.LR")
            adres2 = MyPath & "\" & NameLR & "2.LR"
            My.Computer.FileSystem.CopyFile(adres2, MyPath & "\" & NameLR & "-копия2.LR")
            adres3 = MyPath & "\" & NameLR & "3.LR"
            My.Computer.FileSystem.CopyFile(adres3, MyPath & "\" & NameLR & "-копия3.LR")
            adres4 = MyPath & "\" & NameLR & "4.LR"
            My.Computer.FileSystem.CopyFile(adres4, MyPath & "\" & NameLR & "-копия4.LR")
            PrintLine(hFile1, MyPath & "\" & NameLR & "-копия1.LR")
            PrintLine(hFile1, MyPath & "\" & NameLR & "-копия2.LR")
            PrintLine(hFile1, MyPath & "\" & NameLR & "-копия3.LR")
            PrintLine(hFile1, MyPath & "\" & NameLR & "-копия4.LR")
            FileClose(hFile1)
        Else
            MsgBox("Не найдено шаблона по введеному имени.")
        End If

        Me.TextBox4.Text = NameLR & "-копия"

        LoadDataRichTextBox()
        PopulateDataGridView()

        RichTextBox1.Text = "" : RichTextBox2.Text = "" : RichTextBox3.Text = "" : RichTextBox4.Text = "" : RichTextBox5.Text = ""
        RichTextBox6.Text = "" : RichTextBox7.Text = "" : RichTextBox8.Text = ""
        EnabledEditButtons(False)
        EnabledEditButtons2(False)
    End Sub
    'Добавить новый шаблон
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim NameLR As String

        NameLR = Me.TextBox4.Text
        MyPath = CurDir()

        If My.Computer.FileSystem.FileExists(MyPath & "\" & NameLR & "1.LR") = False Then
            hFile1 = FreeFile()
            FileOpen(hFile1, MyPath & "\" & "range.tmp", OpenMode.Output)
            PrintLine(hFile1, NameLR)
            adres1 = MyPath & "\" & NameLR & "1.LR"
            adres2 = MyPath & "\" & NameLR & "2.LR"
            adres3 = MyPath & "\" & NameLR & "3.LR"
            adres4 = MyPath & "\" & NameLR & "4.LR"
            PrintLine(hFile1, adres1)
            PrintLine(hFile1, adres2)
            PrintLine(hFile1, adres3)
            PrintLine(hFile1, adres4)
            FileClose(hFile1)

            hFile1 = FreeFile()
            FileOpen(hFile1, adres1, OpenMode.Output)
            FileClose(hFile1)

            hFile1 = FreeFile()
            FileOpen(hFile1, adres2, OpenMode.Output)
            FileClose(hFile1)

            hFile1 = FreeFile()
            FileOpen(hFile1, adres3, OpenMode.Output)
            FileClose(hFile1)

            hFile1 = FreeFile()
            FileOpen(hFile1, adres4, OpenMode.Output)
            FileClose(hFile1)
        Else
            MsgBox("Шаблон с таким именем уже существует.")
            Exit Sub
        End If

        LoadDataRichTextBox()
        PopulateDataGridView()

        RichTextBox1.Text = "" : RichTextBox2.Text = "" : RichTextBox3.Text = "" : RichTextBox4.Text = "" : RichTextBox5.Text = ""
        RichTextBox6.Text = "" : RichTextBox7.Text = "" : RichTextBox8.Text = ""

        EnabledEditButtons(False)
        EnabledEditButtons2(False)
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ВыходToolStripMenuItem.Click
        End
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Dim a, b As Double
        Form1_Width = Me.Size.Width
        a = Form1_Width
        Form1_Height = Me.Size.Height
        b = Form1_Height

        'GroupBox4_Width = 560 + (a - 600) * 560 / 600 : GroupBox4.Width = GroupBox4_Width
        GroupBox4_Height = GB4_H + (b - 887) * GB4_H / 887 : GroupBox4.Height = GroupBox4_Height
        'GroupBox4_Left = 12 + (a - 600) * 12 / 600 : GroupBox4.Left = GroupBox4_Left
        GroupBox4_Top = 36 + (b - 887) * 36 / 887 : GroupBox4.Top = GroupBox4_Top

        'GroupBox2_Width = 560 + (a - 600) * 560 / 600 : GroupBox2.Width = GroupBox2_Width
        GroupBox2_Width = 560 + (a - 600) : GroupBox2.Width = GroupBox2_Width
        GroupBox2_Height = GB2_H + (b - 887) * GB2_H / 887 : GroupBox2.Height = GroupBox2_Height
        'GroupBox2_Left = 12 + (a - 600) * 12 / 600 : GroupBox2.Left = GroupBox2_Left
        'GroupBox2_Top = 132 + (b - 887) * 132 / 887 : GroupBox2.Top = GroupBox2_Top
        GroupBox2_Top = (52 + GB4_H) + (b - 887) * (52 + GB4_H) / 887 : GroupBox2.Top = GroupBox2_Top

        'GroupBox1_Width = 560 + (a - 600) * 560 / 600 : GroupBox1.Width = GroupBox1_Width
        GroupBox1_Width = 560 + (a - 600) : GroupBox1.Width = GroupBox1_Width
        GroupBox1_Height = GB1_H + (b - 887) * GB1_H / 887 : GroupBox1.Height = GroupBox1_Height
        'GroupBox1_Left = 12 + (a - 600) * 12 / 600 : GroupBox1.Left = GroupBox1_Left
        'GroupBox1_Top = 256 + (b - 887) * 256 / 887 : GroupBox1.Top = GroupBox1_Top
        GroupBox1_Top = (68 + GB4_H + GB2_H) + (b - 887) * (68 + GB4_H + GB2_H) / 887 : GroupBox1.Top = GroupBox1_Top

        sDGC_H = 311 + (80 - GB4_H) + (108 - GB2_H) + (234 - GB1_H)
        'songsDataGridView_Width = 560 + (a - 600) * 560 / 600 : songsDataGridView.Width = songsDataGridView_Width
        songsDataGridView_Width = 560 + (a - 600) : songsDataGridView.Width = songsDataGridView_Width
        songsDataGridView_Height = sDGC_H + (b - 887) * sDGC_H / 887 : songsDataGridView.Height = songsDataGridView_Height
        'songsDataGridView_Left = 12 + (a - 600) * 12 / 600 : songsDataGridView.Left = songsDataGridView_Left
        'songsDataGridView_Top = 506 + (b - 887) * 506 / 887 : songsDataGridView.Top = songsDataGridView_Top
        songsDataGridView_Top = (84 + GB4_H + GB2_H + GB1_H) + (b - 887) * (84 + GB4_H + GB2_H + GB1_H) / 887 : songsDataGridView.Top = songsDataGridView_Top

        Button3.Left = 279 + (a - 600)
        Button3.Top = 823 + (b - 887) * 823 / 887

        Button5.Left = 416 + (a - 600)
        Button5.Top = 823 + (b - 887) * 823 / 887

        Button4.Left = 497 + (a - 600)
        Button4.Top = 823 + (b - 887) * 823 / 887

        'Button24.Left = 12 + (a - 600) * 12 / 600
        Button24.Top = 823 + (b - 887) * 823 / 887
    End Sub

    Private Sub GroupBox4_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox4.Resize
        Dim a, b As Double
        GroupBox4_Width = Me.GroupBox4.Size.Width : a = GroupBox4_Width
        GroupBox4_Height = Me.GroupBox4.Size.Height : b = GroupBox4_Height

        Label8.Left = 6 + (a - 560) * 6 / 560
        Label8.Top = 26 + (b - 80) * 26 / 80

        TextBox4.Width = 469 + (a - 560) * 469 / 560
        TextBox4.Height = 23 + (b - 80) * 23 / 80
        TextBox4.Left = 56 + (a - 560) * 56 / 560
        TextBox4.Top = 23 + (b - 80) * 23 / 80

        Button10.Left = 531 + (a - 560) * 531 / 560
        Button10.Top = 22 + (b - 80) * 22 / 80

        Button11.Left = 102 + (a - 560)
        Button11.Top = 51 + (b - 80) * 51 / 80

        Button15.Left = 209 + (a - 560)
        Button15.Top = 51 + (b - 80) * 51 / 80

        Button12.Left = 296 + (a - 560)
        Button12.Top = 51 + (b - 80) * 51 / 80

        Button13.Left = 398 + (a - 560)
        Button13.Top = 51 + (b - 80) * 51 / 80

        Button14.Left = 479 + (a - 560)
        Button14.Top = 51 + (b - 80) * 51 / 80

        Button25.Left = 536 + (a - 560)
    End Sub

    Private Sub GroupBox2_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Resize
        Dim a, b As Double
        GroupBox2_Width = Me.GroupBox2.Size.Width : a = GroupBox2_Width
        GroupBox2_Height = Me.GroupBox2.Size.Height : b = GroupBox2_Height

        Label1.Left = 6 + (a - 560) * 6 / 560
        Label1.Top = 27 + (b - 108) * 27 / 108

        TextBox1.Width = 519 + (a - 560) * 519 / 560
        TextBox1.Height = 20 + (b - 108) * 20 / 108
        TextBox1.Left = 6 + (a - 560) * 6 / 560
        TextBox1.Top = 43 + (b - 108) * 43 / 108

        Button1.Left = 531 + (a - 560) * 531 / 560
        Button1.Top = 42 + (b - 108) * 42 / 108

        Label2.Left = 6 + (a - 560) * 6 / 560
        Label2.Top = 66 + (b - 108) * 66 / 108

        TextBox2.Width = 519 + (a - 560) * 519 / 560
        TextBox2.Height = 20 + (b - 108) * 20 / 108
        TextBox2.Left = 6 + (a - 560) * 6 / 560
        TextBox2.Top = 82 + (b - 108) * 82 / 108

        Button2.Left = 531 + (a - 560) * 531 / 560
        Button2.Top = 81 + (b - 108) * 81 / 108

        Button26.Left = 536 + (a - 560)
    End Sub

    Private Sub GroupBox1_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Resize
        Dim a, b As Double
        a = Me.GroupBox1.Size.Width
        b = Me.GroupBox1.Size.Height
        If GroupBox1_Width = Nothing Then GroupBox1_Width = Me.GroupBox1.Size.Width
        If GroupBox1_Height = Nothing Then GroupBox1_Height = Me.GroupBox1.Size.Height

        Label3.Left = 279 + (a - 560) * 279 / 560
        Label3.Top = 20 + (b - 234) * 20 / 234

        Label4.Left = 293 + (a - 560) * 293 / 560
        Label4.Top = 103 + (b - 234) * 103 / 234

        Label5.Left = 42 + (a - 560) * 42 / 560
        Label5.Top = 20 + (b - 234) * 20 / 234

        Label6.Left = 173 + (a - 560) * 173 / 560
        Label6.Top = 20 + (b - 234) * 20 / 234

        Label7.Left = 389 + (a - 560) * 389 / 560
        Label7.Top = 20 + (b - 234) * 20 / 234

        RichTextBox1.Left = 267 + (a - 560) * 267 / 560
        RichTextBox1.Top = 36 + (b - 234) * 36 / 234
        RichTextBox1.Width = 123 + (a - 560) * 123 / 560
        RichTextBox1.Height = 60 + (b - 234) * 60 / 234

        RichTextBox2.Left = 267 + (a - 560) * 267 / 560
        RichTextBox2.Top = 119 + (b - 234) * 119 / 234
        RichTextBox2.Width = 123 + (a - 560) * 123 / 560
        RichTextBox2.Height = 83 + (b - 234) * 83 / 234

        RichTextBox3.Width = 139 + (a - 560) * 139 / 560
        RichTextBox3.Height = 60 + (b - 234) * 60 / 234
        RichTextBox3.Left = 6 + (a - 560) * 6 / 560
        RichTextBox3.Top = 36 + (b - 234) * 36 / 234

        RichTextBox4.Width = 120 + (a - 560) * 120 / 560
        RichTextBox4.Height = 60 + (b - 234) * 60 / 234
        RichTextBox4.Left = 146 + (a - 560) * 146 / 560
        RichTextBox4.Top = 36 + (b - 234) * 36 / 234

        RichTextBox5.Width = 125 + (a - 560) * 125 / 560
        RichTextBox5.Height = 82 + (b - 234) * 82 / 234
        RichTextBox5.Left = 391 + (a - 560) * 391 / 560
        RichTextBox5.Top = 36 + (b - 234) * 36 / 234

        RichTextBox6.Width = 139 + (a - 560) * 139 / 560
        RichTextBox6.Height = 61 + (b - 234) * 61 / 234
        RichTextBox6.Left = 6 + (a - 560) * 6 / 560
        RichTextBox6.Top = 119 + (b - 234) * 119 / 234

        RichTextBox7.Width = 120 + (a - 560) * 120 / 560
        RichTextBox7.Height = 61 + (b - 234) * 61 / 234
        RichTextBox7.Left = 146 + (a - 560) * 146 / 560
        RichTextBox7.Top = 119 + (b - 234) * 119 / 234

        RichTextBox8.Width = 125 + (a - 560) * 125 / 560
        RichTextBox8.Height = 83 + (b - 234) * 83 / 234
        RichTextBox8.Left = 391 + (a - 560) * 391 / 560
        RichTextBox8.Top = 119 + (b - 234) * 119 / 234

        RichTextBox9.Width = 236 + (a - 560) * 236 / 560
        RichTextBox9.Height = 21 + (b - 234) * 21 / 234
        RichTextBox9.Left = 30 + (a - 560) * 30 / 560
        RichTextBox9.Top = 97 + (b - 234) * 97 / 234

        RichTextBox10.Width = 236 + (a - 560) * 236 / 560
        RichTextBox10.Height = 21 + (b - 234) * 21 / 234
        RichTextBox10.Left = 30 + (a - 560) * 30 / 560
        RichTextBox10.Top = 181 + (b - 234) * 181 / 234

        Label9.Left = 6 + (a - 560) * 6 / 560
        Label9.Top = 211 + (b - 234) * 211 / 234

        Label11.Left = 83 + (a - 560) * 83 / 560
        Label11.Top = 211 + (b - 234) * 211 / 234

        TextBox3.Left = 163 + (a - 560) * 163 / 560
        TextBox3.Top = 206 + (b - 234) * 206 / 234

        Button6.Left = 236 + (a - 560)
        Button6.Top = 206 + (b - 234) * 206 / 234

        Button9.Left = 317 + (a - 560)
        Button9.Top = 206 + (b - 234) * 206 / 234

        Button7.Left = 398 + (a - 560)
        Button7.Top = 206 + (b - 234) * 206 / 234

        Button8.Left = 479 + (a - 560)
        Button8.Top = 206 + (b - 234) * 206 / 234

        Button16.Left = 522 + (a - 560)
        Button16.Top = 10 + (b - 234) * 10 / 234

        Button17.Left = 522 + (a - 560)
        Button17.Top = 42 + (b - 234) * 42 / 234

        Button18.Left = 522 + (a - 560)
        Button18.Top = 74 + (b - 234) * 74 / 234

        Button19.Left = 522 + (a - 560)
        Button19.Top = 106 + (b - 234) * 106 / 234

        Button20.Left = 522 + (a - 560)
        Button20.Top = 138 + (b - 234) * 138 / 234

        Button21.Left = 522 + (a - 560)
        Button21.Top = 170 + (b - 234) * 170 / 234

        Button22.Left = 6 + (a - 560) * 6 / 560
        Button22.Top = 97 + (b - 234) * 97 / 234

        Button23.Left = 6 + (a - 560) * 6 / 560
        Button23.Top = 181 + (b - 234) * 181 / 234

        Button27.Left = 498 + (a - 560)

        GroupBox1_Width = a
        GroupBox1_Height = b
    End Sub

    Public Sub iObjectResizeSize(ByRef iObject_Parent As Object, ByRef iObject_Child As Object,
                                 ByVal X_Parent As Object, ByVal X_Child As Object,
                                 ByVal Y_Parent As Object, ByVal Y_Child As Object)
        iObject_Child.Width = X_Child + (iObject_Parent.Width - X_Parent) * X_Child / X_Parent
        iObject_Child.Height = Y_Child + (iObject_Parent.Height - Y_Parent) * Y_Child / Y_Parent
    End Sub

    Public Sub iObjectResizeLocation(ByRef iObject As Object, ByVal delX As Object, ByVal delY As Object)

    End Sub

    Private Sub ВернутьСтандартныйРазмерФормыToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ВернутьСтандартныйРазмерФормыToolStripMenuItem.Click
        Me.Width = 600
        Me.Height = 887
    End Sub

    'Информация
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim AppToRun As Object
        AppToRun = CurDir() & "\help.pdf"
        Diagnostics.Process.Start(AppToRun)
    End Sub

    Private Sub TextBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Click
        If Trim(Me.Label10.Text) <> "" Then
            Me.TextBox3.Enabled = True
        End If
    End Sub

    'Сохраняем ШАБЛОН в читаемом текстовом виде на листе EXCEL согласно структуры:
    ' |   A    |          B         |     C     |   D   |   E   |   F   |    G   |          H         |     I     |    J    |    K    |    L    |    M    |    N    |    O    |P|...
    '1|Word1(1)|PhoneticNotation1(1)|Version1(1)| запас | запас | запас |Word2(1)|PhoneticNotation2(1)|Version2(1)|Value1(1)|Value2(1)|Value3(1)|Value4(1)|Value5(1)|Value6(1)|
    '2|Word1(2)|PhoneticNotation1(2)|Version1(2)| запас | запас | запас |Word2(2)|PhoneticNotation2(2)|Version2(2)|Value1(2)|Value2(2)|Value3(2)|Value4(2)|Value5(2)|Value6(2)|    
    Private Sub ЭкспортироватьШаблонToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ЭкспортироватьШаблонToolStripMenuItem.Click
        Dim adres As String
        SaveFileDialog1.Filter = "Все файлы (*.*)|*.*|MS Excel 2010 (*.xlsx)|*.xlsx|MS Excel 2003 (*.xls)|*.xls|Текстовые файлы (*.txt)|*.txt"
        SaveFileDialog1.FilterIndex = 1
        SaveFileDialog1.RestoreDirectory = True
        SaveFileDialog1.FileName = Trim(Me.TextBox4.Text)
        SaveFileDialog1.ShowDialog()
        adres = SaveFileDialog1.FileName

        If SaveFileDialog1.FilterIndex = 1 Or SaveFileDialog1.FilterIndex = 2 Or SaveFileDialog1.FilterIndex = 3 Then
            Dim xlsAppE As New Microsoft.Office.Interop.Excel.Application With {.Visible = False}
            Dim xlsBookE As Microsoft.Office.Interop.Excel.Workbook
            xlsBookE = xlsAppE.Workbooks.Add()
            Dim xlsSheetE As Microsoft.Office.Interop.Excel.Worksheet
            xlsSheetE = xlsBookE.Sheets(1)
            Dim i As Long
            For i = 1 To Gindex
                xlsSheetE.Cells(i, 1) = Word1(i - 1)
                xlsSheetE.Cells(i, 2) = PhoneticNotation1(i - 1)
                xlsSheetE.Cells(i, 3) = Version1(i - 1)

                xlsSheetE.Cells(i, 4) = Value1(i - 1)
                xlsSheetE.Cells(i, 5) = Value2(i - 1)
                xlsSheetE.Cells(i, 6) = Value3(i - 1)

                xlsSheetE.Cells(i, 7) = Word2(i - 1)
                xlsSheetE.Cells(i, 8) = PhoneticNotation2(i - 1)
                xlsSheetE.Cells(i, 9) = Version2(i - 1)

                xlsSheetE.Cells(i, 10) = Value4(i - 1)
                xlsSheetE.Cells(i, 11) = Value5(i - 1)
                xlsSheetE.Cells(i, 12) = Value6(i - 1)
            Next
            xlsBookE = xlsAppE.ActiveWorkbook
            xlsBookE.SaveAs(adres)
            xlsBookE.Close()
        End If
        Me.RichTextBox1.Text = "" : Me.RichTextBox2.Text = "" : Me.RichTextBox3.Text = "" : Me.RichTextBox4.Text = "" : Me.RichTextBox5.Text = ""
        Me.RichTextBox6.Text = "" : Me.RichTextBox7.Text = "" : Me.RichTextBox8.Text = ""
        EnabledEditButtons(False)
        EnabledEditButtons2(False)
    End Sub

    'Загрузка шаблона хранящегося во внещнем файле EXCEL
    Private Sub ИмпортироватьШаблонToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ИмпортироватьШаблонToolStripMenuItem.Click
        Dim adresI, stroka, Stroka1, Stroka2, stroka3 As Object
        Dim i As Integer
        Dim MyPath As Object
        Dim Name As Object
        Dim NameLR As String

        'ищем файл в проводнике
        OpenFileDialog1.Filter = "Все файлы (*.*)|*.*|MS Excel 2010 (*.xlsx)|*.xlsx|MS Excel 2003 (*.xls)|*.xls|Текстовые файлы (*.txt)|*.txt"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = True
        OpenFileDialog1.ShowDialog()
        adresI = OpenFileDialog1.FileName

        Gindex = 10 ^ 6
        ReDim Word1(Gindex - 1)
        ReDim PhoneticNotation1(Gindex - 1)
        ReDim Version1(Gindex - 1)

        ReDim Value1(Gindex - 1)
        ReDim Value2(Gindex - 1)
        ReDim Value3(Gindex - 1)

        ReDim Word2(Gindex - 1)
        ReDim PhoneticNotation2(Gindex - 1)
        ReDim Version2(Gindex - 1)

        ReDim Value4(Gindex - 1)
        ReDim Value5(Gindex - 1)
        ReDim Value6(Gindex - 1)

        'считываем данные с листа EXCEL
        Dim xlsAppI As New Microsoft.Office.Interop.Excel.Application With {.Visible = False}
        Dim xlsBookI As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsSheetI As Microsoft.Office.Interop.Excel.Worksheet
        xlsBookI = xlsAppI.Workbooks.Open(adresI)
        xlsSheetI = xlsBookI.Sheets(1)
        i = 0 : stroka = "hera"
        Do While stroka <> ""
            i = i + 1
            stroka = Trim(xlsSheetI.Cells(i, 1).value)
            Word1(i - 1) = xlsSheetI.Cells(i, 1).value
            If Word1(i - 1) = Nothing Then Word1(i - 1) = "Nothing"
            PhoneticNotation1(i - 1) = xlsSheetI.Cells(i, 2).value
            If PhoneticNotation1(i - 1) = Nothing Then PhoneticNotation1(i - 1) = "Nothing"
            Version1(i - 1) = xlsSheetI.Cells(i, 3).value
            If Version1(i - 1) = Nothing Then Version1(i - 1) = "Nothing"
            Value1(i - 1) = xlsSheetI.Cells(i, 4).value
            If Value1(i - 1) = Nothing Then Value1(i - 1) = "Nothing"
            Value2(i - 1) = xlsSheetI.Cells(i, 5).value
            If Value2(i - 1) = Nothing Then Value2(i - 1) = "Nothing"
            Value3(i - 1) = xlsSheetI.Cells(i, 6).value
            If Value3(i - 1) = Nothing Then Value3(i - 1) = "Nothing"
            Word2(i - 1) = xlsSheetI.Cells(i, 7).value
            If Word2(i - 1) = Nothing Then Word2(i - 1) = "Nothing"
            PhoneticNotation2(i - 1) = xlsSheetI.Cells(i, 8).value
            If PhoneticNotation2(i - 1) = Nothing Then PhoneticNotation2(i - 1) = "Nothing"
            Version2(i - 1) = xlsSheetI.Cells(i, 9).value
            If Version2(i - 1) = Nothing Then Version2(i - 1) = "Nothing"
            Value4(i - 1) = xlsSheetI.Cells(i, 10).value
            If Value4(i - 1) = Nothing Then Value4(i - 1) = "Nothing"
            Value5(i - 1) = xlsSheetI.Cells(i, 11).value
            If Value5(i - 1) = Nothing Then Value5(i - 1) = "Nothing"
            Value6(i - 1) = xlsSheetI.Cells(i, 12).value
            If Value6(i - 1) = Nothing Then Value6(i - 1) = "Nothing"
        Loop
        Gindex = i - 1
        xlsBookI.Close()

        'проверяем есть ли шаблон с таким именем
        MyPath = CurDir()
        Name = "" : NameLR = ""
        Name = My.Computer.FileSystem.GetName(adresI)
        For i = Len(Name) - 1 To 1 Step -1
            stroka = Mid(Name, i, 1)
            If stroka = "." Then
                stroka = Mid(Name, 1, i - 1) & "1"
                Exit For
            End If
        Next
        NameLR = MyPath & "\" & stroka & ".LR"
        If My.Computer.FileSystem.FileExists(NameLR) = True Then
            stroka = Mid(Name, 1, i - 1) & "_" 'автоматически переименовываем шаблон
        Else
            stroka = Mid(Name, 1, i - 1)
        End If

        'создаём файл MyBase1.LR создаём файл MyBase2.LR и записываем соответствующую информацию в файл range
        hFile1 = FreeFile()
        FileOpen(hFile1, MyPath & "\" & "range.tmp", OpenMode.Output)
        Me.TextBox4.Text = stroka
        PrintLine(hFile1, stroka)
        NameLR = Name
        adres1 = MyPath & "\" & stroka & "1.LR"
        adres2 = MyPath & "\" & stroka & "2.LR"
        adres3 = MyPath & "\" & stroka & "3.LR"
        adres4 = MyPath & "\" & stroka & "4.LR"
        PrintLine(hFile1, adres1)
        PrintLine(hFile1, adres2)
        PrintLine(hFile1, adres3)
        PrintLine(hFile1, adres4)
        FileClose(hFile1)
        hFile1 = FreeFile()
        FileOpen(hFile1, adres1, OpenMode.Output)
        FileClose(hFile1)
        hFile1 = FreeFile()
        FileOpen(hFile1, adres2, OpenMode.Output)
        FileClose(hFile1)
        hFile1 = FreeFile()
        FileOpen(hFile1, adres3, OpenMode.Output)
        FileClose(hFile1)
        hFile1 = FreeFile()
        FileOpen(hFile1, adres4, OpenMode.Output)
        FileClose(hFile1)

        'загрузка данных в файлы "...1.LR", "...2.LR" и "...3.LR" и "...4.LR"
        hFile1 = FreeFile()
        FileOpen(hFile1, adres1, OpenMode.Append)
        hFile2 = FreeFile()
        FileOpen(hFile2, adres2, OpenMode.Append)
        hFile3 = FreeFile()
        FileOpen(hFile3, adres3, OpenMode.Append)
        hFile4 = FreeFile()
        FileOpen(hFile4, adres4, OpenMode.Append)
        For j = 1 To Gindex
            stroka = ""
            For i = 1 To Len(Word1(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Word1(j - 1)), i, 1))), "13", "32")
                stroka = stroka + "_" + stroka3
            Next i
            Stroka1 = ""
            For i = 1 To Len(PhoneticNotation1(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(PhoneticNotation1(j - 1)), i, 1))), "13", "32")
                Stroka1 = Stroka1 + "_" + stroka3
            Next i
            Stroka2 = ""
            For i = 1 To Len(Version1(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Version1(j - 1)), i, 1))), "13", "32")
                Stroka2 = Stroka2 + "_" + stroka3
            Next i
            stroka = "|" & stroka & "|" & Stroka1 & "|" & Stroka2 & "|"
            PrintLine(hFile1, stroka)

            stroka = ""
            For i = 1 To Len(Word2(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Word2(j - 1)), i, 1))), "13", "32")
                stroka = stroka + "_" + stroka3
            Next i
            Stroka1 = ""
            For i = 1 To Len(PhoneticNotation2(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(PhoneticNotation2(j - 1)), i, 1))), "13", "32")
                Stroka1 = Stroka1 + "_" + stroka3
            Next i
            Stroka2 = ""
            For i = 1 To Len(Version2(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Version2(j - 1)), i, 1))), "13", "32")
                Stroka2 = Stroka2 + "_" + stroka3
            Next i
            stroka = "|" & stroka & "|" & Stroka1 & "|" & Stroka2 & "|"
            PrintLine(hFile2, stroka)

            stroka = ""
            For i = 1 To Len(Value1(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Value1(j - 1)), i, 1))), "13", "32")
                stroka = stroka + "_" + stroka3
            Next i
            Stroka1 = ""
            For i = 1 To Len(Value2(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Value2(j - 1)), i, 1))), "13", "32")
                Stroka1 = Stroka1 + "_" + stroka3
            Next i
            Stroka2 = ""
            For i = 1 To Len(Value3(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Value3(j - 1)), i, 1))), "13", "32")
                Stroka2 = Stroka2 + "_" + stroka3
            Next i
            stroka = "|" & stroka & "|" & Stroka1 & "|" & Stroka2 & "|"
            PrintLine(hFile3, stroka)

            stroka = ""
            For i = 1 To Len(Value4(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Value4(j - 1)), i, 1))), "13", "32")
                stroka = stroka + "_" + stroka3
            Next i
            Stroka1 = ""
            For i = 1 To Len(Value5(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Value5(j - 1)), i, 1))), "13", "32")
                Stroka1 = Stroka1 + "_" + stroka3
            Next i
            Stroka2 = ""
            For i = 1 To Len(Value6(j - 1))
                stroka3 = Replace(System.Convert.ToString(AscW(Mid(Trim(Value6(j - 1)), i, 1))), "13", "32")
                Stroka2 = Stroka2 + "_" + stroka3
            Next i
            stroka = "|" & stroka & "|" & Stroka1 & "|" & Stroka2 & "|"
            PrintLine(hFile4, stroka)
        Next j
        FileClose(hFile1)
        FileClose(hFile2)
        FileClose(hFile3)
        FileClose(hFile4)

        LoadDataRichTextBox()
        PopulateDataGridView()

        Me.RichTextBox1.Text = "" : Me.RichTextBox2.Text = "" : Me.RichTextBox3.Text = "" : Me.RichTextBox4.Text = "" : Me.RichTextBox5.Text = ""
        Me.RichTextBox6.Text = "" : Me.RichTextBox7.Text = "" : Me.RichTextBox8.Text = ""
        EnabledEditButtons(False)
        EnabledEditButtons2(False)
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        EnabledEditButtons3(False)
    End Sub

    'Обращаемся к DataGridView
    Private Sub songsDataGridView_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles songsDataGridView.CellMouseClick
        'Dim messageBoxVB As New System.Text.StringBuilder()
        'messageBoxVB.AppendFormat("{0}={1}", "ColumnIndex", e.ColumnIndex) 'получаем номер колонки index-1
        'messageBoxVB.AppendLine()
        'messageBoxVB.AppendFormat("{0}={1}", "RowIndex", e.RowIndex) 'получаем номер строки index-1
        'messageBoxVB.AppendLine()
        'messageBoxVB.AppendFormat("{0}={1}", "button", e.Button) 'скорее всего получаем нажатую клавишу
        'messageBoxVB.AppendLine()
        ' messageBoxVB.AppendFormat("{0}={1}", "Click", e.Clicks) '
        ' messageBoxVB.AppendLine()
        ' messageBoxVB.AppendFormat("{0}={1}", "X", e.X) '
        'messageBoxVB.AppendLine()
        'messageBoxVB.AppendFormat("{0}={1}", "Y", e.Y) '
        'messageBoxVB.AppendLine()
        'messageBoxVB.AppendFormat("{0}={1}", "Delta", e.Delta) '
        'messageBoxVB.AppendLine()
        'messageBoxVB.AppendFormat("{0}={1}", "Location", e.Location) '
        'messageBoxVB.AppendLine()
        'MessageBox.Show(messageBoxVB.ToString(), "CellMouseClick Event")

        Dim i As Integer

        EnabledEditButtons(True)

        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox4.Text = ""
        RichTextBox5.Text = ""
        RichTextBox3.Text = ""
        RichTextBox7.Text = ""
        RichTextBox8.Text = ""
        RichTextBox6.Text = ""
        RichTextBox9.Text = ""
        RichTextBox10.Text = ""
        TextBox5.Text = ""
        'Предусмотреть 
        '     для
        '    hFile4

        i = e.RowIndex + 1
        SelectIndex = i - 1

        If i > 0 Then
            RichTextBox1.Text = Value2(i - 1)
            RichTextBox2.Text = Value3(i - 1)
            RichTextBox4.Text = PhoneticNotation1(i - 1)
            RichTextBox5.Text = Version1(i - 1)
            RichTextBox3.Text = Word1(i - 1)
            RichTextBox7.Text = PhoneticNotation2(i - 1)
            RichTextBox8.Text = Version2(i - 1)
            RichTextBox6.Text = Word2(i - 1)
            RichTextBox9.Text = Value4(i - 1)
            RichTextBox10.Text = Value5(i - 1)
            TextBox5.Text = Value6(i - 1)
            'Предусмотреть 
            '     для
            '    hFile4

            Select Case Value1(i - 1) 'раскраска кнопок типа копирования
                Case 1 : Button16.BackColor = Color.Red
                Case 2 : Button17.BackColor = Color.Red
                Case 3 : Button18.BackColor = Color.Red
                Case 4 : Button19.BackColor = Color.Red
                Case 5 : Button20.BackColor = Color.Red
                Case 6 : Button21.BackColor = Color.Red
            End Select

            Label10.Text = i.ToString
        End If
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Button16.BackColor = Color.Red
        Button17.BackColor = Me.GroupBox1.BackColor
        Button18.BackColor = Me.GroupBox1.BackColor
        Button19.BackColor = Me.GroupBox1.BackColor
        Button20.BackColor = Me.GroupBox1.BackColor
        Button21.BackColor = Me.GroupBox1.BackColor
    End Sub
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Button16.BackColor = Me.GroupBox1.BackColor
        Button17.BackColor = Color.Red
        Button18.BackColor = Me.GroupBox1.BackColor
        Button19.BackColor = Me.GroupBox1.BackColor
        Button20.BackColor = Me.GroupBox1.BackColor
        Button21.BackColor = Me.GroupBox1.BackColor
    End Sub
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Button16.BackColor = Me.GroupBox1.BackColor
        Button17.BackColor = Me.GroupBox1.BackColor
        Button18.BackColor = Color.Red
        Button19.BackColor = Me.GroupBox1.BackColor
        Button20.BackColor = Me.GroupBox1.BackColor
        Button21.BackColor = Me.GroupBox1.BackColor
    End Sub
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Button16.BackColor = Me.GroupBox1.BackColor
        Button17.BackColor = Me.GroupBox1.BackColor
        Button18.BackColor = Me.GroupBox1.BackColor
        Button19.BackColor = Color.Red
        Button20.BackColor = Me.GroupBox1.BackColor
        Button21.BackColor = Me.GroupBox1.BackColor
    End Sub
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Button16.BackColor = Me.GroupBox1.BackColor
        Button17.BackColor = Me.GroupBox1.BackColor
        Button18.BackColor = Me.GroupBox1.BackColor
        Button19.BackColor = Me.GroupBox1.BackColor
        Button20.BackColor = Color.Red
        Button21.BackColor = Me.GroupBox1.BackColor
    End Sub
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Button16.BackColor = Me.GroupBox1.BackColor
        Button17.BackColor = Me.GroupBox1.BackColor
        Button18.BackColor = Me.GroupBox1.BackColor
        Button19.BackColor = Me.GroupBox1.BackColor
        Button20.BackColor = Me.GroupBox1.BackColor
        Button21.BackColor = Color.Red
    End Sub

    'закрыть файлы источники
    Private Sub Button24_Click(sender As System.Object, e As System.EventArgs) Handles Button24.Click
        Dim oBook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim i As Integer
        Button24.Enabled = False
        On Error Resume Next
        For i = LBound(UniqueBook1) To UBound(UniqueBook1)
            oBook1 = UniqueBook1(i)
            oBook1.Close()
        Next
    End Sub

    Private Sub Button25_Click(sender As System.Object, e As System.EventArgs) Handles Button25.Click
        If GB4_H = 80 Then
            GB4_H = 20
            VisibleGB4(False)
            Button25.Text = "□"
            GoTo m4
        ElseIf GB4_H = 20 Then
            GB4_H = 80
            VisibleGB4(True)
            Button25.Text = "-"
            GoTo m4
        End If
m4:
        Me.Width = Me.Width + 1 : Me.Height = Me.Height + 1
        Me.Width = Me.Width - 1 : Me.Height = Me.Height - 1
    End Sub

    Private Sub Button26_Click(sender As System.Object, e As System.EventArgs) Handles Button26.Click
        If GB2_H = 108 Then
            GB2_H = 20
            VisibleGB2(False)
            Button26.Text = "□"
            GoTo m5
        ElseIf GB2_H = 20 Then
            GB2_H = 108
            VisibleGB2(True)
            Button26.Text = "-"
            GoTo m5
        End If
m5:
        Me.Width = Me.Width + 1 : Me.Height = Me.Height + 1
        Me.Width = Me.Width - 1 : Me.Height = Me.Height - 1
    End Sub

    Private Sub Button27_Click(sender As System.Object, e As System.EventArgs) Handles Button27.Click
        If GB1_H = 234 Then
            GB1_H = 20
            VisibleGB1(False)
            Button27.Text = "□"
            GoTo m6
        ElseIf GB1_H = 20 Then
            GB1_H = 234
            VisibleGB1(True)
            Button27.Text = "-"
            GoTo m6
        End If
m6:
        Me.Width = Me.Width + 1 : Me.Height = Me.Height + 1
        Me.Width = Me.Width - 1 : Me.Height = Me.Height - 1
    End Sub
    'Public GB4_H As Integer = 80
    'Public GB2_H As Integer = 108
    'Public GB1_H As Integer = 234
    'Public sDGC_H As Integer = 311
End Class