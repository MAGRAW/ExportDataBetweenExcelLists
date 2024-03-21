<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ФайлToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИзменитьШаблонToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЭкспортироватьШаблонToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИмпортироватьШаблонToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВыходToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СервисToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВидToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВернутьСтандартныйРазмерФормыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОПрограммеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.RichTextBox3 = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox4 = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox5 = New System.Windows.Forms.RichTextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button27 = New System.Windows.Forms.Button()
        Me.Button23 = New System.Windows.Forms.Button()
        Me.Button22 = New System.Windows.Forms.Button()
        Me.RichTextBox10 = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox9 = New System.Windows.Forms.RichTextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Button21 = New System.Windows.Forms.Button()
        Me.Button20 = New System.Windows.Forms.Button()
        Me.Button19 = New System.Windows.Forms.Button()
        Me.Button18 = New System.Windows.Forms.Button()
        Me.Button17 = New System.Windows.Forms.Button()
        Me.Button16 = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.RichTextBox8 = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox7 = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox6 = New System.Windows.Forms.RichTextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button26 = New System.Windows.Forms.Button()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Button25 = New System.Windows.Forms.Button()
        Me.Button15 = New System.Windows.Forms.Button()
        Me.Button14 = New System.Windows.Forms.Button()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Button12 = New System.Windows.Forms.Button()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Button24 = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.AllowDrop = True
        Me.TextBox1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.TextBox1.Location = New System.Drawing.Point(9, 66)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(776, 26)
        Me.TextBox1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 42)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(621, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Адрес файла Excel, содержащий диапозон данных для экспорта (копирования):"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 102)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(577, 20)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Адрес файла Excel, куда будет импортирован диапазон данных (вставка):"
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.Color.LightBlue
        Me.TextBox2.Location = New System.Drawing.Point(9, 126)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(776, 26)
        Me.TextBox2.TabIndex = 3
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button1.Location = New System.Drawing.Point(796, 65)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(34, 35)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "..."
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button2.Location = New System.Drawing.Point(796, 125)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(34, 35)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.BackColor = System.Drawing.Color.Gray
        Me.Button3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Button3.ForeColor = System.Drawing.Color.SpringGreen
        Me.Button3.Location = New System.Drawing.Point(418, 1252)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(196, 35)
        Me.Button3.TabIndex = 10
        Me.Button3.Text = "ПЕРЕНЕСТИ ДАННЫЕ"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.MenuStrip1.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 2)
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ФайлToolStripMenuItem, Me.СервисToolStripMenuItem, Me.ВидToolStripMenuItem, Me.ОПрограммеToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(876, 35)
        Me.MenuStrip1.TabIndex = 11
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ФайлToolStripMenuItem
        '
        Me.ФайлToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ИзменитьШаблонToolStripMenuItem, Me.ЭкспортироватьШаблонToolStripMenuItem, Me.ИмпортироватьШаблонToolStripMenuItem, Me.ВыходToolStripMenuItem})
        Me.ФайлToolStripMenuItem.Name = "ФайлToolStripMenuItem"
        Me.ФайлToolStripMenuItem.Size = New System.Drawing.Size(69, 29)
        Me.ФайлToolStripMenuItem.Text = "Файл"
        '
        'ИзменитьШаблонToolStripMenuItem
        '
        Me.ИзменитьШаблонToolStripMenuItem.Name = "ИзменитьШаблонToolStripMenuItem"
        Me.ИзменитьШаблонToolStripMenuItem.Size = New System.Drawing.Size(380, 34)
        Me.ИзменитьШаблонToolStripMenuItem.Text = "Выбрать шаблон"
        '
        'ЭкспортироватьШаблонToolStripMenuItem
        '
        Me.ЭкспортироватьШаблонToolStripMenuItem.Name = "ЭкспортироватьШаблонToolStripMenuItem"
        Me.ЭкспортироватьШаблонToolStripMenuItem.Size = New System.Drawing.Size(380, 34)
        Me.ЭкспортироватьШаблонToolStripMenuItem.Text = "Экспортировать шаблон в Excel"
        '
        'ИмпортироватьШаблонToolStripMenuItem
        '
        Me.ИмпортироватьШаблонToolStripMenuItem.Name = "ИмпортироватьШаблонToolStripMenuItem"
        Me.ИмпортироватьШаблонToolStripMenuItem.Size = New System.Drawing.Size(380, 34)
        Me.ИмпортироватьШаблонToolStripMenuItem.Text = "Импортировать шаблон из Excel"
        '
        'ВыходToolStripMenuItem
        '
        Me.ВыходToolStripMenuItem.Name = "ВыходToolStripMenuItem"
        Me.ВыходToolStripMenuItem.Size = New System.Drawing.Size(380, 34)
        Me.ВыходToolStripMenuItem.Text = "Выход"
        '
        'СервисToolStripMenuItem
        '
        Me.СервисToolStripMenuItem.Name = "СервисToolStripMenuItem"
        Me.СервисToolStripMenuItem.Size = New System.Drawing.Size(87, 29)
        Me.СервисToolStripMenuItem.Text = "Сервис"
        '
        'ВидToolStripMenuItem
        '
        Me.ВидToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ВернутьСтандартныйРазмерФормыToolStripMenuItem})
        Me.ВидToolStripMenuItem.Name = "ВидToolStripMenuItem"
        Me.ВидToolStripMenuItem.Size = New System.Drawing.Size(58, 29)
        Me.ВидToolStripMenuItem.Text = "Вид"
        '
        'ВернутьСтандартныйРазмерФормыToolStripMenuItem
        '
        Me.ВернутьСтандартныйРазмерФормыToolStripMenuItem.Name = "ВернутьСтандартныйРазмерФормыToolStripMenuItem"
        Me.ВернутьСтандартныйРазмерФормыToolStripMenuItem.Size = New System.Drawing.Size(419, 34)
        Me.ВернутьСтандартныйРазмерФормыToolStripMenuItem.Text = "Вернуть стандартный размер формы"
        '
        'ОПрограммеToolStripMenuItem
        '
        Me.ОПрограммеToolStripMenuItem.Name = "ОПрограммеToolStripMenuItem"
        Me.ОПрограммеToolStripMenuItem.Size = New System.Drawing.Size(141, 29)
        Me.ОПрограммеToolStripMenuItem.Text = "О программе"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(63, 31)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(89, 20)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Лист Excel"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(260, 31)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(83, 20)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Диапазон"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(584, 31)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(183, 20)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Комментарий к ссылке"
        '
        'RichTextBox3
        '
        Me.RichTextBox3.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.RichTextBox3.EnableAutoDragDrop = True
        Me.RichTextBox3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox3.Location = New System.Drawing.Point(9, 55)
        Me.RichTextBox3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox3.Name = "RichTextBox3"
        Me.RichTextBox3.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.RichTextBox3.Size = New System.Drawing.Size(206, 90)
        Me.RichTextBox3.TabIndex = 18
        Me.RichTextBox3.Text = ""
        '
        'RichTextBox4
        '
        Me.RichTextBox4.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.RichTextBox4.EnableAutoDragDrop = True
        Me.RichTextBox4.Enabled = False
        Me.RichTextBox4.Font = New System.Drawing.Font("Lucida Sans Unicode", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox4.ForeColor = System.Drawing.Color.Green
        Me.RichTextBox4.Location = New System.Drawing.Point(219, 55)
        Me.RichTextBox4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox4.Name = "RichTextBox4"
        Me.RichTextBox4.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.RichTextBox4.Size = New System.Drawing.Size(178, 90)
        Me.RichTextBox4.TabIndex = 19
        Me.RichTextBox4.Text = ""
        '
        'RichTextBox5
        '
        Me.RichTextBox5.BackColor = System.Drawing.Color.LightGreen
        Me.RichTextBox5.EnableAutoDragDrop = True
        Me.RichTextBox5.Enabled = False
        Me.RichTextBox5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox5.ForeColor = System.Drawing.Color.DimGray
        Me.RichTextBox5.Location = New System.Drawing.Point(586, 55)
        Me.RichTextBox5.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox5.Name = "RichTextBox5"
        Me.RichTextBox5.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.RichTextBox5.Size = New System.Drawing.Size(186, 124)
        Me.RichTextBox5.TabIndex = 20
        Me.RichTextBox5.Text = ""
        '
        'Button4
        '
        Me.Button4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Button4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Button4.Location = New System.Drawing.Point(746, 1252)
        Me.Button4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(112, 35)
        Me.Button4.TabIndex = 21
        Me.Button4.Text = "Выход"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button5.BackColor = System.Drawing.Color.DarkSlateGray
        Me.Button5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Button5.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Button5.Location = New System.Drawing.Point(624, 1252)
        Me.Button5.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(112, 35)
        Me.Button5.TabIndex = 22
        Me.Button5.Text = "Справка"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.Enabled = False
        Me.Button6.Location = New System.Drawing.Point(354, 317)
        Me.Button6.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(112, 35)
        Me.Button6.TabIndex = 23
        Me.Button6.Text = "Добавить"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Enabled = False
        Me.Button7.Location = New System.Drawing.Point(597, 317)
        Me.Button7.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(112, 35)
        Me.Button7.TabIndex = 24
        Me.Button7.Text = "Удалить"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Enabled = False
        Me.Button8.Location = New System.Drawing.Point(718, 317)
        Me.Button8.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(112, 35)
        Me.Button8.TabIndex = 25
        Me.Button8.Text = "Отменить"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Enabled = False
        Me.Button9.Location = New System.Drawing.Point(476, 317)
        Me.Button9.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(112, 35)
        Me.Button9.TabIndex = 26
        Me.Button9.Text = "Изменить"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.MediumSeaGreen
        Me.GroupBox1.Controls.Add(Me.Button27)
        Me.GroupBox1.Controls.Add(Me.Button23)
        Me.GroupBox1.Controls.Add(Me.Button22)
        Me.GroupBox1.Controls.Add(Me.RichTextBox10)
        Me.GroupBox1.Controls.Add(Me.RichTextBox9)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.RichTextBox2)
        Me.GroupBox1.Controls.Add(Me.RichTextBox1)
        Me.GroupBox1.Controls.Add(Me.Button21)
        Me.GroupBox1.Controls.Add(Me.Button20)
        Me.GroupBox1.Controls.Add(Me.Button19)
        Me.GroupBox1.Controls.Add(Me.Button18)
        Me.GroupBox1.Controls.Add(Me.Button17)
        Me.GroupBox1.Controls.Add(Me.Button16)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.TextBox3)
        Me.GroupBox1.Controls.Add(Me.RichTextBox8)
        Me.GroupBox1.Controls.Add(Me.RichTextBox7)
        Me.GroupBox1.Controls.Add(Me.RichTextBox6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Button6)
        Me.GroupBox1.Controls.Add(Me.Button9)
        Me.GroupBox1.Controls.Add(Me.RichTextBox3)
        Me.GroupBox1.Controls.Add(Me.Button8)
        Me.GroupBox1.Controls.Add(Me.RichTextBox4)
        Me.GroupBox1.Controls.Add(Me.Button7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.RichTextBox5)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(18, 394)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Size = New System.Drawing.Size(840, 360)
        Me.GroupBox1.TabIndex = 27
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Редактирование ссылок"
        '
        'Button27
        '
        Me.Button27.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button27.Location = New System.Drawing.Point(747, 0)
        Me.Button27.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button27.Name = "Button27"
        Me.Button27.Size = New System.Drawing.Size(27, 28)
        Me.Button27.TabIndex = 47
        Me.Button27.Text = "-"
        Me.Button27.UseVisualStyleBackColor = True
        '
        'Button23
        '
        Me.Button23.Enabled = False
        Me.Button23.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button23.Location = New System.Drawing.Point(9, 278)
        Me.Button23.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button23.Name = "Button23"
        Me.Button23.Size = New System.Drawing.Size(34, 35)
        Me.Button23.TabIndex = 46
        Me.Button23.Text = "..."
        Me.Button23.UseVisualStyleBackColor = True
        '
        'Button22
        '
        Me.Button22.Enabled = False
        Me.Button22.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button22.Location = New System.Drawing.Point(9, 149)
        Me.Button22.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button22.Name = "Button22"
        Me.Button22.Size = New System.Drawing.Size(34, 35)
        Me.Button22.TabIndex = 34
        Me.Button22.Text = "..."
        Me.Button22.UseVisualStyleBackColor = True
        '
        'RichTextBox10
        '
        Me.RichTextBox10.BackColor = System.Drawing.Color.LightBlue
        Me.RichTextBox10.Location = New System.Drawing.Point(45, 278)
        Me.RichTextBox10.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox10.Name = "RichTextBox10"
        Me.RichTextBox10.Size = New System.Drawing.Size(352, 30)
        Me.RichTextBox10.TabIndex = 45
        Me.RichTextBox10.Text = ""
        '
        'RichTextBox9
        '
        Me.RichTextBox9.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.RichTextBox9.Location = New System.Drawing.Point(45, 149)
        Me.RichTextBox9.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox9.Name = "RichTextBox9"
        Me.RichTextBox9.Size = New System.Drawing.Size(352, 30)
        Me.RichTextBox9.TabIndex = 44
        Me.RichTextBox9.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(440, 158)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(79, 20)
        Me.Label4.TabIndex = 43
        Me.Label4.Text = "Формула"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(418, 31)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 20)
        Me.Label3.TabIndex = 42
        Me.Label3.Text = "Имя переменной"
        '
        'RichTextBox2
        '
        Me.RichTextBox2.BackColor = System.Drawing.SystemColors.HotTrack
        Me.RichTextBox2.Enabled = False
        Me.RichTextBox2.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox2.ForeColor = System.Drawing.SystemColors.Window
        Me.RichTextBox2.Location = New System.Drawing.Point(400, 183)
        Me.RichTextBox2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(182, 126)
        Me.RichTextBox2.TabIndex = 41
        Me.RichTextBox2.Text = ""
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.Color.Yellow
        Me.RichTextBox1.Enabled = False
        Me.RichTextBox1.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox1.ForeColor = System.Drawing.Color.Maroon
        Me.RichTextBox1.Location = New System.Drawing.Point(400, 55)
        Me.RichTextBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(182, 90)
        Me.RichTextBox1.TabIndex = 40
        Me.RichTextBox1.Text = ""
        '
        'Button21
        '
        Me.Button21.Enabled = False
        Me.Button21.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button21.Image = Global.WindowsApplication1.My.Resources.Resources._6
        Me.Button21.Location = New System.Drawing.Point(783, 262)
        Me.Button21.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button21.Name = "Button21"
        Me.Button21.Size = New System.Drawing.Size(45, 51)
        Me.Button21.TabIndex = 39
        Me.Button21.UseVisualStyleBackColor = True
        '
        'Button20
        '
        Me.Button20.Enabled = False
        Me.Button20.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button20.Image = Global.WindowsApplication1.My.Resources.Resources._5
        Me.Button20.Location = New System.Drawing.Point(783, 212)
        Me.Button20.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button20.Name = "Button20"
        Me.Button20.Size = New System.Drawing.Size(45, 51)
        Me.Button20.TabIndex = 38
        Me.Button20.UseVisualStyleBackColor = True
        '
        'Button19
        '
        Me.Button19.Enabled = False
        Me.Button19.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button19.Image = Global.WindowsApplication1.My.Resources.Resources._4
        Me.Button19.Location = New System.Drawing.Point(783, 163)
        Me.Button19.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button19.Name = "Button19"
        Me.Button19.Size = New System.Drawing.Size(45, 51)
        Me.Button19.TabIndex = 37
        Me.Button19.UseVisualStyleBackColor = True
        '
        'Button18
        '
        Me.Button18.Enabled = False
        Me.Button18.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button18.Image = Global.WindowsApplication1.My.Resources.Resources._3
        Me.Button18.Location = New System.Drawing.Point(783, 114)
        Me.Button18.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button18.Name = "Button18"
        Me.Button18.Size = New System.Drawing.Size(45, 51)
        Me.Button18.TabIndex = 36
        Me.Button18.UseVisualStyleBackColor = True
        '
        'Button17
        '
        Me.Button17.Enabled = False
        Me.Button17.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button17.Image = Global.WindowsApplication1.My.Resources.Resources._2
        Me.Button17.Location = New System.Drawing.Point(783, 65)
        Me.Button17.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(45, 51)
        Me.Button17.TabIndex = 35
        Me.Button17.UseVisualStyleBackColor = True
        '
        'Button16
        '
        Me.Button16.AccessibleDescription = ""
        Me.Button16.BackColor = System.Drawing.Color.MediumSeaGreen
        Me.Button16.Enabled = False
        Me.Button16.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button16.Image = Global.WindowsApplication1.My.Resources.Resources._11
        Me.Button16.Location = New System.Drawing.Point(783, 15)
        Me.Button16.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button16.Name = "Button16"
        Me.Button16.Size = New System.Drawing.Size(45, 51)
        Me.Button16.TabIndex = 34
        Me.Button16.Tag = ""
        Me.Button16.UseVisualStyleBackColor = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(124, 325)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(108, 20)
        Me.Label11.TabIndex = 33
        Me.Label11.Text = "заменить на:"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(80, 325)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(0, 20)
        Me.Label10.TabIndex = 32
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(4, 325)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(69, 20)
        Me.Label9.TabIndex = 31
        Me.Label9.Text = "Индекс:"
        '
        'TextBox3
        '
        Me.TextBox3.Enabled = False
        Me.TextBox3.Location = New System.Drawing.Point(244, 317)
        Me.TextBox3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(58, 26)
        Me.TextBox3.TabIndex = 28
        '
        'RichTextBox8
        '
        Me.RichTextBox8.BackColor = System.Drawing.Color.LightGreen
        Me.RichTextBox8.EnableAutoDragDrop = True
        Me.RichTextBox8.Enabled = False
        Me.RichTextBox8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox8.ForeColor = System.Drawing.Color.DimGray
        Me.RichTextBox8.Location = New System.Drawing.Point(586, 183)
        Me.RichTextBox8.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox8.Name = "RichTextBox8"
        Me.RichTextBox8.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.RichTextBox8.Size = New System.Drawing.Size(186, 126)
        Me.RichTextBox8.TabIndex = 29
        Me.RichTextBox8.Text = ""
        '
        'RichTextBox7
        '
        Me.RichTextBox7.BackColor = System.Drawing.Color.LightBlue
        Me.RichTextBox7.EnableAutoDragDrop = True
        Me.RichTextBox7.Enabled = False
        Me.RichTextBox7.Font = New System.Drawing.Font("Lucida Sans Unicode", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox7.ForeColor = System.Drawing.Color.Green
        Me.RichTextBox7.Location = New System.Drawing.Point(219, 183)
        Me.RichTextBox7.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox7.Name = "RichTextBox7"
        Me.RichTextBox7.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.RichTextBox7.Size = New System.Drawing.Size(178, 92)
        Me.RichTextBox7.TabIndex = 28
        Me.RichTextBox7.Text = ""
        '
        'RichTextBox6
        '
        Me.RichTextBox6.BackColor = System.Drawing.Color.LightBlue
        Me.RichTextBox6.EnableAutoDragDrop = True
        Me.RichTextBox6.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox6.Location = New System.Drawing.Point(9, 183)
        Me.RichTextBox6.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBox6.Name = "RichTextBox6"
        Me.RichTextBox6.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.RichTextBox6.Size = New System.Drawing.Size(206, 92)
        Me.RichTextBox6.TabIndex = 27
        Me.RichTextBox6.Text = ""
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.MediumSeaGreen
        Me.GroupBox2.Controls.Add(Me.Button26)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.TextBox2)
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.ForeColor = System.Drawing.Color.Black
        Me.GroupBox2.Location = New System.Drawing.Point(18, 203)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox2.Size = New System.Drawing.Size(840, 166)
        Me.GroupBox2.TabIndex = 28
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Адреса файлов"
        '
        'Button26
        '
        Me.Button26.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button26.Location = New System.Drawing.Point(804, 0)
        Me.Button26.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button26.Name = "Button26"
        Me.Button26.Size = New System.Drawing.Size(27, 28)
        Me.Button26.TabIndex = 6
        Me.Button26.Text = "-"
        Me.Button26.UseVisualStyleBackColor = True
        '
        'TextBox4
        '
        Me.TextBox4.BackColor = System.Drawing.SystemColors.ButtonShadow
        Me.TextBox4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox4.ForeColor = System.Drawing.Color.SpringGreen
        Me.TextBox4.Location = New System.Drawing.Point(84, 35)
        Me.TextBox4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(702, 27)
        Me.TextBox4.TabIndex = 29
        '
        'Button10
        '
        Me.Button10.AccessibleDescription = ""
        Me.Button10.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button10.Location = New System.Drawing.Point(796, 34)
        Me.Button10.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(34, 35)
        Me.Button10.TabIndex = 30
        Me.Button10.Tag = ""
        Me.Button10.Text = "..."
        Me.Button10.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 40)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 20)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Шаблон:"
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.MediumSeaGreen
        Me.GroupBox4.Controls.Add(Me.Button25)
        Me.GroupBox4.Controls.Add(Me.Button15)
        Me.GroupBox4.Controls.Add(Me.Button14)
        Me.GroupBox4.Controls.Add(Me.Button13)
        Me.GroupBox4.Controls.Add(Me.Button12)
        Me.GroupBox4.Controls.Add(Me.Button11)
        Me.GroupBox4.Controls.Add(Me.Label8)
        Me.GroupBox4.Controls.Add(Me.TextBox4)
        Me.GroupBox4.Controls.Add(Me.Button10)
        Me.GroupBox4.ForeColor = System.Drawing.Color.Black
        Me.GroupBox4.Location = New System.Drawing.Point(18, 55)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox4.Size = New System.Drawing.Size(840, 123)
        Me.GroupBox4.TabIndex = 33
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "    Редактор шаблонов"
        '
        'Button25
        '
        Me.Button25.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button25.Location = New System.Drawing.Point(804, 0)
        Me.Button25.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button25.Name = "Button25"
        Me.Button25.Size = New System.Drawing.Size(27, 28)
        Me.Button25.TabIndex = 37
        Me.Button25.Text = "-"
        Me.Button25.UseVisualStyleBackColor = True
        '
        'Button15
        '
        Me.Button15.Enabled = False
        Me.Button15.Location = New System.Drawing.Point(314, 78)
        Me.Button15.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button15.Name = "Button15"
        Me.Button15.Size = New System.Drawing.Size(122, 35)
        Me.Button15.TabIndex = 36
        Me.Button15.Text = "Клонировать"
        Me.Button15.UseVisualStyleBackColor = True
        '
        'Button14
        '
        Me.Button14.Enabled = False
        Me.Button14.Location = New System.Drawing.Point(718, 78)
        Me.Button14.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(112, 35)
        Me.Button14.TabIndex = 35
        Me.Button14.Text = "Отменить"
        Me.Button14.UseVisualStyleBackColor = True
        '
        'Button13
        '
        Me.Button13.Enabled = False
        Me.Button13.Location = New System.Drawing.Point(597, 78)
        Me.Button13.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(112, 35)
        Me.Button13.TabIndex = 34
        Me.Button13.Text = "Удалить"
        Me.Button13.UseVisualStyleBackColor = True
        '
        'Button12
        '
        Me.Button12.Enabled = False
        Me.Button12.Location = New System.Drawing.Point(444, 78)
        Me.Button12.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(144, 35)
        Me.Button12.TabIndex = 33
        Me.Button12.Text = "Переименовать"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'Button11
        '
        Me.Button11.Enabled = False
        Me.Button11.Location = New System.Drawing.Point(153, 78)
        Me.Button11.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(152, 35)
        Me.Button11.TabIndex = 32
        Me.Button11.Text = "Добавить новый"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'Button24
        '
        Me.Button24.BackColor = System.Drawing.Color.Coral
        Me.Button24.Enabled = False
        Me.Button24.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button24.ForeColor = System.Drawing.Color.Maroon
        Me.Button24.Location = New System.Drawing.Point(18, 1252)
        Me.Button24.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button24.Name = "Button24"
        Me.Button24.Size = New System.Drawing.Size(254, 35)
        Me.Button24.TabIndex = 34
        Me.Button24.Text = "Закрыть источники"
        Me.Button24.UseVisualStyleBackColor = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SeaGreen
        Me.ClientSize = Global.WindowsApplication1.My.MySettings.Default.размер_Form1
        Me.Controls.Add(Me.Button24)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.MenuStrip1)
        Me.DataBindings.Add(New System.Windows.Forms.Binding("Location", Global.WindowsApplication1.My.MySettings.Default, "позиция_Form1", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.DataBindings.Add(New System.Windows.Forms.Binding("ClientSize", Global.WindowsApplication1.My.MySettings.Default, "размер_Form1", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = Global.WindowsApplication1.My.MySettings.Default.позиция_Form1
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Form1"
        Me.Opacity = 0.97R
        Me.Text = "Перенос данных между листами Excel"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ФайлToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents СервисToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ОПрограммеToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents RichTextBox3 As System.Windows.Forms.RichTextBox
    Friend WithEvents RichTextBox4 As System.Windows.Forms.RichTextBox
    Friend WithEvents RichTextBox5 As System.Windows.Forms.RichTextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RichTextBox8 As System.Windows.Forms.RichTextBox
    Friend WithEvents RichTextBox7 As System.Windows.Forms.RichTextBox
    Friend WithEvents RichTextBox6 As System.Windows.Forms.RichTextBox
    Friend WithEvents ИзменитьШаблонToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Button13 As System.Windows.Forms.Button
    Friend WithEvents Button12 As System.Windows.Forms.Button
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Button15 As System.Windows.Forms.Button
    Friend WithEvents Button14 As System.Windows.Forms.Button
    Friend WithEvents ВыходToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВидToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВернутьСтандартныйРазмерФормыToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ЭкспортироватьШаблонToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ИмпортироватьШаблонToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button21 As System.Windows.Forms.Button
    Friend WithEvents Button20 As System.Windows.Forms.Button
    Friend WithEvents Button19 As System.Windows.Forms.Button
    Friend WithEvents Button18 As System.Windows.Forms.Button
    Friend WithEvents Button17 As System.Windows.Forms.Button
    Friend WithEvents Button16 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents RichTextBox2 As System.Windows.Forms.RichTextBox
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Button23 As System.Windows.Forms.Button
    Friend WithEvents Button22 As System.Windows.Forms.Button
    Friend WithEvents RichTextBox10 As System.Windows.Forms.RichTextBox
    Friend WithEvents RichTextBox9 As System.Windows.Forms.RichTextBox
    Friend WithEvents Button24 As System.Windows.Forms.Button
    Friend WithEvents Button25 As System.Windows.Forms.Button
    Friend WithEvents Button27 As System.Windows.Forms.Button
    Friend WithEvents Button26 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

End Class
