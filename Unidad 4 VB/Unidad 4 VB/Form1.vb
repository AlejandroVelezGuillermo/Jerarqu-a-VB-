Imports Microsoft.Win32
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Security.Policy
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Namespace CSV
    Partial Public Class Form1
        Inherits Form

        Private registros As List(Of registros) = New List(Of registros)()
        Private rutaArchivoActual As String = ""
        Private formato As String

        Public Sub New()
            InitializeComponent()
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        End Sub

        Private Sub aGREGARToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
            AGREGAR()
        End Sub

        Private Sub AGREGAR()
            If String.IsNullOrWhiteSpace(txtNombre.Text) OrElse String.IsNullOrWhiteSpace(txtTelefono.Text) OrElse String.IsNullOrWhiteSpace(txtCorreo.Text) OrElse String.IsNullOrWhiteSpace(comboBox2.Text) Then
                MessageBox.Show("Por favor, complete todos los campos antes de agregar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return
            End If

            Dim nuevoRegistro As registros = New registros With {
                .Nombre = txtNombre.Text,
                .Telefono = txtTelefono.Text,
                .Correo = txtCorreo.Text,
                .Asistencia = comboBox2.Text
            }
            registros.Add(nuevoRegistro)
            dgvDatos.DataSource = Nothing
            dgvDatos.DataSource = registros
            LimpiarCampos()
        End Sub

        Private Sub gUARDARToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
            SAVE()
        End Sub

        Private Sub SAVE()
            Dim NombreA As String = textBox1.Text

            Try
                Dim escritorio As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                Dim rutaArchivo As String = Path.Combine(escritorio, NombreA & "." & formato)

                Using writer As StreamWriter = New StreamWriter(rutaArchivo)
                    writer.WriteLine("Nombre,Telefono,Correo,Asistencia")

                    For Each registro As registros In registros
                        writer.WriteLine($"{registro.Nombre},{registro.Telefono},{registro.Correo}")
                    Next
                End Using

                MessageBox.Show($"Datos guardados exitosamente en el archivo CSV en el escritorio ({rutaArchivo}).", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                LimpiarCampos()
            Catch ex As Exception
                MessageBox.Show($"Error al guardar en el archivo CSV: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End Try

            dgvDatos.DataSource = Nothing
            dgvDatos.Rows.Clear()
            registros.Clear()
        End Sub

        Private Sub LimpiarCampos()
            LIMPIAR()
        End Sub

        Private Sub LIMPIAR()
            txtNombre.Text = ""
            txtTelefono.Text = ""
            txtCorreo.Text = ""
            textBox1.Text = ""
            comboBox2.Text = ""
        End Sub

        Private Sub rEMPLACARToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        End Sub

        Private Sub ActualizarGrafico()
            chart1.Series.Clear()
            Dim seriesAsistencia As Series = New Series("Asistencia")
            Dim conteoAsistencias = registros.GroupBy(Function(r) r.Asistencia).[Select](Function(g) New With {Key
                .Asistencia = g.Key, Key
                .Cantidad = g.Count()
            })

            For Each item In conteoAsistencias
                seriesAsistencia.Points.AddXY(item.Asistencia, item.Cantidad)
            Next

            chart1.Series.Add(seriesAsistencia)
            chart1.ChartAreas(0).AxisX.Title = "Asistencia"
            chart1.ChartAreas(0).AxisY.Title = "Cantidad"
            chart1.Series("Asistencia").ChartType = SeriesChartType.Column
        End Sub

        Private Function GenerarNumeroTelefono() As String
            Dim random As Random = New Random()
            Dim parte1 As String = random.[Next](100, 1000).ToString("000")
            Dim parte2 As String = random.[Next](100, 1000).ToString("000")
            Dim parte3 As String = random.[Next](1000, 10000).ToString("0000")
            Dim numeroCompleto As String = $"({parte1}) {parte2}-{parte3}"
            Return numeroCompleto
        End Function

        Private Function GenerarCorreoElectronico() As String
            Dim random As Random = New Random()
            Dim dominios As String() = {"gmail.com", "yahoo.com", "outlook.com", "example.com", "domain.com"}
            Dim parteInicial As String = Guid.NewGuid().ToString().Substring(0, 8)
            Dim dominio As String = dominios(random.[Next](dominios.Length))
            Dim correoCompleto As String = $"{parteInicial}@{dominio}"
            Return correoCompleto
        End Function

        Private Sub button4_Click(ByVal sender As Object, ByVal e As EventArgs)
            ActualizarGrafico()
        End Sub

        Private Sub InitializeComponent()
            Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
            Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
            Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
            Me.button4 = New System.Windows.Forms.Button()
            Me.button3 = New System.Windows.Forms.Button()
            Me.button2 = New System.Windows.Forms.Button()
            Me.button1 = New System.Windows.Forms.Button()
            Me.comboBox2 = New System.Windows.Forms.ComboBox()
            Me.label6 = New System.Windows.Forms.Label()
            Me.label5 = New System.Windows.Forms.Label()
            Me.comboBox1 = New System.Windows.Forms.ComboBox()
            Me.label4 = New System.Windows.Forms.Label()
            Me.textBox1 = New System.Windows.Forms.TextBox()
            Me.label3 = New System.Windows.Forms.Label()
            Me.label2 = New System.Windows.Forms.Label()
            Me.label1 = New System.Windows.Forms.Label()
            Me.txtCorreo = New System.Windows.Forms.TextBox()
            Me.txtTelefono = New System.Windows.Forms.TextBox()
            Me.txtNombre = New System.Windows.Forms.TextBox()
            Me.dgvDatos = New System.Windows.Forms.DataGridView()
            Me.menuStrip1 = New System.Windows.Forms.MenuStrip()
            Me.aGREGARToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.gUARDARToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.aBRIRToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.eDITARToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
            CType(Me.dgvDatos, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.menuStrip1.SuspendLayout()
            CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'button4
            '
            Me.button4.Location = New System.Drawing.Point(812, 444)
            Me.button4.Name = "button4"
            Me.button4.Size = New System.Drawing.Size(767, 23)
            Me.button4.TabIndex = 38
            Me.button4.Text = "ACTUALIZAR"
            Me.button4.UseVisualStyleBackColor = True
            '
            'button3
            '
            Me.button3.Location = New System.Drawing.Point(355, 94)
            Me.button3.Name = "button3"
            Me.button3.Size = New System.Drawing.Size(75, 23)
            Me.button3.TabIndex = 37
            Me.button3.Text = "Random"
            Me.button3.UseVisualStyleBackColor = True
            '
            'button2
            '
            Me.button2.Location = New System.Drawing.Point(355, 66)
            Me.button2.Name = "button2"
            Me.button2.Size = New System.Drawing.Size(75, 23)
            Me.button2.TabIndex = 36
            Me.button2.Text = "Random"
            Me.button2.UseVisualStyleBackColor = True
            '
            'button1
            '
            Me.button1.Location = New System.Drawing.Point(355, 38)
            Me.button1.Name = "button1"
            Me.button1.Size = New System.Drawing.Size(75, 23)
            Me.button1.TabIndex = 35
            Me.button1.Text = "Random"
            Me.button1.UseVisualStyleBackColor = True
            '
            'comboBox2
            '
            Me.comboBox2.FormattingEnabled = True
            Me.comboBox2.Items.AddRange(New Object() {"Asistio", "No Asistio"})
            Me.comboBox2.Location = New System.Drawing.Point(88, 126)
            Me.comboBox2.Name = "comboBox2"
            Me.comboBox2.Size = New System.Drawing.Size(102, 24)
            Me.comboBox2.TabIndex = 34
            '
            'label6
            '
            Me.label6.AutoSize = True
            Me.label6.Location = New System.Drawing.Point(10, 129)
            Me.label6.Name = "label6"
            Me.label6.Size = New System.Drawing.Size(72, 16)
            Me.label6.TabIndex = 33
            Me.label6.Text = "Asistencia:"
            '
            'label5
            '
            Me.label5.AutoSize = True
            Me.label5.Location = New System.Drawing.Point(510, 67)
            Me.label5.Name = "label5"
            Me.label5.Size = New System.Drawing.Size(110, 16)
            Me.label5.TabIndex = 32
            Me.label5.Text = "Ruta Del Archivo:"
            '
            'comboBox1
            '
            Me.comboBox1.FormattingEnabled = True
            Me.comboBox1.Items.AddRange(New Object() {"csv", "txt", "xml", "json"})
            Me.comboBox1.Location = New System.Drawing.Point(631, 59)
            Me.comboBox1.Name = "comboBox1"
            Me.comboBox1.Size = New System.Drawing.Size(56, 24)
            Me.comboBox1.TabIndex = 31
            '
            'label4
            '
            Me.label4.AutoSize = True
            Me.label4.Location = New System.Drawing.Point(489, 34)
            Me.label4.Name = "label4"
            Me.label4.Size = New System.Drawing.Size(131, 16)
            Me.label4.TabIndex = 30
            Me.label4.Text = "Nombre Del Archivo:"
            '
            'textBox1
            '
            Me.textBox1.Location = New System.Drawing.Point(631, 31)
            Me.textBox1.Name = "textBox1"
            Me.textBox1.Size = New System.Drawing.Size(133, 22)
            Me.textBox1.TabIndex = 29
            '
            'label3
            '
            Me.label3.AutoSize = True
            Me.label3.Location = New System.Drawing.Point(12, 100)
            Me.label3.Name = "label3"
            Me.label3.Size = New System.Drawing.Size(51, 16)
            Me.label3.TabIndex = 28
            Me.label3.Text = "Correo:"
            '
            'label2
            '
            Me.label2.AutoSize = True
            Me.label2.Location = New System.Drawing.Point(12, 72)
            Me.label2.Name = "label2"
            Me.label2.Size = New System.Drawing.Size(64, 16)
            Me.label2.TabIndex = 27
            Me.label2.Text = "Telefono:"
            '
            'label1
            '
            Me.label1.AutoSize = True
            Me.label1.Location = New System.Drawing.Point(12, 41)
            Me.label1.Name = "label1"
            Me.label1.Size = New System.Drawing.Size(59, 16)
            Me.label1.TabIndex = 26
            Me.label1.Text = "Nombre:"
            '
            'txtCorreo
            '
            Me.txtCorreo.Location = New System.Drawing.Point(77, 94)
            Me.txtCorreo.Name = "txtCorreo"
            Me.txtCorreo.Size = New System.Drawing.Size(272, 22)
            Me.txtCorreo.TabIndex = 25
            '
            'txtTelefono
            '
            Me.txtTelefono.Location = New System.Drawing.Point(77, 66)
            Me.txtTelefono.Name = "txtTelefono"
            Me.txtTelefono.Size = New System.Drawing.Size(272, 22)
            Me.txtTelefono.TabIndex = 24
            '
            'txtNombre
            '
            Me.txtNombre.Location = New System.Drawing.Point(77, 38)
            Me.txtNombre.Name = "txtNombre"
            Me.txtNombre.Size = New System.Drawing.Size(272, 22)
            Me.txtNombre.TabIndex = 23
            '
            'dgvDatos
            '
            Me.dgvDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvDatos.Location = New System.Drawing.Point(13, 164)
            Me.dgvDatos.Name = "dgvDatos"
            Me.dgvDatos.RowHeadersWidth = 51
            Me.dgvDatos.RowTemplate.Height = 24
            Me.dgvDatos.Size = New System.Drawing.Size(775, 294)
            Me.dgvDatos.TabIndex = 22
            '
            'menuStrip1
            '
            Me.menuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
            Me.menuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.aGREGARToolStripMenuItem, Me.gUARDARToolStripMenuItem, Me.aBRIRToolStripMenuItem, Me.eDITARToolStripMenuItem})
            Me.menuStrip1.Location = New System.Drawing.Point(0, 0)
            Me.menuStrip1.Name = "menuStrip1"
            Me.menuStrip1.Size = New System.Drawing.Size(1674, 28)
            Me.menuStrip1.TabIndex = 21
            Me.menuStrip1.Text = "menuStrip1"
            '
            'aGREGARToolStripMenuItem
            '
            Me.aGREGARToolStripMenuItem.Name = "aGREGARToolStripMenuItem"
            Me.aGREGARToolStripMenuItem.Size = New System.Drawing.Size(89, 24)
            Me.aGREGARToolStripMenuItem.Text = "AGREGAR"
            '
            'gUARDARToolStripMenuItem
            '
            Me.gUARDARToolStripMenuItem.Name = "gUARDARToolStripMenuItem"
            Me.gUARDARToolStripMenuItem.Size = New System.Drawing.Size(92, 24)
            Me.gUARDARToolStripMenuItem.Text = "GUARDAR"
            '
            'aBRIRToolStripMenuItem
            '
            Me.aBRIRToolStripMenuItem.Name = "aBRIRToolStripMenuItem"
            Me.aBRIRToolStripMenuItem.Size = New System.Drawing.Size(64, 24)
            Me.aBRIRToolStripMenuItem.Text = "ABRIR"
            '
            'eDITARToolStripMenuItem
            '
            Me.eDITARToolStripMenuItem.Name = "eDITARToolStripMenuItem"
            Me.eDITARToolStripMenuItem.Size = New System.Drawing.Size(72, 24)
            Me.eDITARToolStripMenuItem.Text = "EDITAR"
            '
            'Chart1
            '
            ChartArea1.Name = "ChartArea1"
            Me.Chart1.ChartAreas.Add(ChartArea1)
            Legend1.Name = "Legend1"
            Me.Chart1.Legends.Add(Legend1)
            Me.Chart1.Location = New System.Drawing.Point(812, 24)
            Me.Chart1.Name = "Chart1"
            Series1.ChartArea = "ChartArea1"
            Series1.Legend = "Legend1"
            Series1.Name = "Series1"
            Me.Chart1.Series.Add(Series1)
            Me.Chart1.Size = New System.Drawing.Size(767, 414)
            Me.Chart1.TabIndex = 39
            Me.Chart1.Text = "Chart1"
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(1674, 833)
            Me.Controls.Add(Me.Chart1)
            Me.Controls.Add(Me.button4)
            Me.Controls.Add(Me.button3)
            Me.Controls.Add(Me.button2)
            Me.Controls.Add(Me.button1)
            Me.Controls.Add(Me.comboBox2)
            Me.Controls.Add(Me.label6)
            Me.Controls.Add(Me.label5)
            Me.Controls.Add(Me.comboBox1)
            Me.Controls.Add(Me.label4)
            Me.Controls.Add(Me.textBox1)
            Me.Controls.Add(Me.label3)
            Me.Controls.Add(Me.label2)
            Me.Controls.Add(Me.label1)
            Me.Controls.Add(Me.txtCorreo)
            Me.Controls.Add(Me.txtTelefono)
            Me.Controls.Add(Me.txtNombre)
            Me.Controls.Add(Me.dgvDatos)
            Me.Controls.Add(Me.menuStrip1)
            Me.Name = "Form1"
            Me.Text = "Form1"
            CType(Me.dgvDatos, System.ComponentModel.ISupportInitialize).EndInit()
            Me.menuStrip1.ResumeLayout(False)
            Me.menuStrip1.PerformLayout()
            CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Private WithEvents button4 As Windows.Forms.Button
        Private WithEvents button3 As Windows.Forms.Button
        Private WithEvents button2 As Windows.Forms.Button
        Private WithEvents button1 As Windows.Forms.Button
        Private WithEvents comboBox2 As Windows.Forms.ComboBox
        Private WithEvents label6 As Label
        Private WithEvents label5 As Label
        Private WithEvents comboBox1 As Windows.Forms.ComboBox
        Private WithEvents label4 As Label
        Private WithEvents textBox1 As Windows.Forms.TextBox
        Private WithEvents label3 As Label
        Private WithEvents label2 As Label
        Private WithEvents label1 As Label
        Private WithEvents txtCorreo As Windows.Forms.TextBox
        Private WithEvents txtTelefono As Windows.Forms.TextBox
        Private WithEvents txtNombre As Windows.Forms.TextBox
        Private WithEvents dgvDatos As DataGridView
        Private WithEvents menuStrip1 As MenuStrip
        Private WithEvents aGREGARToolStripMenuItem As ToolStripMenuItem
        Private WithEvents gUARDARToolStripMenuItem As ToolStripMenuItem
        Private WithEvents aBRIRToolStripMenuItem As ToolStripMenuItem
        Private WithEvents eDITARToolStripMenuItem As ToolStripMenuItem

        Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        End Sub

        Friend WithEvents Chart1 As Chart
    End Class
End Namespace

