<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form工艺编写
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.工艺制作 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.工艺时间记录 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DWGInfo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CustDWG = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DrawNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Qty = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CustCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.工序号 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.资源名称 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.工序内容和注意事项 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.准备工时h = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.单件工时h = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.该工序预计单件成本 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_生产单号 = New System.Windows.Forms.TextBox()
        Me.ComboBox_未写工艺图号 = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_已选图号 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox_DWGInfo = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextBox_CustCode = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox_参考图号 = New System.Windows.Forms.TextBox()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label_Price = New System.Windows.Forms.Label()
        Me.Label_预计单件成本 = New System.Windows.Forms.Label()
        Me.Label_预计总成本 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.ComboBox_已写工艺图号 = New System.Windows.Forms.ComboBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.工艺制作, Me.工艺时间记录, Me.DWGInfo, Me.CustDWG, Me.DrawNo, Me.Qty, Me.CustCode})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 62)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(631, 158)
        Me.DataGridView1.TabIndex = 0
        '
        '工艺制作
        '
        Me.工艺制作.HeaderText = "工艺制作"
        Me.工艺制作.Name = "工艺制作"
        Me.工艺制作.Width = 80
        '
        '工艺时间记录
        '
        Me.工艺时间记录.HeaderText = "工艺时间记录"
        Me.工艺时间记录.Name = "工艺时间记录"
        '
        'DWGInfo
        '
        Me.DWGInfo.HeaderText = "DWGInfo"
        Me.DWGInfo.Name = "DWGInfo"
        '
        'CustDWG
        '
        Me.CustDWG.HeaderText = "CustDWG"
        Me.CustDWG.Name = "CustDWG"
        '
        'DrawNo
        '
        Me.DrawNo.HeaderText = "图号"
        Me.DrawNo.Name = "DrawNo"
        '
        'Qty
        '
        Me.Qty.HeaderText = "数量"
        Me.Qty.Name = "Qty"
        Me.Qty.Width = 60
        '
        'CustCode
        '
        Me.CustCode.HeaderText = "客户号"
        Me.CustCode.Name = "CustCode"
        Me.CustCode.Width = 70
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.工序号, Me.资源名称, Me.工序内容和注意事项, Me.准备工时h, Me.单件工时h, Me.该工序预计单件成本})
        Me.DataGridView2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView2.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.DataGridView2.Location = New System.Drawing.Point(12, 237)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowTemplate.Height = 23
        Me.DataGridView2.Size = New System.Drawing.Size(1273, 571)
        Me.DataGridView2.TabIndex = 1
        '
        '工序号
        '
        Me.工序号.HeaderText = "工序号"
        Me.工序号.Name = "工序号"
        '
        '资源名称
        '
        Me.资源名称.HeaderText = "资源名称"
        Me.资源名称.Name = "资源名称"
        '
        '工序内容和注意事项
        '
        Me.工序内容和注意事项.HeaderText = "工序内容和注意事项"
        Me.工序内容和注意事项.Name = "工序内容和注意事项"
        Me.工序内容和注意事项.Width = 700
        '
        '准备工时h
        '
        Me.准备工时h.HeaderText = "准备工时h"
        Me.准备工时h.Name = "准备工时h"
        '
        '单件工时h
        '
        Me.单件工时h.HeaderText = "单件工时h"
        Me.单件工时h.Name = "单件工时h"
        '
        '该工序预计单件成本
        '
        Me.该工序预计单件成本.HeaderText = "该工序预计单件成本"
        Me.该工序预计单件成本.Name = "该工序预计单件成本"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(414, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "选择生产单号："
        '
        'TextBox_生产单号
        '
        Me.TextBox_生产单号.Location = New System.Drawing.Point(528, 25)
        Me.TextBox_生产单号.Name = "TextBox_生产单号"
        Me.TextBox_生产单号.Size = New System.Drawing.Size(99, 21)
        Me.TextBox_生产单号.TabIndex = 3
        '
        'ComboBox_未写工艺图号
        '
        Me.ComboBox_未写工艺图号.FormattingEnabled = True
        Me.ComboBox_未写工艺图号.Location = New System.Drawing.Point(801, 12)
        Me.ComboBox_未写工艺图号.Name = "ComboBox_未写工艺图号"
        Me.ComboBox_未写工艺图号.Size = New System.Drawing.Size(291, 20)
        Me.ComboBox_未写工艺图号.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微软雅黑", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(660, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(135, 19)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "选择未写工艺图号："
        '
        'TextBox_已选图号
        '
        Me.TextBox_已选图号.Enabled = False
        Me.TextBox_已选图号.Location = New System.Drawing.Point(801, 88)
        Me.TextBox_已选图号.Name = "TextBox_已选图号"
        Me.TextBox_已选图号.Size = New System.Drawing.Size(291, 21)
        Me.TextBox_已选图号.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(16, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(168, 16)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "选择参考图号的工艺："
        '
        'TextBox3
        '
        Me.TextBox3.Enabled = False
        Me.TextBox3.Location = New System.Drawing.Point(801, 120)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(291, 21)
        Me.TextBox3.TabIndex = 9
        '
        'TextBox_DWGInfo
        '
        Me.TextBox_DWGInfo.Enabled = False
        Me.TextBox_DWGInfo.Location = New System.Drawing.Point(801, 149)
        Me.TextBox_DWGInfo.Name = "TextBox_DWGInfo"
        Me.TextBox_DWGInfo.Size = New System.Drawing.Size(291, 21)
        Me.TextBox_DWGInfo.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(694, 93)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "已选图号："
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label5.Location = New System.Drawing.Point(694, 125)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 16)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "客户图号："
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(691, 149)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(104, 16)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "图号全信息："
        '
        'TextBox_CustCode
        '
        Me.TextBox_CustCode.Enabled = False
        Me.TextBox_CustCode.Location = New System.Drawing.Point(1200, 11)
        Me.TextBox_CustCode.Name = "TextBox_CustCode"
        Me.TextBox_CustCode.Size = New System.Drawing.Size(79, 21)
        Me.TextBox_CustCode.TabIndex = 14
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(1200, 88)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(79, 21)
        Me.TextBox6.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(1107, 93)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "工艺数量："
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Red
        Me.Label8.Location = New System.Drawing.Point(1107, 54)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(104, 16)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "生产单数量："
        '
        'TextBox7
        '
        Me.TextBox7.Enabled = False
        Me.TextBox7.Location = New System.Drawing.Point(1200, 49)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(79, 21)
        Me.TextBox7.TabIndex = 18
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("宋体", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Red
        Me.Button1.Location = New System.Drawing.Point(1110, 125)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(169, 39)
        Me.Button1.TabIndex = 19
        Me.Button1.Text = "工艺确认"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox_参考图号
        '
        Me.TextBox_参考图号.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBox_参考图号.Location = New System.Drawing.Point(177, 26)
        Me.TextBox_参考图号.Name = "TextBox_参考图号"
        Me.TextBox_参考图号.Size = New System.Drawing.Size(208, 26)
        Me.TextBox_参考图号.TabIndex = 21
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button5.Location = New System.Drawing.Point(833, 199)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(123, 32)
        Me.Button5.TabIndex = 28
        Me.Button5.Text = "插入一行"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button4.Location = New System.Drawing.Point(1156, 199)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(123, 32)
        Me.Button4.TabIndex = 27
        Me.Button4.Text = "删除一行"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button3.Location = New System.Drawing.Point(1010, 199)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(123, 32)
        Me.Button3.TabIndex = 26
        Me.Button3.Text = "增加一行"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button2.Location = New System.Drawing.Point(674, 199)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(123, 32)
        Me.Button2.TabIndex = 25
        Me.Button2.Text = "产生空白工艺卡"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label_Price
        '
        Me.Label_Price.AutoSize = True
        Me.Label_Price.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label_Price.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label_Price.Location = New System.Drawing.Point(733, 178)
        Me.Label_Price.Name = "Label_Price"
        Me.Label_Price.Size = New System.Drawing.Size(101, 12)
        Me.Label_Price.TabIndex = 29
        Me.Label_Price.Text = "当前单件报价￥："
        '
        'Label_预计单件成本
        '
        Me.Label_预计单件成本.AutoSize = True
        Me.Label_预计单件成本.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label_预计单件成本.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label_预计单件成本.Location = New System.Drawing.Point(947, 178)
        Me.Label_预计单件成本.Name = "Label_预计单件成本"
        Me.Label_预计单件成本.Size = New System.Drawing.Size(95, 12)
        Me.Label_预计单件成本.TabIndex = 30
        Me.Label_预计单件成本.Text = "预计单件成本￥:"
        '
        'Label_预计总成本
        '
        Me.Label_预计总成本.AutoSize = True
        Me.Label_预计总成本.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label_预计总成本.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label_预计总成本.Location = New System.Drawing.Point(1134, 178)
        Me.Label_预计总成本.Name = "Label_预计总成本"
        Me.Label_预计总成本.Size = New System.Drawing.Size(83, 12)
        Me.Label_预计总成本.TabIndex = 31
        Me.Label_预计总成本.Text = "预计总成本￥:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("微软雅黑", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(660, 50)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(135, 19)
        Me.Label9.TabIndex = 33
        Me.Label9.Text = "选择已写工艺图号："
        '
        'ComboBox_已写工艺图号
        '
        Me.ComboBox_已写工艺图号.FormattingEnabled = True
        Me.ComboBox_已写工艺图号.Location = New System.Drawing.Point(801, 50)
        Me.ComboBox_已写工艺图号.Name = "ComboBox_已写工艺图号"
        Me.ComboBox_已写工艺图号.Size = New System.Drawing.Size(291, 20)
        Me.ComboBox_已写工艺图号.TabIndex = 32
        '
        'Form工艺编写
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1297, 820)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.ComboBox_已写工艺图号)
        Me.Controls.Add(Me.Label_预计总成本)
        Me.Controls.Add(Me.Label_预计单件成本)
        Me.Controls.Add(Me.Label_Price)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TextBox_参考图号)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox7)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.TextBox_CustCode)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox_DWGInfo)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox_已选图号)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ComboBox_未写工艺图号)
        Me.Controls.Add(Me.TextBox_生产单号)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.DataGridView1)
        Me.MaximizeBox = False
        Me.Name = "Form工艺编写"
        Me.Text = "赛峰生产单跟进系统--工艺编写"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_生产单号 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox_未写工艺图号 As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_已选图号 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_DWGInfo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox_CustCode As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox_参考图号 As System.Windows.Forms.TextBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label_Price As System.Windows.Forms.Label
    Friend WithEvents Label_预计单件成本 As System.Windows.Forms.Label
    Friend WithEvents Label_预计总成本 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_已写工艺图号 As ComboBox
    Friend WithEvents 工序号 As DataGridViewTextBoxColumn
    Friend WithEvents 资源名称 As DataGridViewComboBoxColumn
    Friend WithEvents 工序内容和注意事项 As DataGridViewTextBoxColumn
    Friend WithEvents 准备工时h As DataGridViewTextBoxColumn
    Friend WithEvents 单件工时h As DataGridViewTextBoxColumn
    Friend WithEvents 该工序预计单件成本 As DataGridViewTextBoxColumn
    Friend WithEvents 工艺制作 As DataGridViewTextBoxColumn
    Friend WithEvents 工艺时间记录 As DataGridViewTextBoxColumn
    Friend WithEvents DWGInfo As DataGridViewTextBoxColumn
    Friend WithEvents CustDWG As DataGridViewTextBoxColumn
    Friend WithEvents DrawNo As DataGridViewTextBoxColumn
    Friend WithEvents Qty As DataGridViewTextBoxColumn
    Friend WithEvents CustCode As DataGridViewTextBoxColumn
End Class
