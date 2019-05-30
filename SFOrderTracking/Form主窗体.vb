Public Class Form主窗体
    Private Sub UserLogin_Click(sender As Object, e As EventArgs) Handles UserLogin.Click
        Me.Hide()
        Form用户登录.Show()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AutoUpdater.Mandatory = True
        AutoUpdater.Start("http://192.168.0.76:8013/ERPUpdater.xml")
        StatusStrip1.Items(0).Text = "Ver " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
        StatusStrip1.Items(2).Alignment = ToolStripItemAlignment.Right
        StatusStrip1.Items(2).Text = "          用户名：匿名"
        Me.Width = 900
    End Sub
    Private Sub OrderQuryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrderQuryToolStripMenuItem.Click
        Me.Hide()
        Form订单查询.Show()
    End Sub
    Private Sub OrderScanMenuItem_Click(sender As Object, e As EventArgs) Handles OrderScanMenuItem.Click
        Me.Hide()
        Form主管工序扫描.Show()
    End Sub
    Private Sub 收料扫描ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 收料扫描ToolStripMenuItem.Click
        'Me.Hide()
        Me.Visible = False
        Form收料扫描.Show()
    End Sub
    Private Sub 输入生产单ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 输入生产单ToolStripMenuItem.Click
        Me.Hide()
        Form输入生产单.Show()
    End Sub

    Private Sub 取消订单ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 取消订单ToolStripMenuItem.Click
        Me.Hide()
        Form取消生产单.Show()
    End Sub
    Private Sub 用户查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 用户查询ToolStripMenuItem.Click
        Me.Hide()
        Form用户查询.Show()
    End Sub
    Private Sub 增加用户ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 增加用户ToolStripMenuItem.Click
        Me.Hide()
        Form增加用户.Show()
    End Sub

    Private Sub 订单入库ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 订单入库ToolStripMenuItem.Click
        Me.Hide()
        Form订单入库.Show()
    End Sub

    Private Sub 客户清单ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 客户清单ToolStripMenuItem1.Click
        Me.Hide()
        Form客户信息查询及维护.Show()
    End Sub
    Private Sub Form1_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        'DataGridView1.Visible = False
    End Sub

    Private Sub 收图记录ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 收图记录ToolStripMenuItem.Click
        Me.Hide()
        FormPMC图纸记录.Show()
        FormPMC图纸记录.Label2.Text = "收图清单:"
        FormPMC图纸记录.Label2.Visible = True
        FormPMC图纸记录.Button1.Visible = True
        FormPMC图纸记录.Text = "收图记录"
        FormPMC图纸记录.DrawingReceiveStatus = "Receive"
        FormPMC图纸记录.Button4.Visible = True
    End Sub

    Private Sub 给工艺发图记录ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 给工艺发图记录ToolStripMenuItem.Click
        Me.Hide()
        FormPMC图纸记录.Show()
        FormPMC图纸记录.DrawingReceiveStatus = "Release"
        FormPMC图纸记录.Text = "给工艺发图记录"
        FormPMC图纸记录.Label2.Text = "发图清单:"
        FormPMC图纸记录.Button2.Visible = True
        FormPMC图纸记录.ComboBox1.Visible = True
        FormPMC图纸记录.Label3.Visible = True
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * from UserInfo where UserType='process' and (isnull(ChineseName,'')<>'')"
        rs.Open(SQLstring, cn, 1, 2)
        rs.MoveFirst()
        Do While rs.EOF = False
            FormPMC图纸记录.ComboBox1.Items.Add(rs.Fields("ChineseName").Value)
            rs.MoveNext()
        Loop
    End Sub

    Private Sub 工艺回图记录ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 工艺回图记录ToolStripMenuItem.Click
        Me.Hide()
        FormPMC图纸记录.Show()
        FormPMC图纸记录.DrawingReceiveStatus = "Return"
        FormPMC图纸记录.Text = "工艺回图记录"
        FormPMC图纸记录.Label2.Text = "回图清单:"
        FormPMC图纸记录.Button3.Visible = True
        FormPMC图纸记录.ComboBox1.Visible = True
        'Form15.Label3.Visible = True
        FormPMC图纸记录.ComboBox1.Visible = False
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * from UserInfo where UserType='process' and (isnull(ChineseName,'')<>'')"
        rs.Open(SQLstring, cn, 1, 2)
        rs.MoveFirst()
        Do While rs.EOF = False
            FormPMC图纸记录.ComboBox1.Items.Add(rs.Fields("ChineseName").Value)
            rs.MoveNext()
        Loop
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim ss As String
        'ss = "net use \\192.168.0.186\sfdata" & " " & "sfbase123456" & " /user:" & """" & "erp\SFOrderBase" & """"
        ss = "net use \\192.168.0.186\sfdata 'sfbase123456' /user:'erp\SFOrderBase'"
        Shell(ss, vbHide)
        ss = "net use z: \\192.168.0.186\sfdata"
        Shell(ss, vbHide)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim ss As String
        ss = "net use z: \\192.168.0.186\sfdata /del"
        Shell(ss, vbHide)
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub 未处理图纸查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 未处理图纸查询ToolStripMenuItem.Click
        Me.Hide()
        Form未分发及未收回图纸查询.Show()
    End Sub
    Private Sub 工艺编制ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 工艺编制ToolStripMenuItem1.Click
        Me.Hide()
        Form工艺编写及修改.Show()
    End Sub
    Private Sub 打印工艺卡ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打印工艺卡ToolStripMenuItem.Click
        Me.Hide()
        Form工艺卡打印.Show()
    End Sub
    Private Sub 外放记录输入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 外放记录输入ToolStripMenuItem.Click
        Me.Hide()
        Form外发记录.Show()
    End Sub
    Private Sub 外发收回输入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 外发收回输入ToolStripMenuItem.Click
        Me.Hide()
        Form外发收回记录.Show()
    End Sub
    Private Sub 外发记录查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 外发记录查询ToolStripMenuItem.Click
        Me.Hide()
        Form外发记录查询.Show()
        SubReordQuery = "PMC"
    End Sub

    Private Sub 更改图号ToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles 更改图号ToolStripMenuItem.Click
        Me.Hide()
        Form更改图号.Show()
    End Sub
    Private Sub 取消工艺ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 取消工艺ToolStripMenuItem.Click
        Me.Hide()
        Form取消工艺.Show()
    End Sub
    Private Sub 增加供应商ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 增加供应商ToolStripMenuItem.Click
        Me.Hide()
        Form增加供应商.Show()
    End Sub
    Private Sub 扫描录入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 扫描录入ToolStripMenuItem.Click
        Me.Hide()
        Form品质部终检记录.Show()
    End Sub
    Private Sub 单号图号方式录入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 单号图号方式录入ToolStripMenuItem.Click
        Me.Hide()
        Form按单号图号录入终检记录.Show()
    End Sub
    Private Sub 分类查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 分类查询ToolStripMenuItem.Click
        Me.Hide()
        Form查询终检记录.Show()
        Form查询终检记录.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 全部查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 全部查询ToolStripMenuItem.Click
        Me.Hide()
        Form查询所有终检记录.Show()
    End Sub

    Private Sub 下料完成后收料ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下料完成后收料ToolStripMenuItem.Click
        Me.Hide()
        Form下料完成记录.Show()
    End Sub

    Private Sub 录入图纸ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 录入图纸ToolStripMenuItem.Click
        Me.Hide()
        Form37.Show()
    End Sub

    Private Sub 更改外发供应商ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 更改外发供应商ToolStripMenuItem.Click
        Me.Hide()
        Form修改供应商.Show()
    End Sub

    Private Sub 更新价格ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 更新价格ToolStripMenuItem.Click
        Me.Hide()
        Form更新价格.Show()
    End Sub

    Private Sub 查询价格ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 查询价格ToolStripMenuItem.Click
        Me.Hide()
        Form查询订单信息.Show()
    End Sub

    Private Sub 返工记录输入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 返工记录输入ToolStripMenuItem.Click
        Me.Hide()
        Form返工记录.Show()
    End Sub

    Private Sub 返工收回QAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 返工收回QAToolStripMenuItem.Click
        Me.Hide()
        Form返工收回.Show()
    End Sub

    Private Sub 查询返工记录ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 查询返工记录ToolStripMenuItem.Click
        Me.Hide()
        Form查询返工记录.Show()
    End Sub
    Private Sub 嘉睦零件交期查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦零件交期查询ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦零件交期信息查询.Show()
        Form嘉睦零件交期信息查询.WindowState = FormWindowState.Maximized
    End Sub
    Private Sub 嘉睦零件最新信息ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦零件最新信息ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦零件最新信息查询.Show()
    End Sub
    Private Sub 出库输入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 出库输入ToolStripMenuItem.Click
        Me.Hide()
        Form订单出货记录.Show()
        Form订单出货记录.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 外发记录查询ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 外发记录查询ToolStripMenuItem1.Click
        Me.Hide()
        Form外发记录查询.Show()
        SubReordQuery = "SpecialQuery"
    End Sub

    Private Sub 出货货查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 出货货查询ToolStripMenuItem.Click
        Me.Hide()
        Form出货查询.Show()
        Form出货查询.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 补入订单价格ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 补入订单价格ToolStripMenuItem.Click
        Me.Hide()
        Form补入订单价格.Show()
        Form补入订单价格.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 嘉睦订单录入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦订单录入ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦订单录入.Show()
    End Sub

    Private Sub 嘉睦出货录入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦出货录入ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦出货记录.Show()
    End Sub

    Private Sub 嘉睦订单查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦订单查询ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦订单查询.Show()
    End Sub

    Private Sub 嘉睦零件库存管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦零件库存管理ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦零件库存管理.Show()
    End Sub

    Private Sub 嘉睦出货查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦出货查询ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦出货查询.Show()
        Form嘉睦出货查询.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 嘉睦库存查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦库存查询ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦零件库存查询.Show()
        Form嘉睦零件库存查询.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 系统图纸清单ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 系统图纸清单ToolStripMenuItem.Click
        Me.Hide()
        Form系统图纸清单.Show()
    End Sub

    Private Sub 库存管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 库存管理ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦零件库存管理.Show()
    End Sub

    Private Sub 单号图号定价格ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 单号图号定价格ToolStripMenuItem.Click
        Dim cn As New ADODB.Connection
        Dim rs, rs1 As New ADODB.Recordset
        Dim SQLstring, CNString1 As String
        '********************************** Textbox1 connection ***************************************************************
        On Error Resume Next
        cn.Open(CNsfmdb)
        SQLstring = "Select * From TempPrice"
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        Do While rs.EOF = False
            '按单号图号更新订单价格
            CNString1 = "Select * from SFOrderbase where SFOrder=" & "'" & rs("SFOrder").Value & "'" & " and DrawNo=" & "'" & rs("DrawNo").Value & "'"
            rs1.Open(CNString1, cn, 1, 3)
            rs1("UnitP").Value = rs("UnitP").Value
            rs1.Update()
            rs1.Close()

            '按单号图号更新出货价格
            CNString1 = "Select * from OutputRecord where SFOrder=" & "'" & rs("SFOrder").Value & "'" & " and DrawNo=" & "'" & rs("DrawNo").Value & "'"
            rs1.Open(CNString1, cn, 1, 3)
            rs1("UnitP").Value = rs("UnitP").Value
            rs1.Update()
            rs1.Close()

            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        rs1 = Nothing
        cn = Nothing
        MsgBox("Price input！")

    End Sub

    Private Sub 图号客户代码定价格ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 图号客户代码定价格ToolStripMenuItem.Click
        Me.Hide()
        Form按客户代码及图号输入价格.Show()
    End Sub

    Private Sub 按客户代码和图号更新出货价格ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 按客户代码和图号更新出货价格ToolStripMenuItem.Click
        Me.Hide()
        Form补充出货价格.Show()
    End Sub
    Private Sub 材料信息查找ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 材料信息查找ToolStripMenuItem.Click
        Me.Hide()
        Form根据材料信息查图号.Show()
    End Sub
    Private Sub 工艺信息查找ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 工艺信息查找ToolStripMenuItem.Click
        Me.Hide()
        Form查询工艺信息.Show()
    End Sub
    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles 嘉睦未交货订单查询.Click
        Me.Hide()
        Form嘉睦未交货订单查询.Show()
        Form嘉睦未交货订单查询.WindowState = FormWindowState.Maximized
    End Sub
    Private Sub 客户退货输入_Click(sender As Object, e As EventArgs) Handles 客户退货输入.Click
        Me.Hide()
        Form客户退货输入.Show()
    End Sub

    Private Sub 删除出货记录ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除出货记录ToolStripMenuItem.Click
        Me.Hide()
        Form删除出货记录.Show()
    End Sub

    Private Sub 按客户月出货查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 按客户月出货查询ToolStripMenuItem.Click
        Me.Hide()
        Form按客户月出货查询.Show()
    End Sub

    Private Sub 嘉睦月出货查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦月出货查询ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦月出货汇总.Show()
    End Sub

    Private Sub ToolStripMenuItem2_Click_1(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Me.Hide()
        Form按客户月出货查询.Show()
    End Sub

    Private Sub 查询图纸ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 查询图纸ToolStripMenuItem.Click
        Me.Hide()
        Form查询图纸.Show()
    End Sub

    Private Sub 嘉睦退货输入_Click(sender As Object, e As EventArgs) Handles 嘉睦退货输入.Click
        Me.Hide()
        Form嘉睦退货录入.Show()
    End Sub

    Private Sub 删除订单项ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除订单项ToolStripMenuItem.Click
        Me.Hide()
        Form删除订单项.Show()
    End Sub
    Private Sub 按客户月订单汇总查询_Click(sender As Object, e As EventArgs) Handles 按客户月订单汇总查询.Click
        Me.Hide()
        Form按客户月订单汇总查询.Show()
    End Sub

    Private Sub 组员增减ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 组员增减ToolStripMenuItem.Click
        Me.Hide()
        Form管理者挑选组员.Show()
    End Sub

    Private Sub 新员工录入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 新员工录入ToolStripMenuItem.Click
        Me.Hide()
        Form员工信息维护.Show()
    End Sub

    Private Sub 员工离职录入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 员工离职录入ToolStripMenuItem.Click
        Me.Hide()
        Form员工离职录入.Show()
    End Sub

    Private Sub 嘉睦月订单查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦月订单查询ToolStripMenuItem.Click
        Me.Hide()
        Form83.Show()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Me.Hide()
        Form缺少价格的生产单查询.Show()
    End Sub

    Private Sub 按日期查询生产单ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 按日期查询生产单ToolStripMenuItem.Click
        Me.Hide()
        Form按日期查询所下生产单.Show()
    End Sub

    Private Sub 嘉睦出货记录更改ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 嘉睦出货记录更改ToolStripMenuItem.Click
        Me.Hide()
        Form嘉睦出货记录更改.Show()
    End Sub

    Private Sub 订单跟进备注ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 订单跟进备注ToolStripMenuItem.Click
        Me.Hide()
        Form跟单备注.Show()
        Form跟单备注.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 按工序个人工时查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 按工序个人工时查询ToolStripMenuItem.Click
        Me.Hide()
        Form工序个人工时查询.Show()
        Form工序个人工时查询.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 查询价格ToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub 查询价格ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles 查询价格ToolStripMenuItem2.Click
        Me.Hide()
        Form查询价格.Show()
        Form查询价格.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 打印工艺卡ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 打印工艺卡ToolStripMenuItem1.Click
        Me.Hide()
        Form工艺卡打印.Show()
    End Sub
End Class
