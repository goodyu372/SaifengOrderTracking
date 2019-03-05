Module Module1
    Public ProcessWatch As Process() = Process.GetProcessesByName("Excel")
    Public UserType, UserName, ChineseName, ProcessName As String
    'Public CNsfmdb = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=czlee859288;Initial Catalog=SFOrderBase;Server =\\192.168.0.186;Data Source=ERP\SAIFENGDB"
    'Public CNsfmdb = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=czlee859288;Initial Catalog=SFAll;Server =\\192.168.0.186;Data Source=ERP\SAIFENGDB"
    Public CNsfmdb = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123;Initial Catalog=SFAll;Server =\\sf;Data Source=SF"
    Public erpBasePath = "\\sf\【可执行文件共享】\ERP Data"
    Public PrintCard = erpBasePath + "\Data\NewPrintCard.xlsx"  '赛峰工艺卡打印模板
    Public TempDataPath = "D:\TempData\"  '用户本地计算机D盘存放一些临时文件的目录'
    Public DWGPath = erpBasePath + "\sfData\DWG\"  'DWG等图纸文件存放本机的路径'
    'Public SFOrderBaseNetConnect = "net use \\192.168.0.186" & " " & "sfbase123456" & " /user:" & """" & "erp\SFOrderBase" & """"
    'Public SFOrderBaseNetConnect = "net use \\192.168.0.186\ipc$" & " " & "sfbase123456" & " /user:" & """" & "erp\SFOrderBase" & """"
    Public SFOrderBaseNetConnect = "net use \\erp\ipc$" & " " & "sfbase123456" & " /user:" & """" & "erp\SFOrderBase" & """"
    ' Public SFOrderBaseNetDisconnect = "net use \\192.168.0.186\ipc$ /del"
    Public SFOrderBaseNetDisconnect = "net use \\erp\ipc$ /del"
    Public ProcessForm As String
    Public SubReordQuery, JamuStockForm As String

    Public Sub LoginUser(UserType)
        Select Case UserType
            Case "lathe"
                ProcessName = "车工"
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.组员增减ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.Width = 600
            Case "milling"
                ProcessName = "铣工"
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.组员增减ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.Width = 800
            Case "grinding"
                ProcessName = "磨工"
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.组员增减ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.Width = 600
            Case "cnc"
                ProcessName = "CNC"
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.组员增减ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.Width = 600
            Case "fitting"
                ProcessName = "钳工"
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.组员增减ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.Width = 600
            Case "wc"
                ProcessName = "线割"
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.组员增减ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.Width = 600
            Case "cylinder"
                ProcessName = "内外圆磨"
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.组员增减ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.Width = 600
            Case "edm"
                ProcessName = "火花机"
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.组员增减ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.Width = 600
            Case "admin"
                ProcessName = ""
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.组员增减ToolStripMenuItem.Visible = True
                Form主窗体.OrderQuryToolStripMenuItem.Visible = True
                Form主窗体.OrderInputMenuItem.Visible = True
                Form主窗体.ProcessInputMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = True
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = True
                Form主窗体.SystemInfoToolStripMenuItem.Visible = True
                Form主窗体.PMC图纸记录ToolStripMenuItem.Visible = True
                Form主窗体.返工记录ToolStripMenuItem.Visible = True
                Form主窗体.客户清单ToolStripMenuItem1.Visible = True
                Form主窗体.嘉睦出货记录更改ToolStripMenuItem.Visible = True
                Form主窗体.增加用户ToolStripMenuItem.Visible = True
                Form主窗体.更改图号ToolStripMenuItem.Visible = True
                Form主窗体.Width = 1400
                Form订单查询.Button1.Visible = True
                Form主窗体.查询价格ToolStripMenuItem2.Enabled = True
            Case "sales"
                Form主窗体.OrderInputMenuItem.Visible = True
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.ProcessInputMenuItem.Visible = False
                Form主窗体.UserAdminToolStripMenuItem.Visible = True
                Form主窗体.用户查询ToolStripMenuItem.Visible = False
                Form主窗体.增加用户ToolStripMenuItem.Visible = False
                Form主窗体.全部查询ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.查询价格ToolStripMenuItem2.Enabled = True
                Form主窗体.SystemInfoToolStripMenuItem.Visible = True
            Case "process"
                Form主窗体.OrderInputMenuItem.Visible = False
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.ProcessInputMenuItem.Visible = True
                Form主窗体.UserAdminToolStripMenuItem.Visible = True
                Form主窗体.用户查询ToolStripMenuItem.Visible = False
                Form主窗体.增加用户ToolStripMenuItem.Visible = False
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.查询价格ToolStripMenuItem2.Enabled = True
            Case "plan"
                Form主窗体.OrderInputMenuItem.Visible = False
                'Form1.OrderScanMenuItem.Visible = True
                Form主窗体.ProcessInputMenuItem.Visible = False
                Form主窗体.UserAdminToolStripMenuItem.Visible = True
                Form主窗体.用户查询ToolStripMenuItem.Visible = False
                Form主窗体.增加用户ToolStripMenuItem.Visible = False
                Form主窗体.PMC图纸记录ToolStripMenuItem.Visible = True
                Form主窗体.全部查询ToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.查询价格ToolStripMenuItem2.Enabled = True
            Case "manager"
                Form主窗体.OrderInputMenuItem.Visible = False
                Form主窗体.OrderScanMenuItem.Visible = False
                Form主窗体.ProcessInputMenuItem.Visible = False
                Form主窗体.UserAdminToolStripMenuItem.Visible = True
                Form主窗体.UserLogin.Visible = True
                Form主窗体.用户查询ToolStripMenuItem.Visible = False
                Form主窗体.增加用户ToolStripMenuItem.Visible = False
                Form主窗体.全部查询ToolStripMenuItem.Visible = True
                Form主窗体.SystemInfoToolStripMenuItem.Visible = True
                Form主窗体.客户清单ToolStripMenuItem1.Visible = False
                Form主窗体.员工信息查看ToolStripMenuItem1.Visible = False
                Form主窗体.更改图号ToolStripMenuItem.Visible = False
                Form主窗体.更新价格ToolStripMenuItem.Visible = False
                Form主窗体.取消工艺ToolStripMenuItem.Visible = False
                Form主窗体.删除出货记录ToolStripMenuItem.Visible = False
                Form主窗体.删除订单项ToolStripMenuItem.Visible = False
                Form主窗体.查询图纸ToolStripMenuItem.Visible = False
                Form主窗体.系统图纸清单ToolStripMenuItem.Visible = False
                Form主窗体.查询价格ToolStripMenuItem2.Enabled = True
                Form主窗体.Width = 1400
            Case "scan"
                Form主窗体.OrderInputMenuItem.Visible = False
                'Form1.OrderScanMenuItem.Visible = True
                Form主窗体.ProcessInputMenuItem.Visible = False
                Form主窗体.UserAdminToolStripMenuItem.Visible = True
                Form主窗体.UserLogin.Visible = True
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.用户查询ToolStripMenuItem.Visible = False
                Form主窗体.增加用户ToolStripMenuItem.Visible = False
            Case "HRInput"
                Form主窗体.OrderInputMenuItem.Visible = False
                'Form1.OrderScanMenuItem.Visible = True
                Form主窗体.ProcessInputMenuItem.Visible = False
                Form主窗体.UserAdminToolStripMenuItem.Visible = True
                Form主窗体.UserLogin.Visible = True
                Form主窗体.OrderScanMenuItem.Visible = True
                Form主窗体.用户查询ToolStripMenuItem.Visible = False
                Form主窗体.增加用户ToolStripMenuItem.Visible = False
                Form主窗体.SystemInfoToolStripMenuItem.Visible = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
            Case "JMStock"
                Form主窗体.嘉睦零件管理ToolStripMenuItem.Enabled = True
                Form主窗体.ToolStripMenuItem1.Visible = False
                Form主窗体.品质部终检记录ToolStripMenuItem.Visible = False
                Form主窗体.Width = 800
        End Select
    End Sub
End Module
