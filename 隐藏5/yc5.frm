VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "����5"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Check1 
      Caption         =   "ͬ����Լ���κ������������޹أ���"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3360
      Top             =   480
   End
   Begin VB.Label Label1 
      Caption         =   "�����ּ��رջ�򿪼�����ӽ��ҡ��������졣�����˴����أ���Ŀ¼���½�guanbi.txt���رգ���"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Function find() As Integer
find = 0

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set ps = objWMIService.ExecQuery("Select * from Win32_Process")

For Each p In ps
If p.Name = "StudentEx.exe" Then
find = 1
End If
Next

End Function

Private Sub Check1_Click()
If Check1.Value = 0 Then End
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "88", , "ϵͳ��ʾ"
End Sub

Private Sub Label1_Click()
MsgBox "�������ڷǷ�;����Υ������֮�£����˶Դ˲����κ����Ρ�", , "����"
Me.Hide
End Sub

Private Sub Timer1_Timer()
Dim b As Integer
If Check1.Value = 1 Then
x = GetAsyncKeyState(145)
If x < 0 Then

If a = 0 Then
b = find
  If b = 1 Then
  k
  Else
  s
  End If
End If

a = a + 1

Else
a = 0
End If
End If
If Dir("guanbi") <> "" Then Unload Me
End Sub
Private Sub k()
Shell "taskkill -f -im StudentEx.exe"
End Sub
Private Sub s()
Shell "C:\Program Files\TopDomain\StudentEx.exe"
End Sub
