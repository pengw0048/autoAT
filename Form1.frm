VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7875
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   6960
      TabIndex        =   11
      Text            =   "5"
      Top             =   600
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5040
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   5760
      TabIndex        =   6
      Text            =   "5"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2040
      TabIndex        =   5
      Text            =   "������֣�"
      Top             =   600
      Width           =   1935
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   6588
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label7 
      Caption         =   "��ʱ�룺"
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "���ȣ�"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "ÿ���ظ�AT��������"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "ÿ���ظ�֮ǰ�����֣�"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "���ţ����������ѡ�"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "��¼����˺ţ�����һ��״̬��Ȼ����㿪����״̬�Ļظ��б�"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim names(5000) As String, uid(5000) As String, url1 As String
Dim d As HTMLDocument
Dim un As Long, pos As Long
Dim run As Boolean
Dim stage As Integer, letter As Integer, page As Integer
Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long



Private Sub Command1_Click()
If run Then Exit Sub
If Val(Text2.Text) < 1 Or Val(Text2.Text) > 6 Then
MsgBox "AT���������԰�"
Exit Sub
End If
run = True
url1 = WebBrowser1.LocationURL
stage = 1
letter = 0
page = 0
WebBrowser1.Navigate "http://3g.renren.com/status/search.do?type=letter&name=A"
End Sub

Private Sub Command2_Click()
run = False
End Sub

Private Sub Form_Load()
Me.Show
Form_Resize
WebBrowser1.Navigate "http://3g.renren.com"
End Sub

Private Sub Form_Resize()
WebBrowser1.Width = Form1.Width - 200
WebBrowser1.Height = Form1.Height - WebBrowser1.Top - 500
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Set d = WebBrowser1.Document
If stage = 1 Then
  For i = 0 To d.All.length - 1
    If d.All.Item(i).innerText = "û���ҵ�����" Then
      letter = letter + 1
      If letter = 26 Then
        stage = 2
        For j = 0 To un - 1
          Debug.Print names(j) & "(" & uid(j) & ")"
        Next j
        WebBrowser1.Navigate url1
        Exit Sub
      End If
      page = 0
      WebBrowser1.Navigate "http://3g.renren.com/status/search.do?type=letter&name=" & Chr(Asc("A") + letter)
      Exit Sub
    End If
  Next i
  For i = 0 To d.getElementsByTagName("A").length - 1
    If Left(d.getElementsByTagName("A").Item(i).innerHTML, 1) = "@" Then
      uid(un) = Split(Mid(d.getElementsByTagName("A").Item(i).outerHTML, InStr(1, d.getElementsByTagName("A").Item(i).outerHTML, "atuid=")), "&")(0)
      uid(un) = Mid(uid(un), 7)
      names(un) = d.getElementsByTagName("A").Item(i).innerHTML
      un = un + 1
      Label6.Caption = "��ȡ�û���Ϣ" & Trim(un)
      DoEvents
    End If
  Next i
  page = page + 1
  WebBrowser1.Navigate "http://3g.renren.com/status/search.do?type=letter&name=" & Chr(Asc("A") + letter) & "&curpage=" & Trim(page)
ElseIf stage = 2 Then
  ts = Text1.Text
  For i = 0 To Val(Text2.Text) - 1
    If pos < un Then
      ts = ts + names(pos) & "(" & uid(pos) & ") "
    End If
    pos = pos + 1
  Next i
  Label6.Caption = "������Ϣ" & Trim(pos) & "/" & Trim(un)
  DoEvents
  d.getElementsByName("status")(0).Value = ts
  d.getElementsByName("update")(0).Click
  stage = 3
  Exit Sub
ElseIf stage = 3 Then
  Debug.Print pos
  If pos >= un Then
    MsgBox "����"
    End
    Exit Sub
  End If
  Label6.Caption = "�ȴ���һ��..."
  DoEvents
  SleepEx 1000 * Val(Text3.Text), 1
  stage = 2
  WebBrowser1.Navigate url1
End If
End Sub
