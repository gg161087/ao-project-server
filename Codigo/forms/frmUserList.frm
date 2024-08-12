VERSION 5.00
Begin VB.Form frmUserList 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug de Userlist"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEcharTodosLosNoLogged 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Echar todos los no Logged"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdActualiza 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Actualiza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2175
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActualiza_Click()
    Dim LoopC As Integer
    Text2.Text = "MaxUsers: " & MaxUsers & vbCrLf
    Text2.Text = Text2.Text & "LastUser: " & LastUser & vbCrLf
    Text2.Text = Text2.Text & "NumUsers: " & NumUsers & vbCrLf
    List1.Clear
    For LoopC = 1 To MaxUsers
        List1.AddItem Format(LoopC, "000") & " " & IIf(UserList(LoopC).flags.UserLogged, UserList(LoopC).Name, "")
        List1.ItemData(List1.NewIndex) = LoopC
    Next LoopC
End Sub

Private Sub cmdEcharTodosLosNoLogged_Click()
    Dim LoopC As Integer
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).ConnID <> -1 And Not UserList(LoopC).flags.UserLogged Then
            Call CloseSocket(LoopC)
        End If
    Next LoopC
End Sub

Private Sub List1_Click()
    Dim Userindex As Integer
    If List1.ListIndex <> -1 Then
        Userindex = List1.ItemData(List1.ListIndex)
        If Userindex > 0 And Userindex <= MaxUsers Then
            With UserList(Userindex)
                Text1.Text = "UserLogged: " & .flags.UserLogged & vbCrLf
                Text1.Text = Text1.Text & "IdleCount: " & .Counters.IdleCount & vbCrLf
                Text1.Text = Text1.Text & "ConnId: " & .ConnID & vbCrLf
                Text1.Text = Text1.Text & "ConnIDValida: " & .ConnIDValida & vbCrLf
            End With
        End If
    End If
End Sub
