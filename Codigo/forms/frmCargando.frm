VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCargando 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   2880
   ClientLeft      =   1410
   ClientTop       =   3000
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   135.849
   ScaleMode       =   0  'User
   ScaleWidth      =   440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar cargar 
      Height          =   255
      Left            =   1354
      TabIndex        =   0
      Top             =   2544
      Width           =   3891
      _ExtentX        =   6853
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Min             =   1e-4
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3000
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " aa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblCargando 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando, por favor espere..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   2020
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private VersionNumberMaster As String
Private VersionNumberLocal As String

Private Sub Form_Load()
    lblVersion(2).Caption = GetVersionOfTheServer()
    Me.VerifyIfUsingLastVersion
End Sub

Function VerifyIfUsingLastVersion()
    On Error Resume Next
    If Not (CheckIfRunningLastVersion) Then
        If MsgBox("Tu version no es la actual, Deseas ejecutar el actualizador?. - Tu version: " & VersionNumberLocal & " Ultima version: " & VersionNumberMaster & " -- Your version is not up to date, open the launcher to update? ", vbYesNo) = vbYes Then
            Call ShellExecute(Me.hWnd, "open", App.Path & "\Autoupdate.exe", "", "", 1)
            End
        End If
    End If
End Function

Private Function CheckIfRunningLastVersion() As Boolean
    Dim responseGithub As String
    Dim JsonObject     As Object
    responseGithub = Inet1.OpenURL("https://api.github.com/repos/gg161087/ao-project-server/releases/latest")
    If Len(responseGithub) = "" Then Exit Function
    Set JsonObject = modJSON.parse(responseGithub)
    VersionNumberMaster = JsonObject.Item("tag_name")
    MsgBox (VersionNumberMaster)
    VersionNumberLocal = GetVar(App.Path & "\Server.ini", "INIT", "VersionTagRelease")
    If VersionNumberMaster = VersionNumberLocal Then
        CheckIfRunningLastVersion = True
    Else
        CheckIfRunningLastVersion = False
    End If
End Function
