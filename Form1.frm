VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Exit Program"
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Menu Button 2"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Menu Button 1"
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   10
         Height          =   495
         Left            =   360
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Welcome to Menu Selection Program"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   525
         TabIndex        =   5
         Top             =   2760
         Width           =   3285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "MENU OPTIONS"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   120
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* MENU HIGHLIGHT BUTTONS made by Victor Detaro
'* Sample Program to show 3 different ways to highlight buttons
'* when mouse moves over them
'*
'* FEEL FREE TO USE OR MODIFY,
'* BUT DON'T FORGET TO INCLUDE MY NAME IN YOUR CREDITS IF YOU DO USE

Option Explicit

'* Array vic is used as temporary strings to store button captions
'* det is standard integer
Dim vic(2) As String, det As Integer

Private Sub Command1_Click(Index As Integer)
    'BUTTONS CLICKED - CREATE EVENTS
    Select Case Index
    Case 0
        MsgBox "You selected Menu Button No. 1"
    Case 1
        MsgBox "You selected Menu Button No. 2"
    Case 2
        MsgBox "Bye for now"
        End
    End Select
End Sub

'MOUSE MOVED OVER BUTTONS
Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    'TO AVOID FLICKERING OF HIGHLIGHT
    If Shape1.Visible = False Then
        
        'HIGHLIGHT BY WHITE SHAPE BOX
        Shape1.Top = Command1(Index).Top
        Shape1.Left = Command1(Index).Left
        Shape1.Width = Command1(Index).Width
        Shape1.Height = Command1(Index).Height
        Shape1.Visible = True
        
        'HIGHLIGHT BY CAPITALIZING AND BOLDING BUTTON CAPTIONS
        Command1(Index).FontBold = True
        Command1(Index).Caption = Format(vic(Index), ">")
        
        'HIGHLIGHT BY DEFINING BUTTON ACTION
        Select Case Index
        Case 0
            Label2 = "Menu Button 1"
        Case 1
            Label2 = "Menu Button 2"
        Case 2
            Label2 = "Exit Program"
        End Select
    End If
End Sub

Private Sub Form_Load()
    'INITIALIZING BUTTON CAPTIONS
    For det = 0 To 2
        vic(det) = Command1(det).Caption
    Next det
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'REMOVING WHITE BOX HIGHLIGHT
    Shape1.Visible = False
    
    'REMOVING CAPITALIZATION AND BOLD OF BUTTON CAPTIONS
    For det = 0 To 2
        If Command1(det).FontBold = True Then
            Command1(det).FontBold = False
            Command1(det).Caption = vic(det)
        End If
    Next det
    
    'REMOVING DEFINITION OF BUTTON ACTIONS
    Label2 = "Welcome to Menu Selection Program"
End Sub
