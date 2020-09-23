VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "x"
      Height          =   135
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton Command3 
      Height          =   735
      Left            =   1440
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   720
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   0
      Picture         =   "Form1.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Open application"
      Height          =   255
      Left            =   30
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Send screenshot to:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ConvertWordDoc
Unload Form1
End Sub

Private Sub Command2_Click()
ConvertExcel
Unload Form1
End Sub

Private Sub Command3_Click()
ConvertPower
Unload Form1
End Sub

Private Sub Command4_Click()
End
End Sub

