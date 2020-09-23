VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "File Type Generator"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   735
      Left            =   2160
      TabIndex        =   17
      Top             =   4560
      Width           =   5535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox printprog 
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox openprog 
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   1200
      Width           =   4575
   End
   Begin VB.TextBox content 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox ext 
      Height          =   285
      Left            =   120
      MaxLength       =   3
      TabIndex        =   8
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox full 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox prop 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create File Type"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   7575
   End
   Begin VB.Label Label9 
      Caption         =   $"Form1.frx":030A
      Height          =   1095
      Left            =   3000
      TabIndex        =   15
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label Label8 
      Caption         =   "Program to Print With (Path and Filename)"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "Program to Open With (Path and Filename)"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "Content Type"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "File Extension"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Full File Name"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Proper File Name (NO SPACES)"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0425
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0549
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   7695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Form1.frx":064B
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   -480
      Shape           =   3  'Circle
      Top             =   -480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim myfiletype As filetype

myfiletype.ProperName = prop.Text
myfiletype.FullName = full.Text
myfiletype.ContentType = content.Text
myfiletype.extension = "." & ext.Text
myfiletype.Commands.Captions.Add "Open"
myfiletype.Commands.Commands.Add openprog.Text & " ""%1"""
myfiletype.Commands.Captions.Add "Print"
myfiletype.Commands.Commands.Add printprog.Text & " ""%1"" /P"
CreateExtension myfiletype
End Sub

Private Sub Command2_Click()
MsgBox "File Type Generator v1.0" & vbCrLf & "©2001 Adam Tillotson" & vbCrLf & "All Rights Reserved" & vbCrLf & "Feel free to take this code and use it in an install program or your program.  BE SURE TO VOTE!!!", vbInformation, "About..."
End Sub

Private Sub Command3_Click()
End
End Sub

