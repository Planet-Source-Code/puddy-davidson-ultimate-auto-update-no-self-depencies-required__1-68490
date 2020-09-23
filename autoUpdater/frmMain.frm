VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Your Project Starts Here"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Thank-you for taking the time to look at my example. You may use it freely, and you may vote and comment freely also."
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        frmAbout.Show
End Sub
