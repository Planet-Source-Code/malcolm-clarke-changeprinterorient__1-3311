VERSION 5.00
Begin VB.Form frmChangePrinterOrient 
   Caption         =   "Change Printer Orientation"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangePrinterOrient 
      Caption         =   "Change Printer Orientation"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.OptionButton optLand 
      Caption         =   "Landscape"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Portrait"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmChangePrinterOrient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdChangePrinterOrient_Click()
      
    If optPort.Value = True Then
        ChngPrinterOrientationPortrait Me
    ElseIf optLand.Value = True Then
        ChngPrinterOrientationLandscape Me
    End If
        
End Sub

