VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgShowTrial 
   Caption         =   "Show Trial Solution"
   ClientHeight    =   1752
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   5700
   OleObjectBlob   =   "dlgShowTrial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgShowTrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@folder("Solver.Source")
Option Explicit

Private showTrialAction As SlvShowTrial
Private stopRestorePrevious As Boolean

Public Function GetShowTrialAction() As SlvShowTrial
    GetShowTrialAction = showTrialAction
End Function

Public Function GetStopRestorePrevious() As Boolean
    GetStopRestorePrevious = stopRestorePrevious
End Function

Private Sub cmdContinue_Click()
    showTrialAction = SlvShowTrial.slvContinue
    stopRestorePrevious = False
    Hide
End Sub

Private Sub cmdStopAndRestore_Click()
    showTrialAction = SlvShowTrial.slvStop
    stopRestorePrevious = True
    Hide
End Sub

Private Sub cmdStopAndKeep_Click()
    showTrialAction = SlvShowTrial.slvStop
    stopRestorePrevious = False
    Hide
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Show Trial Solution"
    Me.cmdContinue.Caption = "Continue"
    Me.cmdStopAndRestore.Caption = "Stop and Restore"
    Me.cmdStopAndKeep.Caption = "Stop and Keep"
    Me.cmdContinue.Accelerator = "C"
    Me.cmdStopAndRestore.Accelerator = "R"
    Me.cmdStopAndKeep.Accelerator = "K"
End Sub

Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then cmdContinue_Click
End Sub

