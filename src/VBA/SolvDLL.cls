VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolvDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This private class is used to communicate directly with DLL and is not exposed to User."
'@ModuleDescription "This private class is used to communicate directly with DLL and is not exposed to User."
'@folder("SolverWrapper.Source")
' ==========================================================================
' SolverWrapper v0.9
'
' A wrapper for automating MS Excel's Solver Add-in
'
' https://github.com/GCuser99/SolverWrapper
'
' Contact Info:
'
' https://github.com/GCUser99
' ==========================================================================
' MIT License
'
' Copyright (c) 2024, GCUser99 (https://github.com/GCuser99/SolverWrapper)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==========================================================================

'This private class is needed because the callback must reside in a class module and be declared public
'for it to be called by the Solver DLL. We do not want the callback to be exposed, hence residing in this private class.
Option Explicit

'declare the Solver DLL that does most of the work...
Private Declare PtrSafe Function Solv Lib "Solver32.dll" (ByVal callbackClass As Variant, ByVal app As Variant, ByVal wkb As Variant, ByVal mode As Long) As Long

Private Declare PtrSafe Function SetDllDirectory Lib "kernel32.dll" Alias "SetDllDirectoryW" (ByVal lpNewDirectory As LongPtr) As Boolean
Private Declare PtrSafe Function GetDllDirectory Lib "kernel32.dll" Alias "GetDllDirectoryW" (ByVal cBufferLength As Long, ByVal lpOldDirectory As LongPtr) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public oSolverSheet As Worksheet
Public saveSolutions As Boolean

'store ShowTrial dialog results
Public stopRestorePrevious As Boolean

'store ShowTrial dialog results
Public enableInternalEvents As Boolean
Public enableUserDefinedCallback As Boolean

'internal events that triggers chained public event in SolvProblem
Public Event ShowTrialInternal(ByVal reason As Long, ByVal trialNum As Long, cancel As Boolean)
Public Event BeforeSolveInternal(cancel As Boolean)
Public Event AfterSolveInternal(ByVal returnMsgCode As Long, ByVal trialNum As Long)
'internal event that triggers chained public callback function in SolvProblem
Public Event ShowTrialCallbackInternal(ByVal reason As Long, ByVal trialNum As Long, cancel As Boolean)

Private trialNum As Long
Private saveSolverShow As Boolean
Private solutions() As Variant
Private solverOpened As Boolean
Private solverClosed As Boolean

Private Enum SlvSolverMode
    SolveMode = 0
    CloseMode = 1
    CancelRestoreMode = 2
End Enum

Public Function RunSolver() As Long
    Dim saveScreenUpdating As Long
    Dim returnMsgCode As Integer
    Dim dllFolderPath As String
    Dim cancel As Boolean
    Const pathBufSize = 1024
    Dim saveDllFolderPathBuf As String * pathBufSize
    Dim saveDllFolderPath As String
    
    saveSolverShow = SolverShow
    
    If Not saveSolverShow Then
        If saveSolutions Or enableInternalEvents Then
            SolverShow = True
        End If
    End If
    
    If enableInternalEvents Then
        RaiseEvent BeforeSolveInternal(cancel)
        If cancel Then Exit Function
        RaiseEvent ShowTrialInternal(1, trialNum, cancel)
        If cancel Then Exit Function
    End If
    
    'remember the screen updating state
    saveScreenUpdating = Application.ScreenUpdating
    'turn off updating and alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If saveSolutions Then
        'load header and initial solution into solutions array
        AddHeaderToSolutions solutions, oSolverSheet
        AddIterationToSolutions solutions, oSolverSheet
    End If
    
    'get the previous DLL directory to restore after solver is called
    GetDllDirectory pathBufSize, StrPtr(saveDllFolderPathBuf)
    saveDllFolderPath = Left$(saveDllFolderPathBuf, InStr(saveDllFolderPathBuf, vbNullChar) - 1)

    'this is required as solver DLL is not in the search path list
    dllFolderPath = Application.LibraryPath & Application.PathSeparator & "Solver"
    SetDllDirectory StrPtr(dllFolderPath)
    
    'run a premptive close mode just in case solver did not get closed on a previous run due to user stopping before complete
    Solv callbackClass:=Me, app:=Application, wkb:=oSolverSheet.Parent, mode:=SlvSolverMode.CloseMode
    'now run solver solve mode
    solverOpened = True
    returnMsgCode = Solv(callbackClass:=Me, app:=Application, wkb:=oSolverSheet.Parent, mode:=SlvSolverMode.SolveMode)

    If IsError(returnMsgCode) Then returnMsgCode = 9
    
    Select Case returnMsgCode
    Case 4, 5, 7, 8, 9, 11, 12, 13, 18, 19, 20
        'Cancel/Restore previous decision variable values:
        '-------------------------------------------------
        '(4) The Objective Cell values do not converge
        '(5) Solver could not find a feasible solution
        '(7) The linearity conditions required by this LP Solver are not satisfied
        '(8) The problem is too large for Solver to handle
        '(9) Solver encountered an error value in the objective cell or a constraint cell
        '(11) There is not enough memory available to solve the problem
        '(12) No documentation
        '(13) Error in model. Please verify that all cells and constraints are valid
        '(18) All variables must have both upper and lower bounds
        '(19) Variable bounds conflict in binary or alldifferent constraint
        '(20) Lower and upper bounds on variables allow no feasible solution
        '-------------------------------------------------
        'solver cancel/restore original decision vars
        Solv callbackClass:=Me, app:=Application, wkb:=oSolverSheet.Parent, mode:=SlvSolverMode.CancelRestoreMode
    Case 6
        'Solver stopped at user's request
        If stopRestorePrevious Then Solv callbackClass:=Me, app:=Application, wkb:=oSolverSheet.Parent, mode:=SlvSolverMode.CancelRestoreMode
    End Select

    'close solver
    Solv callbackClass:=Me, app:=Application, wkb:=oSolverSheet.Parent, mode:=SlvSolverMode.CloseMode
    solverClosed = True 'this will be used in terminate event to determine if solver has been closed, and take action if needed
    
    'load final solution (after solver is finished)
    If saveSolutions Then AddIterationToSolutions solutions, oSolverSheet
    
    'restore the previous DLL directory
    SetDllDirectory StrPtr(saveDllFolderPath)
    
    If SolverShow <> saveSolverShow Then
        SolverShow = saveSolverShow
    End If
    
    'put worksheet state back to before calls to DLL
    Application.DisplayAlerts = True
    Application.ScreenUpdating = saveScreenUpdating
    
    RunSolver = returnMsgCode

    'post return message to status bar
    Application.StatusBar = "Excel Solver - " & ReturnSolverMsg(returnMsgCode)
    
    'finish up on events
    If enableInternalEvents Then
        trialNum = trialNum + 1
        If returnMsgCode <> 6 Then RaiseEvent ShowTrialInternal(1, trialNum, cancel)
        RaiseEvent AfterSolveInternal(returnMsgCode, trialNum)
    End If
End Function

Public Function Callback(callbackReason As Integer) As Long
    'This DLL callback executes if option StepThru = True, user pressing the esc key,
    'or any one of five reasons (see the SlvShowTrial Enum).
    'The Callback function must return 1 if Solver should stop, or 0 if Solver should continue running.
    
    'The purpose of this callback is to:
    '1) help user step through each Solver trial via dialog (StepThru = True or user hit esc)
    '2) save intermediate solutions if user set SaveAllTrialSolutions = True
    '3) trigger the ShowTrial event if user set EnableEvents = True

    Dim cancel As Boolean
    
    If saveSolutions Then AddIterationToSolutions solutions, oSolverSheet
        
    'check for conditions under which ShowTrial dialog should be shown. Note that if
    'enableInternalEvents=True then it is assumed that user will handle callbackReason>=2
    If GetAsyncKeyState(vbKeyEscape) <> 0 Or saveSolverShow Or (callbackReason >= 2 And Not enableInternalEvents) Then
        'either esc key was pressed, or user specified StepThru = True, or MaxTime limit reached
        'display the interrupt dialog
        Callback = ShowDialog(callbackReason)
        If Callback = SlvShowTrial.slvStop Then Exit Function
    End If
    
    If saveSolutions Then
        'continue because we are saving all solutions
        Callback = SlvShowTrial.slvContinue
    End If
    
    'increment counter to pass on ShowTrial event
    trialNum = trialNum + 1
    
    If enableInternalEvents Then RaiseEvent ShowTrialInternal(callbackReason, trialNum, cancel)
    
    If enableUserDefinedCallback Then
        RaiseEvent ShowTrialCallbackInternal(callbackReason, trialNum, cancel)
    End If
    
    If cancel Then Callback = SlvShowTrial.slvStop Else Callback = SlvShowTrial.slvContinue
End Function

Private Function ShowDialog(ByVal reason As Long) As Long
    Dim saveScreenUpdating As Boolean
    Dim engine As SlvSolveMethod
    
    saveScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
    
    engine = NameToLong("solver_eng", oSolverSheet)
    
    Select Case reason
    Case 1
        dlgShowTrial.lblMessage.Caption = "Solver paused - current solution values displayed on worksheet."
    Case 2
        dlgShowTrial.lblMessage.Caption = "The maximum time limit was reached - continue anyway?"
    Case 3
        dlgShowTrial.lblMessage.Caption = "The maximum iteration limit was reached - continue anyway?"
    Case 4
        If engine = slvEvolutionary Then
            dlgShowTrial.lblMessage.Caption = "The maximum number of subproblems was reached - continue anyway?"
        Else
            dlgShowTrial.lblMessage.Caption = "The maximum number of integer subproblems was reached - continue anyway?"
        End If
    Case 5
        If engine = slvEvolutionary Then
            dlgShowTrial.lblMessage.Caption = "The maximum number of feasible solutions was reached - continue anyway?"
        Else
            dlgShowTrial.lblMessage.Caption = "The maximum number of integer solutions was reached - continue anyway?"
        End If
    End Select
    
    'display the ShowTrial dialog
    'even though VBA defaults to vbModal, twinBASIC form default to vbModeless, so be explicit
    dlgShowTrial.Show vbModal
    
    'store the ShowTrial dialog results
    ShowDialog = dlgShowTrial.GetShowTrialAction
    stopRestorePrevious = dlgShowTrial.GetStopRestorePrevious
    
    'unload dialog from memory (triggers terminate event if important)
    #If TWINBASIC Then
    dlgShowTrial.Close
    Set dlgShowTrial = Nothing
    #Else
    Unload dlgShowTrial
    #End If
    
    Application.ScreenUpdating = saveScreenUpdating
End Function

Public Function GetSolutions() As Variant()
    GetSolutions = solutions
End Function

Private Property Let SolverShow(ByVal val As Boolean)
    With oSolverSheet
        If val Then
           .Names("solver_sho").value = 1
        Else
           .Names("solver_sho").value = 2
        End If
    End With
End Property
Private Property Get SolverShow() As Boolean
    On Error Resume Next
    SolverShow = (NameToLong("solver_sho", oSolverSheet) = 1)
End Property

Private Sub Class_Terminate()
    If solverOpened And Not solverClosed Then
        MsgBox "Solver DLL not closed properly - closing now..."
        Solv callbackClass:=Me, app:=Application, wkb:=oSolverSheet.Parent, mode:=SlvSolverMode.CloseMode
    End If
    Set oSolverSheet = Nothing
End Sub
