Attribute VB_Name = "test_Using_Callback_Macro"
'@folder("SolverWrapper.Examples")

Option Explicit

'This example automates solving the problem in SOLVSAMP.XLS on the "Portfolio of Securities" worksheet.
'
'SOLVSAMP.XLS is distributed with MS Office Excel and can be found in:
'
'Application.LibraryPath & "\..\SAMPLES\SOLVSAMP.XLS"
'
'which on many systems can be found here:
'C:\Program Files\Microsoft Office\root\Office16\SAMPLES\SOLVSAMP.XLS
'
'Import this module into the sample workbook, set a reference to the SolverWrapper code library
'and then save SOLVSAMP.XLS to SOLVSAMP.XLSM.

'To use SolverWrapper callback feature, user must write their own callback function (see below)
Sub Solve_Portfolio_of_Securities_with_Events()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Portfolio of Securities")
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "E18", slvMaximize
    
    oProblem.DecisionVars.Add "E10:E14"
    oProblem.DecisionVars.Initialize 0.2, 0.2, 0.2, 0.2, 0.2
    
    'add some constraints
    With oProblem.Constraints
        .AddBounded "E10:E14", 0#, 1#
        .Add "E16", slvEqual, 1#
        .Add "G18", slvLessThanEqual, 0.071
    End With
    
    oProblem.Solver.Method = slvGRG_Nonlinear
    
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.EnableEvents = False
    oProblem.Solver.UserCallbackMacroName = "ShowTrial"

    oProblem.Solver.SaveAllTrialSolutions = True
    
    oProblem.SolveIt
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("o2:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("o2")
    End If
End Sub

'this is the call signature for the callback. Can name the function by user preference.
'Must have the following three input arguments, as well as return a boolean value whether
'to stop (true) or continue (false)
Function ShowTrial(ByVal reason As Long, ByVal trialNum As Long, oProblem As SolvProblem) As Boolean
    Dim i As Long
    Dim stopSolver As Boolean
    
    If trialNum = 1 Then Debug.Print "Solver started on Worksheet: " & oProblem.SolverSheet.Name
    
    Debug.Print "Trial number: " & trialNum
    Debug.Print "Objective: " & oProblem.Objective.CellRange.value
    
    For i = 1 To oProblem.DecisionVars.Count
        Debug.Print oProblem.DecisionVars.CellRange(i).Address, oProblem.DecisionVars.CellRange(i).value
    Next i
    
    Debug.Print "Constraints Satisfied? " & oProblem.Constraints.AreSatisfied
    
    'decide whether to stop solver based on the reason for the event trigger
    Select Case reason
        Case SlvCallbackReason.slvShowIterations 'new iteration has completed or user hit esc key
            stopSolver = False
        Case SlvCallbackReason.slvMaxTimeLimit
            stopSolver = True 'if set to True then solver is stopped!
        Case SlvCallbackReason.slvMaxIterationsLimit
            stopSolver = False
        Case SlvCallbackReason.slvMaxSubproblemsLimit
            stopSolver = False
        Case SlvCallbackReason.slvMaxSolutionsLimit
            stopSolver = False
    End Select
    
    ShowTrial = stopSolver
End Function
