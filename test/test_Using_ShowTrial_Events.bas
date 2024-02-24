Attribute VB_Name = "test_Using_ShowTrial_Events"
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

'To use SolverWrapper events, user must write their own event sink class (see example SolverEventSink class in test folder)
'Then connect to that class as in below example. Be sure to set EnableEvents of the Solver class to True.
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
    
    With oProblem.Solver.Options
        .AssumeNonNeg = False
        .RandomSeed = 7
        .MaxTime = 1
    End With
    
    'must enable events to use this
    oProblem.Solver.EnableEvents = True
    If oProblem.Solver.EnableEvents Then
        'connect-up events proccessing class
        Dim eventSink As SolverEventSink
        Set eventSink = New SolverEventSink
        Set eventSink.Problem = oProblem
    End If
    
    oProblem.Solver.SaveAllTrialSolutions = True
    
    oProblem.SolveIt
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("o2:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("o2")
    End If
End Sub
