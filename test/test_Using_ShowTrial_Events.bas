Attribute VB_Name = "test_Using_ShowTrial_Events"
'@folder("SolverWrapper.Examples")

'This example automates solving the problem in SOLVSAMP.XLS on the "Portfolio of Securities" worksheet.
'
'SOLVSAMP.XLS is distributed with MS Office Excel and can be found in:
'
'Application.LibraryPath & "\..\SAMPLES\SOLVSAMP.XLS"
'
'which on some systems can be found here:
'C:\Program Files\Microsoft Office\root\Office16\SAMPLES\SOLVSAMP.XLS
'
'Import this module into the sample workbook, set a reference to the SolverWrapper code library
'and then save SOLVSAMP.XLS to SOLVSAMP.XLSM.

'To use SolverWrapper events, user must write their own event sink class (see example SolverEventSink class in test folder)
'Then connect to that class as in below example. Be sure to set EnableEvents of the Solver class to True.
Sub Solve_Portfolio_of_Securities_with_Events()
    Dim Problem As New SolvProblem
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Portfolio of Securities")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "E18", slvMaximize
    
    Problem.DecisionVars.Add "E10:E14"
    Problem.DecisionVars.Initialize 0.2, 0.2, 0.2, 0.2, 0.2
    
    Problem.Constraints.AddBounded "E10:E14", 0#, 1#
    Problem.Constraints.Add "E16", slvEqual, 1
    Problem.Constraints.Add "G18", slvLessThanEqual, 0.071
    
    Problem.Solver.Method = slvGRG_Nonlinear
    
    Problem.Solver.Options.AssumeNonNeg = False
    Problem.Solver.Options.RandomSeed = 7
    Problem.Solver.Options.StepThru = False
    Problem.Solver.Options.MaxTime = 1
    
    'must enable events to use this
    Problem.Solver.EnableEvents = True
    If Problem.Solver.EnableEvents Then
        'connect-up ShowTrial event proccessing class (optional!)
        Dim eventSink As SolverEventSink
        Set eventSink = New SolverEventSink
        Set eventSink.Problem = Problem
    End If
    
    Problem.Solver.SaveAllTrialSolutions = True
    
    Problem.SolveIt
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("q1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("q1")
    End If
End Sub
