Attribute VB_Name = "test_Quick_Tour"
'@folder("SolverWrapper.Examples")

Option Explicit

'This example automates solving the problem in SOLVSAMP.XLS on the "Quick Tour" worksheet.
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

'slvGRG_Nonlinear converges and is fast
Sub Solve_Quick_Tour_Unconstrained()
    Dim Problem As SolvProblem
    Dim ws As Worksheet
    
    Set Problem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Quick Tour")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "F15", slvMaximize
    
    Problem.DecisionVars.Add "$B$11:$E$11"
    Problem.DecisionVars.Initialize 10000#
    
    Problem.Solver.Method = slvGRG_Nonlinear
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True

    Problem.SolveIt
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("m1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("m1"), keepOnlyValid:=True
    End If
End Sub

Sub Solve_Quick_Tour_Constrained()
    Dim Problem As SolvProblem
    Dim ws As Worksheet
    
    Set Problem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Quick Tour")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "F15", slvMaximize
    
    Problem.DecisionVars.Add "$B$11:$E$11"
    Problem.DecisionVars.Initialize 10000#
    
    Problem.Constraints.AddBounded "$B$11:$E$11", 0, 40000
    Problem.Constraints.Add "$F$11", slvLessThanEqual, 40000
    
    Problem.Solver.Method = slvGRG_Nonlinear
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True

    Problem.SolveIt
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("m1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("m1"), keepOnlyValid:=True
    End If
End Sub
