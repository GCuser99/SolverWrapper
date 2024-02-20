Attribute VB_Name = "test_Shipping_Routes"
'@folder("SolverWrapper.Examples")

Option Explicit

'This example automates solving the problem in SOLVSAMP.XLS on the "Shipping Routes" worksheet.
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

'this is a linear problem so slvSimplex_LP is faster
'however, this problem has more than one solution - use slvGRG_Nonlinear for multiple solutions
Sub Solve_Shipping_Routes_Slower()
    Dim Problem As SolvProblem
    Dim ws As Worksheet
    
    Set Problem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Shipping Routes")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "B20", slvMinimize
    
    Problem.DecisionVars.Add "C8:G10"
    Problem.DecisionVars.Initialize 60
    
    Problem.Constraints.AddBounded "$C$8:$G$10", 0, 250
    Problem.Constraints.Add "$C$8:$G$10", slvInt
    Problem.Constraints.Add "$B$8:$B$10", slvLessThanEqual, "$B$16:$B$18"
    Problem.Constraints.Add "$C$12:$G$12", slvGreaterThanEqual, "$C$14:$G$14"
    
    Problem.Solver.Method = slvGRG_Nonlinear
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True

    Problem.SolveIt
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:ae10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub

Sub Solve_Shipping_Routes_Faster()
    Dim Problem As SolvProblem
    Dim ws As Worksheet
    
    Set Problem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Shipping Routes")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "B20", slvMinimize
    
    Problem.DecisionVars.Add "C8:G10"
    Problem.DecisionVars.Initialize 60
    
    Problem.Constraints.AddBounded "$C$8:$G$10", 0, 250
    Problem.Constraints.Add "$C$8:$G$10", slvInt
    Problem.Constraints.Add "$B$8:$B$10", slvLessThanEqual, "$B$16:$B$18"
    Problem.Constraints.Add "$C$12:$G$12", slvGreaterThanEqual, "$C$14:$G$14"
    
    Problem.Solver.Method = slvSimplex_LP
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True

    Problem.SolveIt
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:ae10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub
