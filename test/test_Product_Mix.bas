Attribute VB_Name = "test_Product_Mix"
'@folder("SolverWrapper.Examples")

Option Explicit

'This example automates solving the problem in SOLVSAMP.XLS on the "Product Mix" worksheet.
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

'slvGRG_Nonlinear finds the optimum fast
Sub Solve_Product_Mix_Non_Linear()
    Dim Problem As SolvProblem
    Dim ws As Worksheet
    
    Set Problem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Product Mix")
    'this makes the problem non-linear
    ws.Range("$H$15").value = 0.9
    
    Problem.Initialize ws
    
    Problem.Objective.Define "D18", slvMaximize
    
    Problem.DecisionVars.Add "$D$9:$F$9"
    Problem.DecisionVars.Initialize 100
    
    Problem.Constraints.AddBounded "$D$9:$F$9", 0, 800
    Problem.Constraints.Add "$C$11:$C$15", slvLessThanEqual, "$B$11:$B$15"
    Problem.Constraints.Add "$D$9:$F$9", slvInt
    
    Problem.Solver.Method = slvGRG_Nonlinear
    
    Problem.Solver.Options.MaxTimeNoImp = 2
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True

    Problem.SolveIt

    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("n1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("n1"), keepOnlyValid:=True
    End If
End Sub

Sub Solve_Product_Mix_Linear()
    Dim Problem As SolvProblem
    Dim ws As Worksheet
    
    Set Problem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Product Mix")
    'this makes the problem linear
    ws.Range("$H$15").value = 1#
    
    Problem.Initialize ws
    
    Problem.Objective.Define "D18", slvMaximize
    
    Problem.DecisionVars.Add "$D$9:$F$9"
    Problem.DecisionVars.Initialize 100
    
    Problem.Constraints.AddBounded "$D$9:$F$9", 0, 800
    Problem.Constraints.Add "$C$11:$C$15", slvLessThanEqual, "$B$11:$B$15"
    Problem.Constraints.Add "$D$9:$F$9", slvInt
    
    Problem.Solver.Method = slvSimplex_LP
    
    Problem.Solver.Options.MaxTimeNoImp = 2
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True

    Problem.SolveIt

    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("n1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("n1"), keepOnlyValid:=True
    End If
End Sub

