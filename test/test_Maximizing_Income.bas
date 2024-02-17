Attribute VB_Name = "test_Maximizing_Income"
'@folder("SolverWrapper.Examples")

'This example automates solving the problem in SOLVSAMP.XLS on the "Maximizing Income" worksheet.
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

'this is a linear problem, hence slvSimplex_LP is best method
'but slvGRG_Nonlinear is almost as fast and accurate
Sub Solve_Maximizing_Income()
    Dim Problem As New SolvProblem
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Maximizing Income")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "H8", slvMaximize
    
    Problem.DecisionVars.Add "B14:G14", "B15:B16", "E15"
    Problem.DecisionVars.Initialize 50000
    
    Problem.Constraints.Add "B14:G14", slvGreaterThanEqual, 0
    Problem.Constraints.Add "B15:B16", slvGreaterThanEqual, 0
    Problem.Constraints.Add "E15", slvGreaterThanEqual, 0
    Problem.Constraints.Add "B18:H18", slvGreaterThanEqual, 100000
    
    Problem.Solver.Method = slvSimplex_LP
    
    Problem.Solver.Options.AssumeNonNeg = False
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True
    
    Problem.SolveIt
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub

Sub Solve_Maximizing_Income_with_Optional_Constraint()
    Dim Problem As New SolvProblem
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Maximizing Income")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "H8", slvMaximize
    
    Problem.DecisionVars.Add "B14:G14", "B15:B16", "E15"
    Problem.DecisionVars.Initialize 50000
    
    Problem.Constraints.Add "B14:G14", slvGreaterThanEqual, 0
    Problem.Constraints.Add "B15:B16", slvGreaterThanEqual, 0
    Problem.Constraints.Add "E15", slvGreaterThanEqual, 0
    Problem.Constraints.Add "B18:H18", slvGreaterThanEqual, 100000
    
    'add optional constraint that the average maturity of the investments
    'held in month 1 should not be more 4 months
    Problem.Constraints.Add "B20", slvEqual, 0
    
    Problem.Solver.Method = slvSimplex_LP
    
    Problem.Solver.Options.AssumeNonNeg = False
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True
    
    Problem.SolveIt
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub
