Attribute VB_Name = "test_Maximizing_Income"
'@folder("SolverWrapper.Examples")

Option Explicit

'This example automates solving the problem in SOLVSAMP.XLS on the "Maximizing Income" worksheet.
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

'Notes:
'this is a linear problem, hence slvSimplex_LP is best method
'but slvGRG_Nonlinear is almost as fast and accurate
Sub Solve_Maximizing_Income()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Maximizing Income")
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "H8", slvMaximize
    
    oProblem.DecisionVars.Add "B14:G14", "B15:B16", "E15"
    oProblem.DecisionVars.Initialize 50000
    
    With oProblem.Constraints
        .Add "B14:G14", slvGreaterThanEqual, 0
        .Add "B15:B16", slvGreaterThanEqual, 0
        .Add "E15", slvGreaterThanEqual, 0
        .Add "B18:H18", slvGreaterThanEqual, 100000
    End With
    
    oProblem.Solver.Method = slvSimplex_LP
    
    oProblem.Solver.Options.AssumeNonNeg = False
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True
    
    oProblem.SolveIt
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub

Sub Solve_Maximizing_Income_with_Optional_Constraint()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Maximizing Income")
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "H8", slvMaximize
    
    oProblem.DecisionVars.Add "B14:G14", "B15:B16", "E15"
    oProblem.DecisionVars.Initialize 50000
    
    With oProblem.Constraints
        .Add "B14:G14", slvGreaterThanEqual, 0
        .Add "B15:B16", slvGreaterThanEqual, 0
        .Add "E15", slvGreaterThanEqual, 0
        .Add "B18:H18", slvGreaterThanEqual, 100000
        'add optional constraint that the average maturity of the investments
        'held in month 1 should not be more 4 months
        .Add "B20", slvEqual, 0
    End With
    
    oProblem.Solver.Method = slvSimplex_LP
    
    oProblem.Solver.Options.AssumeNonNeg = False
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True
    
    oProblem.SolveIt
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub
