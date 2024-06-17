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

'Notes:
'this is a linear problem so slvSimplex_LP is faster
'however, this problem has more than one BEST solution - use slvGRG_Nonlinear to see multiple solutions
Sub Solve_Shipping_Routes_Slower()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Shipping Routes")
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "B20", slvMinimize
    
    oProblem.DecisionVars.Add "C8:G10"
    oProblem.DecisionVars.Initialize 60
    
    With oProblem.Constraints
        .AddBounded "$C$8:$G$10", 0, 250
        .Add "$C$8:$G$10", slvInt
        .Add "$B$8:$B$10", slvLessThanEqual, "$B$16:$B$18"
        .Add "$C$12:$G$12", slvGreaterThanEqual, "$C$14:$G$14"
    End With
    
    oProblem.Solver.Method = slvGRG_Nonlinear
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True

    oProblem.SolveIt
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:ae10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub

Sub Solve_Shipping_Routes_Faster()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Shipping Routes")
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "B20", slvMinimize
    
    oProblem.DecisionVars.Add "C8:G10"
    oProblem.DecisionVars.Initialize 60
    
    With oProblem.Constraints
        .AddBounded "$C$8:$G$10", 0, 250
        .Add "$C$8:$G$10", slvInt
        .Add "$B$8:$B$10", slvLessThanEqual, "$B$16:$B$18"
        .Add "$C$12:$G$12", slvGreaterThanEqual, "$C$14:$G$14"
    End With
    
    oProblem.Solver.Method = slvSimplex_LP
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True

    oProblem.SolveIt
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:ae10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub
