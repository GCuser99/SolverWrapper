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

'Notes:
'slvGRG_Nonlinear finds the optimum fast
'there are (at least) two solutions that are nearly tied for non-linear case
Sub Solve_Product_Mix_Non_Linear()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Product Mix")
    'this makes the problem non-linear
    ws.Range("$H$15").value = 0.9
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "D18", slvMaximize
    
    oProblem.DecisionVars.Add "$D$9:$F$9"
    oProblem.DecisionVars.Initialize 100
    
    With oProblem.Constraints
        .AddBounded "$D$9:$F$9", 0, 800
        .Add "$C$11:$C$15", slvLessThanEqual, "$B$11:$B$15"
        .Add "$D$9:$F$9", slvInt
    End With
    
    oProblem.Solver.Method = slvGRG_Nonlinear
    
    oProblem.Solver.Options.MaxTimeNoImp = 2
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True

    oProblem.SolveIt

    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("n1:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("n1"), keepOnlyValid:=True
    End If
End Sub

Sub Solve_Product_Mix_Linear()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Product Mix")
    'this makes the problem linear
    ws.Range("$H$15").value = 1#
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "D18", slvMaximize
    
    oProblem.DecisionVars.Add "$D$9:$F$9"
    oProblem.DecisionVars.Initialize 100
    
    With oProblem.Constraints
        .AddBounded "$D$9:$F$9", 0, 800
        .Add "$C$11:$C$15", slvLessThanEqual, "$B$11:$B$15"
        .Add "$D$9:$F$9", slvInt
    End With
    
    oProblem.Solver.Method = slvSimplex_LP
    
    oProblem.Solver.Options.MaxTimeNoImp = 2
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True

    oProblem.SolveIt

    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("n1:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("n1"), keepOnlyValid:=True
    End If
End Sub

