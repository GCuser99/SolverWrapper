Attribute VB_Name = "test_Portfolio_of_Securities"
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

'Notes:
'This is a non-linear problem - use slvGRG_Nonlinear
Sub Solve_Portfolio_of_Securities()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Portfolio of Securities")
    
    'initialize the problem by passing a reference to the worksheet of interest
    oProblem.Initialize ws
    
    'define the objective cell to be optimized
    oProblem.Objective.Define "E18", slvMaximize
    
    'define and initialize the decision cell(s)
    oProblem.DecisionVars.Add "E10:E14"
    oProblem.DecisionVars.Initialize 0.2, 0.2, 0.2, 0.2, 0.2
    
    'add some constraints
    With oProblem.Constraints
        .AddBounded "E10:E14", 0#, 1#
        .Add "E16", slvEqual, 1#
        .Add "G18", slvLessThanEqual, 0.071
    End With
    
    'set the solver engine to use
    oProblem.Solver.Method = slvGRG_Nonlinear
    
    'set solver option
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True

    'solve the optimization problem
    oProblem.SolveIt
    
    'save all trial solutions that passed the constraints to the worksheet
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("o2:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("o2"), keepOnlyValid:=True
    End If
End Sub
