Attribute VB_Name = "test_Engineering_Design"
'@folder("SolverWrapper.Examples")

Option Explicit

'This example automates solving the problem in SOLVSAMP.XLS on the "Engineering Design" worksheet.
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
'This is a non-linear problem and is best solved with slvGRG_Nonlinear
Sub Solve_Engineering_Design()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Engineering Design")
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "G15", slvTargetValue, 0.09
    
    oProblem.DecisionVars.Add "G12"
    oProblem.DecisionVars.Initialize 100
    
    oProblem.Solver.Method = slvGRG_Nonlinear
    
    oProblem.Solver.Options.AssumeNonNeg = True
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True
    
    oProblem.SolveIt
    
    'leave no trace behind of SolverWrapper (hidden Solver names)
    oProblem.CleanUp
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub
