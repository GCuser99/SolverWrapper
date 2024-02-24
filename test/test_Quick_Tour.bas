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

'Notes:
'slvGRG_Nonlinear converges and is fast
Sub Solve_Quick_Tour_Unconstrained()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Quick Tour")
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "F15", slvMaximize
    
    oProblem.DecisionVars.Add "$B$11:$E$11"
    oProblem.DecisionVars.Initialize 10000#
    
    oProblem.Solver.Method = slvGRG_Nonlinear
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True

    oProblem.SolveIt
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("m1:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("m1"), keepOnlyValid:=True
    End If
End Sub

Sub Solve_Quick_Tour_Constrained()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Quick Tour")
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "F15", slvMaximize
    
    oProblem.DecisionVars.Add "$B$11:$E$11"
    oProblem.DecisionVars.Initialize 10000#
    
    oProblem.Constraints.AddBounded "$B$11:$E$11", 0, 40000
    oProblem.Constraints.Add "$F$11", slvLessThanEqual, 40000
    
    oProblem.Solver.Method = slvGRG_Nonlinear
    oProblem.Solver.Options.RandomSeed = 7
    
    oProblem.Solver.SaveAllTrialSolutions = True

    oProblem.SolveIt
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("m1:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("m1"), keepOnlyValid:=True
    End If
End Sub
