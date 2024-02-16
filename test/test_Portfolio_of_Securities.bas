Attribute VB_Name = "test_Portfolio_of_Securities"
'@folder("SolverWrapper.Examples")

'This example automates solving the problem in SOLVSAMP.XLS on the "Portfolio of Securities" worksheet.
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

'This is a non-linear problem so cannot use use slvSimplex_LP
'so use slvGRG_Nonlinear
'slvEvolutionary does not converge (fails on E16=1)
Sub Solve_Portfolio_of_Securities()
    Dim Problem As New SolvProblem
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Portfolio of Securities")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "E18", slvMaximize
    
    Problem.DecisionVars.Add "E10:E14"
    Problem.DecisionVars.Initialize 0.2, 0.2, 0.2, 0.2, 0.2
    
    Problem.Constraints.AddBounded "E10:E14", 0#, 1#
    Problem.Constraints.Add "E16", slvEqual, 1
    Problem.Constraints.Add "G18", slvLessThanEqual, 0.071
    
    Problem.Solver.Method = slvGRG_Nonlinear
    
    Problem.Solver.Options.AssumeNonNeg = False
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True

    Problem.SolveIt
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("q1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("q1"), keepOnlyValid:=True
    End If
End Sub
