Attribute VB_Name = "test_Engineering_Design"
'@folder("SolverWrapper.Examples")

'This example automates solving the problem in SOLVSAMP.XLS on the "Engineering Design" worksheet.
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

'This is a non-linear problem and is best solved with slvGRG_Nonlinear
Sub Solve_Engineering_Design()
    Dim Problem As New SolvProblem
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Engineering Design")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "G15", slvTargetValue, 0.09
    
    Problem.DecisionVars.Add "G12"
    Problem.DecisionVars.Initialize 100
    
    Problem.Solver.Method = slvGRG_Nonlinear
    
    Problem.Solver.Options.AssumeNonNeg = True
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True
    
    Problem.SolveIt
    
    'leave to trace of SolverWrapper (hidden Solver names) behind
    Problem.CleanUp
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub
