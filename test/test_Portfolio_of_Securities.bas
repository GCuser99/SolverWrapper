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
Sub Solve_Portfolio_of_Securities()
    Dim Problem As New SolvProblem
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Portfolio of Securities")
    
    'initialize the problem by passing a reference to the worksheet of interest
    Problem.Initialize ws
    
    'define the objective cell to be optimized
    Problem.Objective.Define "E18", slvMaximize
    
    'define and initialize the decision cell(s)
    Problem.DecisionVars.Add "E10:E14"
    Problem.DecisionVars.Initialize 0.2, 0.2, 0.2, 0.2, 0.2
    
    'add some constraints
    Problem.Constraints.AddBounded "E10:E14", 0#, 1#
    Problem.Constraints.Add "E16", slvEqual, 1#
    Problem.Constraints.Add "G18", slvLessThanEqual, 0.071
    
    'set the solver engine to use
    Problem.Solver.Method = slvGRG_Nonlinear
    
    'set some solver options
    Problem.Solver.Options.AssumeNonNeg = True
    Problem.Solver.Options.RandomSeed = 7
    
    Problem.Solver.SaveAllTrialSolutions = True

    'solve the optimization problem
    Problem.SolveIt
    
    'save all trial solutions that passed the constraints to the worksheet
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("o2:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("o2"), keepOnlyValid:=True
    End If
End Sub
