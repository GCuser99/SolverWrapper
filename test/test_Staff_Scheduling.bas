Attribute VB_Name = "test_Staff_Scheduling"
'@folder("SolverWrapper.Examples")

'This example automates solving the problem in SOLVSAMP.XLS on the "Staff Scheduling" worksheet.
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

'this is a linear problem so slvSimplex_LP conversges the fastest
'but this problem is interesting because there are many tied solutions that
'you would not know without looking at all of solutions tried
Sub Solve_Staff_Scheduling()
    Dim Problem As New SolvProblem
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Staff Scheduling")
    
    Problem.Initialize ws
    
    Problem.Objective.Define "D20", slvMinimize
    
    Problem.DecisionVars.Add "D7:D13"
    Problem.DecisionVars.Initialize 4
    
    Problem.Constraints.AddBounded "D7:D13", 0, 15
    Problem.Constraints.Add "D7:D13", slvInt
    Problem.Constraints.Add "F15:L15", slvGreaterThanEqual, "$F$17:$L$17"

    Problem.Solver.Method = slvSimplex_LP
    
    Problem.Solver.Options.RandomSeed = 7
    Problem.Solver.Options.Precision = 0.000001

    Problem.Solver.SaveAllTrialSolutions = True

    Problem.SolveIt
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("s1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("s1"), keepOnlyValid:=True
    End If
End Sub
