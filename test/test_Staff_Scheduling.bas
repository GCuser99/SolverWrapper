Attribute VB_Name = "test_Staff_Scheduling"
'@folder("SolverWrapper.Examples")

Option Explicit

'This example automates solving the problem in SOLVSAMP.XLS on the "Staff Scheduling" worksheet.
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
'this is a linear problem so slvSimplex_LP converges the fastest
'but this problem is interesting because there are many tied solutions that
'you would not know without looking at all of solutions tried
Sub Solve_Staff_Scheduling()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet
    
    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Staff Scheduling")
    
    oProblem.Initialize ws
    
    oProblem.Objective.Define "D20", slvMinimize
    
    oProblem.DecisionVars.Add "D7:D13"
    oProblem.DecisionVars.Initialize 4
    
    With oProblem.Constraints
        .AddBounded "D7:D13", 0, 15
        .Add "D7:D13", slvInt
        .Add "F15:L15", slvGreaterThanEqual, "$F$17:$L$17"
    End With

    oProblem.Solver.Method = slvSimplex_LP
    
    oProblem.Solver.Options.RandomSeed = 7
    oProblem.Solver.Options.Precision = 0.000001

    oProblem.Solver.SaveAllTrialSolutions = True

    oProblem.SolveIt
    
    If oProblem.Solver.SaveAllTrialSolutions Then
        ws.Range("s1:az10000").ClearContents
        oProblem.SaveSolutionsToRange ws.Range("s1"), keepOnlyValid:=True
    End If
End Sub
