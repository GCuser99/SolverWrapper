# SolverWrapper
MS Excel VBA Object Model and [twinBASIC](https://twinbasic.com/preview.html) DLL for automating Solver

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

The Solver Add-in (from [FrontLine Systems](https://www.solver.com/)) that is installed with Microsoft Excel is a powerful tool for linear and non-linear spreadsheet model optimization. However, automating the Solver via VBA can be awkward due to Solver's non-OOP "functional" design, and the requirement that the Add-in must be installed (activated) before a VBA reference can be made to it (see Peltier Tech for [details](https://peltiertech.com/Excel/SolverVBA.html)).

This repo offers two compatible solutions for automating Solver via VBA. One consists of SolverWrapper object model in VBA code, and the other is by referencing an ActiveX DLL from VBA projects. The DLL, compiled in [twinBASIC](https://twinbasic.com/preview.html), can either be installed/registered and referenced within your Excel VBA project, or be called without registration as well if intellisense and Object Browser are not important. Both of these solutions control Solver by communicating directly with the SOLVER32.DLL, thus in effect circumventing the SOLVER.XLAM add-in, and eliminating having to check if the SOLVER.XLAM has been loaded into Excel. 

## Features

- Uses an OOP design, making it easier to understand and code with
- Unique design that communicates directly with SOLVER32.DLL
- Can be impelemented as VBA library or ActiveX DLL object model
- Capability to save intermediate trial solutions, as opposed to one BEST solution (often there are more than one!)
- Enhanced Solver callback protocol
- An alternative event-based means of monitoring solution progress versus the callback
- Other miscellaneous enhancements
- Help documentation is available in the [SolverWrapper Wiki](https://github.com/GCuser99/SolverWrapper/wiki)

Be aware that one disadvantage of marshalling communication direcly with the Solver DLL (as opposed to the Solver Add-in) is that Solver Report creation is lost. This is because those reports were created by the SOLVER.XLAM Add-in, not the DLL.

## Examples

```vba
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
    
    'leave no trace of SolverWrapper (hidden Solver names) behind
    Problem.CleanUp
    
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("o1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("o1"), keepOnlyValid:=True
    End If
End Sub
```

Credits

[RubberDuck](https://rubberduckvba.com/) by Mathieu Guindon

[twinBASIC](https://twinbasic.com/preview.html) by Wayne Phillips

[Inno Setup](https://jrsoftware.org/isinfo.php) by Jordan Russell and [UninsIS](https://github.com/Bill-Stewart/UninsIS) by Bill Stewart


 

   
