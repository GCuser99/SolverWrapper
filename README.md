# SolverWrapper
MS Excel VBA Object Model and [twinBASIC](https://twinbasic.com/preview.html) DLL for automating Solver

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Introduction

The Solver Add-in (from [FrontLine Systems](https://www.solver.com/)) that is installed with Microsoft Excel is a powerful tool for linear and non-linear spreadsheet model optimization. However, automating the Solver via VBA can be awkward due to Solver's non-OOP "functional" design, and the requirement that the Add-in must be installed (activated) before a VBA reference can be made to it (see Peltier Tech for [details](https://peltiertech.com/Excel/SolverVBA.html)).

This repo offers two compatible solutions for automating Solver via VBA. One consists of SolverWrapper object model in pure VBA code, and the other is an ActiveX DLL referenced from within your VBA projects. The DLL, compiled in [twinBASIC](https://twinbasic.com/preview.html), can either be [installed/registered](https://github.com/GCuser99/SolverWrapper/tree/main/dist), or be called without registration if use of IntelliSense and the Object Browser are not important. These solutions control Solver by communicating directly with the SOLVER32.DLL, thus in effect circumventing the SOLVER.XLAM add-in, and eliminating having to insure that the SOLVER.XLAM has been loaded into Excel. 

## Features

- Uses an OOP design, making it easier to understand and code with
- Unique design that communicates directly with SOLVER32.DLL
- Can be implemented as a pure VBA code library or [twinBASIC](https://twinbasic.com/preview.html) ActiveX DLL object model
- Capability to save intermediate trial solutions, as opposed to one BEST solution (often there are more than one!)
- Enhanced Solver callback protocol
- An alternative event-based means of monitoring solution progress versus the callback
- Other miscellaneous enhancements
- Help documentation is available in the [SolverWrapper Wiki](https://github.com/GCuser99/SolverWrapper/wiki)

Be aware that one disadvantage of marshalling communication directly with the Solver DLL (as opposed to the Solver Add-in) is that Solver Report creation is lost. This is because those reports were created by the Add-in, not the DLL.

## Example

The following example automates solving the problem in SOLVSAMP.XLS on the "Portfolio of Securities" worksheet which is distributed with MS Office Excel and can usually be found in "C:\Program Files\Microsoft Office\root\Office16\SAMPLES".

```vba
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
    
    'save all valid solutions to the worksheet
    If Problem.Solver.SaveAllTrialSolutions Then
        ws.Range("q1:az10000").ClearContents
        Problem.SaveSolutionsToRange ws.Range("q1"), keepOnlyValid:=True
    End If
End Sub
```
<img src="https://github.com/GCuser99/SolverWrapper/blob/main/dev/images/portfolio_of_securities.png" alt="EngineeringDesign" width=100% height=100%>

## Requirements:

- 64-bit MS Windows
- MS Office 2010 or later, 32/64-bit

## Acknowledgements

[RubberDuck](https://rubberduckvba.com/) by Mathieu Guindon

[twinBASIC](https://twinbasic.com/preview.html) by Wayne Phillips

[Inno Setup](https://jrsoftware.org/isinfo.php) by Jordan Russell and [UninsIS](https://github.com/Bill-Stewart/UninsIS) by Bill Stewart


 

   
