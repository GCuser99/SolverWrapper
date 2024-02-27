# SolverWrapper
VBA/[twinBASIC](https://twinbasic.com/preview.html) Object Models for automating MS Excel's Solver

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Introduction

The SOLVER Add-in (from [FrontLine Systems](https://www.solver.com/)) that comes installed with Microsoft Excel is a powerful tool for linear and non-linear spreadsheet model optimization. However, automating the Solver via VBA can be awkward due to Solver's non-OOP "functional" design, and the requirement that the Add-in must be installed (activated) before a VBA reference can be made to it (see Peltier Tech for [details](https://peltiertech.com/Excel/SolverVBA.html)).

This repo offers two compatible solutions for automating Solver via VBA. One consists of SolverWrapper object model in VBA code, and the other is an ActiveX DLL referenced from within your VBA projects. The DLL, compiled in [twinBASIC](https://twinbasic.com/preview.html), can either be [installed/registered](https://github.com/GCuser99/SolverWrapper/tree/main/dist), or be called without registration if the use of IntelliSense and the Object Browser are not important. 

These unique solutions control Solver by **communicating directly with the SOLVER32.DLL**, thus circumventing the SOLVER Add-in, and eliminating having to ensure that the Add-in has been loaded into Excel. Another advantage is that we can rewrite the user-DLL interface to add improvements such as an enhanced Solver callback protocol, and an event-model for monitoring optimization progress.

## Features

- OOP design, making it easier to understand and code with
- Unique implementation that communicates directly with SOLVER32.DLL (bypassing SOLVER Add-in)
- Can either be implemented as a VBA code library or [twinBASIC](https://twinbasic.com/preview.html) ActiveX DLL object model
- Capability to save intermediate trial solutions, as opposed to just one BEST solution (often more than one exists!)
- Enhanced Solver [callback protocol](https://github.com/GCuser99/SolverWrapper/wiki#using-the-enhanced-callback)
- Alternative [event-based model](https://github.com/GCuser99/SolverWrapper/wiki#using-solverwrapper-events) for monitoring solution progress versus using the callback
- Other miscellaneous enhancements
- Help documentation is available in the [SolverWrapper Wiki](https://github.com/GCuser99/SolverWrapper/wiki)

Be aware that one disadvantage of marshalling communication directly with the Solver DLL (as opposed to the Solver Add-in) is that Solver Report creation is lost. This is because those reports were created by the Add-in, not the DLL.

## Example

The following example automates solving the problem in SOLVSAMP.XLS on the "Portfolio of Securities" worksheet which is distributed with MS Office Excel and can usually be found in "C:\Program Files\Microsoft Office\root\Office16\SAMPLES".

```vba
Sub Solve_Portfolio_of_Securities()
    Dim oProblem As SolvProblem
    Dim ws As Worksheet

    Set oProblem = New SolvProblem
    
    Set ws = ThisWorkbook.Worksheets("Portfolio of Securities")
    
    'initialize the problem by passing a reference to the worksheet of interest
    oProblem.Initialize ws
    
    'define the objective cell to be optimized
    oProblem.Objective.Define "E18", slvMaximize
    
    'add and initialize the decision cell(s)
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
```
The image below shows the result of running the above optimization procedure.
<img src="https://github.com/GCuser99/SolverWrapper/blob/main/dev/images/portfolio_of_securities.png" alt="EngineeringDesign" width=100% height=100%>

## Requirements:

- 64-bit MS Windows
- MS Office 2010 or later, 32/64-bit

## Acknowledgements

[RubberDuck](https://rubberduckvba.com/) by Mathieu Guindon

[twinBASIC](https://twinbasic.com/preview.html) by Wayne Phillips

[Inno Setup](https://jrsoftware.org/isinfo.php) by Jordan Russell and [UninsIS](https://github.com/Bill-Stewart/UninsIS) by Bill Stewart


 

   
