# SolverWrapper
MS Excel VBA Object Model and twinBASIC DLL for automating Solver

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

The Solver Add-in (from FrontLine Systems) that is installed with Microsoft Excel is a powerful tool for linear and non-linear spreadsheet model optimization. However, automating the Solver via VBA can be awkward due to Solver's non-OOP "functional" design, and the requirement that the Add-in must be installed (activated) before a VBA reference can be made to it (see Peltier Tech for [details](https://peltiertech.com/Excel/SolverVBA.html)).

This repo offers two solutions for automating Solver via VBA. One is an .xlsm macro file containing the SolverWrapper object model VBA code, and the other is by referencing an ActiveX DLL from VBA projects. The DLL, compiled in twinBASIC, can be called without registration as well if intellisense and Object Browser are not important. Both of these solutions control Solver by communicating directly with the SOLVER32.DLL, thus in effect circumventing the SOLVER.XLAM add-in, and eliminating having to check if the SOLVER.XLAM has been loaded into Excel. 

The SolverWrapper object model uses an OOP design, making it easier to understand and code with. Also, a few functional enhancements have been implemented, such as the capability to save intermediate trial solutions for further analysis of the results.

Be aware that one disadvantage of marshalling communication direcly with the Solver DLL (as opposed to the Solver Add-in) is that Solver Report creation is lost. This is because those reports were created by the SOLVER.XLAM Add-in, not the DLL. 

Credits

[RubberDuck](https://rubberduckvba.com/) by Mathieu Guindon

[twinBASIC](https://twinbasic.com/preview.html) by Wayne Phillips

[Inno Setup](https://jrsoftware.org/isinfo.php) by Jordan Russell and [UninsIS](https://github.com/Bill-Stewart/UninsIS) by Bill Stewart


 

   
