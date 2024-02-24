VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolvOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class to set Solver Options."
'%ModuleDescription "A class to set Solver Options."
'@Exposed
'@folder("SolverWrapper.Source")
' ==========================================================================
' SolverWrapper v0.8
'
' A wrapper for automating MS Excel's Solver Add-in
'
' https://github.com/GCuser99/SolverWrapper
'
' Contact Info:
'
' https://github.com/GCUser99
' ==========================================================================
' MIT License
'
' Copyright (c) 2024, GCUser99 (https://github.com/GCuser99/SolverWrapper)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==========================================================================
Option Explicit

Private oSolverSheet As Worksheet

' ==========================================================================
' Public members
' ==========================================================================

'%Description("Resets all Option values to defaults.")
Public Sub Reset()
Attribute Reset.VB_Description = "Resets all Option values to defaults."
    DeleteSolverNames oSolverSheet, False, True
    InitSolver oSolverSheet, False, True
End Sub

'%Description("The maximum amount of time (in seconds) Solver will spend solving the problem. The value must be a positive integer.")
Public Property Let MaxTime(ByVal val As Long)
Attribute MaxTime.VB_Description = "The maximum amount of time (in seconds) Solver will spend solving the problem. The value must be a positive integer."
    With oSolverSheet.Names
        If val <> 0 Then
           .Add Name:="solver_tim", RefersTo:=CDbl(val), Visible:=nameVisible
        Else
           .Add Name:="solver_tim", RefersTo:=MAXLONG, Visible:=nameVisible
        End If
    End With
End Property
Public Property Get MaxTime() As Long
    On Error Resume Next
    MaxTime = NameToLong("solver_tim", oSolverSheet)
End Property

'%Description("The maximum number of iterations Solver will use in solving the problem. The value must be a positive integer.")
Public Property Let MaxIterations(ByVal val As Long)
Attribute MaxIterations.VB_Description = "The maximum number of iterations Solver will use in solving the problem. The value must be a positive integer."
    With oSolverSheet.Names
        If MaxIterations <> 0 Then
            .Add Name:="solver_itr", RefersTo:=CDbl(val), Visible:=nameVisible
        Else
            .Add Name:="solver_itr", RefersTo:=MAXLONG, Visible:=nameVisible
        End If
    End With
End Property
Public Property Get MaxIterations() As Long
    On Error Resume Next
    MaxIterations = NameToLong("solver_itr", oSolverSheet)
End Property

'%Description("A number between 0 (zero) and 1 that specifies the precision with which constraints (including integer constraints) must be satisfied. The default precision is 0.000001. A smaller number of decimal places (for example, 0.0001) indicates a lower degree of precision. In general, the higher the degree of precision you specify (the smaller the number), the more time Solver will take to reach solutions.")
Public Property Let Precision(ByVal val As Double)
Attribute Precision.VB_Description = "A number between 0 (zero) and 1 that specifies the precision with which constraints (including integer constraints) must be satisfied. The default precision is 0.000001. A smaller number of decimal places (for example, 0.0001) indicates a lower degree of precision. In general, the higher the degree of precision you specify (the smaller the number), the more time Solver will take to reach solutions."
    With oSolverSheet.Names
        .Add Name:="solver_pre", RefersToLocal:=val, Visible:=nameVisible
    End With
End Property
Public Property Get Precision() As Double
    On Error Resume Next
    Precision = NameToDbl("solver_pre", oSolverSheet)
End Property

'%Description("True to have Solver pause on each trial solution. The default value is False.")
Public Property Let StepThru(ByVal val As Boolean)
Attribute StepThru.VB_Description = "True to have Solver pause on each trial solution. The default value is False."
    With oSolverSheet.Names
        If val Then
           .Add Name:="solver_sho", RefersToLocal:=1, Visible:=nameVisible
        Else
           .Add Name:="solver_sho", RefersToLocal:=2, Visible:=nameVisible
        End If
    End With
End Property
Public Property Get StepThru() As Boolean
    On Error Resume Next
    StepThru = (NameToLong("solver_sho", oSolverSheet) = 1)
End Property

'%Description("Specifies the approach used to obtain initial estimates of the basic variables in each one-dimensional search: 1 represents tangent estimates, and 2 represents quadratic estimates. Tangent estimates use linear extrapolation from a tangent vector. Quadratic estimates use quadratic extrapolation; this may improve the results for highly nonlinear problems. The default value is 1 (tangent estimates).")
Public Property Let Estimates(ByVal val As SlvEstimates)
Attribute Estimates.VB_Description = "Specifies the approach used to obtain initial estimates of the basic variables in each one-dimensional search: 1 represents tangent estimates, and 2 represents quadratic estimates. Tangent estimates use linear extrapolation from a tangent vector. Quadratic estimates use quadratic extrapolation; this may improve the results for highly nonlinear problems. The default value is 1 (tangent estimates)."
    With oSolverSheet.Names
        .Add Name:="solver_est", RefersToLocal:=val, Visible:=nameVisible
    End With
End Property
Public Property Get Estimates() As SlvEstimates
    On Error Resume Next
    Estimates = NameToLong("solver_est", oSolverSheet)
End Property

'%Description("Specifies forward differencing or central differencing for estimates of partial derivatives of the objective and constraint functions: 1 represents forward differencing, and 2 represents central differencing. Central differencing requires more worksheet recalculations, but it may help with problems that generate a message saying that Solver could not improve the solution. With constraints whose values change rapidly near their limits, you should use central differencing. The default value is 1 (forward differencing).")
Public Property Let Derivatives(ByVal val As SlvDerivatives)
Attribute Derivatives.VB_Description = "Specifies forward differencing or central differencing for estimates of partial derivatives of the objective and constraint functions: 1 represents forward differencing, and 2 represents central differencing. Central differencing requires more worksheet recalculations, but it may help with problems that generate a message saying that Solver could not improve the solution. With constraints whose values change rapidly near their limits, you should use central differencing. The default value is 1 (forward differencing)."
    With oSolverSheet.Names
        .Add Name:="solver_drv", RefersToLocal:=val, Visible:=nameVisible
    End With
End Property
Public Property Get Derivatives() As SlvDerivatives
    On Error Resume Next
    Derivatives = NameToLong("solver_drv", oSolverSheet)
End Property

'%Description("Use the Search options to specify the search algorithm that will be used at each iteration to decide which direction to search in: 1 represents the Newton search method, and 2 represents the conjugate search method. Newton, which uses a quasi-Newton method, is the default search method.")
Public Property Let SearchOption(ByVal val As SlvSearchOption)
Attribute SearchOption.VB_Description = "Use the Search options to specify the search algorithm that will be used at each iteration to decide which direction to search in: 1 represents the Newton search method, and 2 represents the conjugate search method. Newton, which uses a quasi-Newton method, is the default search method."
     With oSolverSheet.Names
        .Add Name:="solver_nwt", RefersToLocal:=val, Visible:=nameVisible
    End With
End Property
Public Property Get SearchOption() As SlvSearchOption
    On Error Resume Next
    SearchOption = NameToLong("solver_nwt", oSolverSheet)
End Property

'%Description("A decimal number between 0 (zero) and 100 that specifies the Integer Optimality percentage tolerance. This argument applies only if integer constraints have been defined; it specifies that Solver can stop if it has found a feasible integer solution whose objective is within this percentage of the best known bound on the objective of the true integer optimal solution.  A larger percentage tolerance would tend to speed up the solution process.")
Public Property Let IntTolerance(ByVal val As Double)
Attribute IntTolerance.VB_Description = "A decimal number between 0 (zero) and 100 that specifies the Integer Optimality percentage tolerance. This argument applies only if integer constraints have been defined; it specifies that Solver can stop if it has found a feasible integer solution whose objective is within this percentage of the best known bound on the objective of the true integer optimal solution.  A larger percentage tolerance would tend to speed up the solution process."
    With oSolverSheet.Names
        .Add Name:="solver_tol", RefersToLocal:=(val / 100#), Visible:=nameVisible
    End With
End Property
Public Property Get IntTolerance() As Double
    On Error Resume Next
    IntTolerance = NameToDbl("solver_tol", oSolverSheet)
End Property

'%Description("If the objective or constraints differ by several orders of magnitude - for example, maximizing percentage of profit based on million-dollar investments - set this option True to have Solver internally rescale the objective and constraint values to similar orders of magnitude during computation. If this option is False, Solver will perform its computations with the original values of the objective and constraints. The default value is True.")
Public Property Let Scaling(ByVal val As Boolean)
Attribute Scaling.VB_Description = "If the objective or constraints differ by several orders of magnitude - for example, maximizing percentage of profit based on million-dollar investments - set this option True to have Solver internally rescale the objective and constraint values to similar orders of magnitude during computation. If this option is False, Solver will perform its computations with the original values of the objective and constraints. The default value is True."
    With oSolverSheet.Names
        If val Then
           .Add Name:="solver_scl", RefersToLocal:=1, Visible:=nameVisible
        Else
           .Add Name:="solver_scl", RefersToLocal:=0, Visible:=nameVisible
        End If
    End With
End Property
Public Property Get Scaling() As Boolean
    On Error Resume Next
    Scaling = (NameToLong("solver_scl", oSolverSheet) = 1)
End Property

'%Description("A number between 0 (zero) and 1 that specifies the convergence tolerance for the GRG Nonlinear Solving and Evolutionary Solving methods.  For the GRG method, when the relative change in the target cell value is less than this tolerance for the last five iterations, Solver stops.  For the Evolutionary method, when 99% or more of the members of the population have fitness values whose relative (i.e. percentage) difference is less than this tolerance, Solver stops. In both cases, Solver displays the message Solver converged to the current solution. All constraints are satisfied.")
Public Property Let Convergence(ByVal val As Double)
Attribute Convergence.VB_Description = "A number between 0 (zero) and 1 that specifies the convergence tolerance for the GRG Nonlinear Solving and Evolutionary Solving methods.  For the GRG method, when the relative change in the target cell value is less than this tolerance for the last five iterations, Solver stops.  For the Evolutionary method, when 99% or more of the members of the population have fitness values whose relative (i.e. percentage) difference is less than this tolerance, Solver stops. In both cases, Solver displays the message Solver converged to the current solution. All constraints are satisfied."
    With oSolverSheet.Names
        .Add Name:="solver_cvg", RefersToLocal:=val, Visible:=nameVisible
    End With
End Property
Public Property Get Convergence() As Double
    On Error Resume Next
    Convergence = NameToDbl("solver_cvg", oSolverSheet)
End Property

'%Description("True to have Solver assume a lower limit of zero for all decision variable cells that do not have explicit lower limits Constraints (the cells must contain nonnegative values). False to have Solver use only the limits specified in the Constraints.")
Public Property Let AssumeNonNeg(ByVal val As Boolean)
Attribute AssumeNonNeg.VB_Description = "True to have Solver assume a lower limit of zero for all decision variable cells that do not have explicit lower limits Constraints (the cells must contain nonnegative values). False to have Solver use only the limits specified in the Constraints."
    'this option is actually a Constraint!
    With oSolverSheet.Names
        If val Then
           .Add Name:="solver_neg", RefersToLocal:=1, Visible:=nameVisible 'True is default
        Else
           .Add Name:="solver_neg", RefersToLocal:=2, Visible:=nameVisible
        End If
     End With
End Property
Public Property Get AssumeNonNeg() As Boolean
    On Error Resume Next
    AssumeNonNeg = (NameToLong("solver_neg", oSolverSheet) = 1)
End Property

'%Description("The number of different starting points (values for the decision variables) you want the multistart method to consider. The minimum population size is 10; if you supply a value less than 10, or leave it blank, the multistart method uses a population size of 10 times the number of decision variables, but no more than 200.")
Public Property Let PopulationSize(ByVal val As Long)
Attribute PopulationSize.VB_Description = "The number of different starting points (values for the decision variables) you want the multistart method to consider. The minimum population size is 10; if you supply a value less than 10, or leave it blank, the multistart method uses a population size of 10 times the number of decision variables, but no more than 200."
    With oSolverSheet.Names
        .Add Name:="solver_ssz", RefersToLocal:=val, Visible:=nameVisible
    End With
End Property
Public Property Get PopulationSize() As Long
    On Error Resume Next
    PopulationSize = NameToLong("solver_ssz", oSolverSheet)
End Property

'%Description("A positive integer specifies a fixed seed for the random number generator used by the Evolutionary Solving method and the GRG Nonlinear multistart method for global optimization.  This means that Solver will find the same solution each time it is run on a model that has not changed.  A zero value specifies that Solver should use a different seed for the random number generator each time it runs, which may yield different solutions each time it is run on a model that has not changed.")
Public Property Let RandomSeed(ByVal val As Long)
Attribute RandomSeed.VB_Description = "A positive integer specifies a fixed seed for the random number generator used by the Evolutionary Solving method and the GRG Nonlinear multistart method for global optimization.  This means that Solver will find the same solution each time it is run on a model that has not changed.  A zero value specifies that Solver should use a different seed for the random number generator each time it runs, which may yield different solutions each time it is run on a model that has not changed."
    With oSolverSheet.Names
        .Add Name:="solver_rsd", RefersToLocal:=val, Visible:=nameVisible
    End With
End Property
Public Property Get RandomSeed() As Long
    On Error Resume Next
    PopulationSize = NameToLong("solver_rsd", oSolverSheet)
End Property

'%Description("True to have Solver use multistart method for global optimization with the GRG Nonlinear Solving method, when SolverSolve is called.  False to have Solver run the GRG Solving method only once, without multistart, when SolverSolve is called.")
Public Property Let MultiStart(ByVal val As Boolean)
Attribute MultiStart.VB_Description = "True to have Solver use multistart method for global optimization with the GRG Nonlinear Solving method, when SolverSolve is called.  False to have Solver run the GRG Solving method only once, without multistart, when SolverSolve is called."
    With oSolverSheet.Names
        If val Then
           .Add Name:="solver_msl", RefersToLocal:=1, Visible:=nameVisible
        Else
           .Add Name:="solver_msl", RefersToLocal:=2, Visible:=nameVisible
        End If
    End With
End Property
Public Property Get MultiStart() As Boolean
    On Error Resume Next
    PopulationSize = (NameToLong("solver_msl", oSolverSheet) = 1)
End Property

'%Description("True to cause the Evolutionary Solving method and the multistart method to return immediately from a call to SolveIt with a value of 18 if any of the variables do not have both lower and upper bounds defined.  False to have these methods attempt to solve the problem without bounds on all of the variables.")
Public Property Let RequireBounds(ByVal val As Boolean)
Attribute RequireBounds.VB_Description = "True to cause the Evolutionary Solving method and the multistart method to return immediately from a call to SolveIt with a value of 18 if any of the variables do not have both lower and upper bounds defined.  False to have these methods attempt to solve the problem without bounds on all of the variables."
    With oSolverSheet.Names
        If val Then
           .Add Name:="solver_rbv", RefersToLocal:=1, Visible:=nameVisible
        Else
           .Add Name:="solver_rbv", RefersToLocal:=2, Visible:=nameVisible
        End If
    End With
End Property
Public Property Get RequireBounds() As Boolean
    On Error Resume Next
    RequireBounds = (NameToLong("solver_rbv", oSolverSheet) = 1)
End Property

'%Description("A number between 0 (zero) and 1 that specifies the rate at which the Evolutionary Solving method will make mutations to existing population members.  A higher Mutation rate tends to increase the diversity of the population, and may yield better solutions.")
Public Property Let MutationRate(ByVal val As Double)
Attribute MutationRate.VB_Description = "A number between 0 (zero) and 1 that specifies the rate at which the Evolutionary Solving method will make mutations to existing population members.  A higher Mutation rate tends to increase the diversity of the population, and may yield better solutions."
    With oSolverSheet.Names
        .Add Name:="solver_mrt", RefersToLocal:=val, Visible:=nameVisible
    End With
End Property
Public Property Get MutationRate() As Double
    On Error Resume Next
    MutationRate = NameToDbl("solver_mrt", oSolverSheet)
End Property

'%Description("The maximum number of subproblems Solver will explore in problems with integer constraints, and problems solved via the Evolutionary Solving method. The value must be a positive integer.")
Public Property Let MaxSubproblems(ByVal val As Long)
Attribute MaxSubproblems.VB_Description = "The maximum number of subproblems Solver will explore in problems with integer constraints, and problems solved via the Evolutionary Solving method. The value must be a positive integer."
    With oSolverSheet.Names
        If val <> 0 Then
            .Add Name:="solver_nod", RefersToLocal:=val, Visible:=nameVisible
        Else
            .Add Name:="solver_nod", RefersToLocal:=MAXLONG, Visible:=nameVisible
        End If
    End With
End Property
Public Property Get MaxSubproblems() As Long
    On Error Resume Next
    MaxSubproblems = NameToLong("solver_nod", oSolverSheet)
End Property

'%Description("The maximum number of feasible (or integer feasible) solutions Solver will consider in problems with integer constraints, and problems solved via the Evolutionary Solving method. The value must be a positive integer.")
Public Property Let MaxIntegerSols(ByVal val As Long)
Attribute MaxIntegerSols.VB_Description = "The maximum number of feasible (or integer feasible) solutions Solver will consider in problems with integer constraints, and problems solved via the Evolutionary Solving method. The value must be a positive integer."
    With oSolverSheet.Names
        If val <> 0 Then
           .Add Name:="solver_mip", RefersToLocal:=val, Visible:=nameVisible
        Else
           .Add Name:="solver_mip", RefersToLocal:=MAXLONG, Visible:=nameVisible
        End If
    End With
End Property
Public Property Get MaxIntegerSols() As Long
    On Error Resume Next
    MaxIntegerSols = NameToLong("solver_mip", oSolverSheet)
End Property

'%Description("True to have Solver ignore any integer constraints and solve the relaxation of the problem.  False to have Solver use the integer constraints in solving the problem.")
Public Property Let SolveWithout(ByVal val As Boolean)
Attribute SolveWithout.VB_Description = "True to have Solver ignore any integer constraints and solve the relaxation of the problem.  False to have Solver use the integer constraints in solving the problem."
    With oSolverSheet.Names
        If val Then
           .Add Name:="solver_rlx", RefersToLocal:=1, Visible:=nameVisible
        Else
           .Add Name:="solver_rlx", RefersToLocal:=2, Visible:=nameVisible
        End If
    End With
End Property
Public Property Get SolveWithout() As Boolean
    On Error Resume Next
    SolveWithout = (NameToLong("solver_rlx", oSolverSheet) = 1)
End Property

'%Description("When the Evolutionary Solving method is used, the maximum amount of time (in seconds) Solver will continue solving without finding significantly improved solutions to add to the population. The value must be a positive integer.")
Public Property Let MaxTimeNoImp(ByVal val As Long)
Attribute MaxTimeNoImp.VB_Description = "When the Evolutionary Solving method is used, the maximum amount of time (in seconds) Solver will continue solving without finding significantly improved solutions to add to the population. The value must be a positive integer."
    With oSolverSheet.Names
        .Add Name:="solver_mni", RefersToLocal:=val, Visible:=nameVisible
    End With
End Property
Public Property Get MaxTimeNoImp() As Long
    On Error Resume Next
    MaxTimeNoImp = NameToLong("solver_mni", oSolverSheet)
End Property

' ==========================================================================
' Private members
' ==========================================================================

Friend Sub SetSolverSheet(ws As Worksheet)
    Set oSolverSheet = ws
End Sub

Private Sub Class_Terminate()
    Set oSolverSheet = Nothing
End Sub
