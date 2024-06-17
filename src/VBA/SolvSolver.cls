VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolvSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class to manage Solver Engine and Options."
'@ModuleDescription "A class to manage Solver Engine and Options."
'@Exposed
'@folder("SolverWrapper.Source")
' ==========================================================================
' SolverWrapper v1.0
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

'About the Solver Methods from https://www.solver.com/excel-solver-algorithms-and-methods-used:
'The GRG Nonlinear Solving Method for nonlinear optimization uses the Generalized Reduced Gradient (GRG2) code,
'which was developed by Leon Lasdon, University of Texas at Austin, and Alan Waren, Cleveland State University,
'and enhanced by Frontline Systems, Inc.

'The Simplex LP Solving Method for linear programming uses the Simplex and dual Simplex method with bounds on the variables,
'and problems with integer constraints use the branch and bound method, as implemented by John Watson and Daniel Fylstra,
'Frontline Systems, Inc.

'The Evolutionary Solving Method for non-smooth optimization uses a variety of genetic algorithm and local search methods,
'implemented by several individuals at Frontline Systems, Inc.

Private oSolvOptions As SolvOptions
Private saveSolutions As Boolean
Private oSolverSheet As Worksheet
Private enableSolverEvents As Boolean
Private userDefinedCallback As String

' ==========================================================================
' Public members
' ==========================================================================

'@Description("Exposes the Solver Options object.")
Public Function Options() As SolvOptions
Attribute Options.VB_Description = "Exposes the Solver Options object."
    Set Options = oSolvOptions
    oSolvOptions.SetSolverSheet oSolverSheet
End Function

'@Description("Resets all Option values to defaults.")
Public Sub Reset()
Attribute Reset.VB_Description = "Resets all Option values to defaults."
    DeleteSolverNames oSolverSheet, False, True
    InitSolver oSolverSheet, False, True
End Sub

'@Description("Sets/Gets the solver method (engine) to use to solve the problem.")
Public Property Let Method(val As SlvSolveMethod)
Attribute Method.VB_Description = "Sets/Gets the solver method (engine) to use to solve the problem."
    oSolverSheet.Names.Add Name:="solver_eng", RefersToLocal:=val, Visible:=nameVisible
End Property
Public Property Get Method() As SlvSolveMethod
    On Error Resume Next
    Method = NameToLong("solver_eng")
End Property

'@Description("Sets whether all intermediate solutions attempted should be saved - see GetAllSolutions to return solution array.")
Public Property Let SaveAllTrialSolutions(val As Boolean)
Attribute SaveAllTrialSolutions.VB_Description = "Sets whether all intermediate solutions attempted should be saved - see GetAllSolutions to return solution array."
    saveSolutions = val
End Property
Public Property Get SaveAllTrialSolutions() As Boolean
    SaveAllTrialSolutions = saveSolutions
End Property

'@Description("Enables the ShowTrial, BeforeSolve, AfterSolve events of the SolvProblem class (default = False).")
Public Property Let EnableEvents(val As Boolean)
Attribute EnableEvents.VB_Description = "Enables the ShowTrial, BeforeSolve, AfterSolve events of the SolvProblem class (default = False)."
    enableSolverEvents = val
End Property
Public Property Get EnableEvents() As Boolean
    EnableEvents = enableSolverEvents
End Property

'@Description("Sets the name of the user-defined callback function that will be called on each trial solution. If this is used in conjunction with th ShowTrial event, then the return value from the callback function takes precedence over the return value from the event. ")
Public Property Let UserCallbackMacroName(ByVal val As String)
Attribute UserCallbackMacroName.VB_Description = "Sets the name of the user-defined callback function that will be called on each trial solution. If this is used in conjunction with th ShowTrial event, then the return value from the callback function takes precedence over the return value from the event. "
    If InStr(val, "!") Then
        userDefinedCallback = val
    Else
        userDefinedCallback = "'" & oSolverSheet.Parent.Name & "'!" & val
    End If
End Property
Public Property Get UserCallbackMacroName() As String
    UserCallbackMacroName = userDefinedCallback 'End Property
End Property

'@Description("Returns the string message associated with the Solver code returned by SolveIt method.")
Public Function MsgCodeToString(ByVal code As Long) As String
Attribute MsgCodeToString.VB_Description = "Returns the string message associated with the Solver code returned by SolveIt method."
    MsgCodeToString = ReturnSolverMsg(code)
End Function

' ==========================================================================
' Private members
' ==========================================================================

Friend Sub SetSolverSheet(ws As Worksheet)
    Set oSolverSheet = ws
End Sub

Private Sub Class_Initialize()
    Set oSolvOptions = New SolvOptions
    userDefinedCallback = vbNullString
End Sub

Private Sub Class_Terminate()
    Set oSolvOptions = Nothing
    Set oSolverSheet = Nothing
End Sub
