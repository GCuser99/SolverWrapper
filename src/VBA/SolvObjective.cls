VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolvObjective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class to define the Solver Objective function."
'@ModuleDescription "A class to define the Solver Objective function."
'@Exposed
'@folder("SolverWrapper.Source")
' ==========================================================================
' SolverWrapper v0.9
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

'@DefaultMember
'@Description("Sets the Objective (Target) cell and parameters.")
Public Sub Define(ByVal setCell As Variant, ByVal goal As SlvGoalType, Optional ByVal valueOf As Double)
Attribute Define.VB_Description = "Sets the Objective (Target) cell and parameters."
Attribute Define.VB_UserMemId = 0
    'This procedure processes the objective and sets the corresponding the sheet name
    Dim ws As Worksheet
    
    Set ws = oSolverSheet
    
    If Not IsError(setCell) Then
        setCell = CellRefToString(setCell, ws)
    Else
        Err.Raise vbObjectError + 4001, , "Please specify a valid cell reference"
    End If
    
    If ws.Range(setCell).Count > 1 Then
        Err.Raise vbObjectError + 4001, , "Objective Cell must be a single cell on the active sheet"
    End If

    If goal < SlvGoalType.[_First] Or goal > SlvGoalType.[_Last] Then
        Err.Raise vbObjectError + 4001, , "Please specify an objective Goal value between 1 and 3."
    End If
        
    If IsMissing(valueOf) Or IsError(valueOf) Then
        valueOf = 0
    End If
        
    If goal = 3 And Not IsNumeric(valueOf) Then
        Err.Raise vbObjectError + 4001, , "Please specify a vaid ValueOf value."
    End If

    With ws.Names
        .Add Name:="solver_opt", RefersTo:="=" & ws.Range(setCell).Address, Visible:=nameVisible
        .Add Name:="solver_typ", RefersToLocal:=goal, Visible:=nameVisible
        .Add Name:="solver_val", RefersToLocal:=valueOf, Visible:=nameVisible
    End With
End Sub

'@Description("Gets the Objective variable cell value.")
Public Property Get CellRange() As Range
Attribute CellRange.VB_Description = "Gets the Objective variable cell value."
    Dim ws As Worksheet
    Set ws = oSolverSheet
    If NameExists("solver_opt", ws) Then
        Set CellRange = NameToRange(ws.Names("solver_opt"))
    Else
        Err.Raise vbObjectError + 4001, , "Must define problem objective before invoking CellRange property."
    End If
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
