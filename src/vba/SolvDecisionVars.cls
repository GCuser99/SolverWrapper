VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolvDecisionVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class to define the Decision variable, i.e. the cells that are changed in order to satisfy the Objective."
'@ModuleDescription "A class to define the Decision variable, i.e. the cells that are changed in order to satisfy the Objective."
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
'@Description("Adds a cell(s) range that will be changed to achieve the desired result in the Objective cell - limited to 200 cells.")
Public Sub Add(ParamArray cellRefs() As Variant)
Attribute Add.VB_Description = "Adds a cell(s) range that will be changed to achieve the desired result in the Objective cell - limited to 200 cells."
Attribute Add.VB_UserMemId = 0
    Dim rCellRef As Range
    'initialize rCellRef - there can be more than one call to Add so check to see if its been called before...
    If NameExists("solver_adj", oSolverSheet) Then
        Set rCellRef = oSolverSheet.Names("solver_adj").RefersToRange()
        Set rCellRef = CellRefToRange(Array(rCellRef, Array(cellRefs)(0)), oSolverSheet)
    Else 'first time called
        Set rCellRef = CellRefToRange(Array(cellRefs)(0), oSolverSheet)
    End If
    
    If Me.Count + rCellRef.Count > 200 Then
        Err.Raise vbObjectError + 4001, , "Number of Decision Variable cells must be less than or equal to 200"
    End If
    
    ProcessDecisionVars rCellRef
End Sub

'@Description("Removes a cell or range of cells from the Decision variable set")
Public Sub Remove(ParamArray cellRefs() As Variant)
Attribute Remove.VB_Description = "Removes a cell or range of cells from the Decision variable set"
    Dim oldRange As Range
    Dim deleteRange As Range
    Dim rCellRef As Range
    Dim i As Long
    Set rCellRef = CellRefToRange(Array(cellRefs)(0), oSolverSheet)
    Set deleteRange = rCellRef
    Set oldRange = oSolverSheet.Names("solver_adj").RefersToRange
    oSolverSheet.Names("solver_adj").RefersTo = SubtractRangeFromRange(deleteRange, oldRange)
End Sub

'@Description("Removes all Decision variables.")
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Removes all Decision variables."
    DeleteName "solver_adj", oSolverSheet
End Sub

'@Description("Initializes Decision Variables to user-specified values - can be a constant to assign to all, or a list of values to assign to each variable.")
Public Sub Initialize(ParamArray cellVals() As Variant)
Attribute Initialize.VB_Description = "Initializes Decision Variables to user-specified values - can be a constant to assign to all, or a list of values to assign to each variable."
    Dim adjRange As Range, i As Long, cell As Range
    If NameExists("solver_adj", oSolverSheet) Then
        Set adjRange = oSolverSheet.Names("solver_adj").RefersToRange
        'assign cell values to decision var cells
        i = LBound(cellVals)
        For Each cell In adjRange
            cell.value = cellVals(i)
            i = Min(i + 1, UBound(cellVals))
        Next cell
    Else
        Err.Raise vbObjectError + 4001, , "Cannot Initialize the solution - no Decision variables are defined"
    End If
End Sub

'@Description("Gets the total number of Decision variables.")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the total number of Decision variables."
    If NameExists("solver_adj", oSolverSheet) Then
        Count = oSolverSheet.Names("solver_adj").RefersToRange().Count
    Else
        Count = 0
    End If
End Property

'@Description("Gets the Decision variable cell range.")
Public Function CellRange() As Range
Attribute CellRange.VB_Description = "Gets the Decision variable cell range."
    If NameExists("solver_adj", oSolverSheet) Then
        Set CellRange = oSolverSheet.Names("solver_adj").RefersToRange
    Else
        Set CellRange = Nothing
    End If
End Function

' ==========================================================================
' Private members
' ==========================================================================

Private Sub ProcessDecisionVars(ByVal byChange As Variant)
    'This procedure processes the decision variables and sets the corresponding sheet name
    Dim ws As Worksheet
    
    Set ws = oSolverSheet
    
    If Not IsError(byChange) Then
        byChange = CellRefToString(byChange, ws)
    Else
        Err.Raise vbObjectError + 4001, , "Please specify a valid cell reference"
    End If
    
    On Error Resume Next
    If Not ws.Range(byChange).Count > 0 Then
        Exit Sub
    End If
    On Error GoTo 0
    
    ws.Names.Add Name:="solver_adj", RefersTo:="=" & ws.Range(byChange).Address, Visible:=nameVisible
End Sub

Friend Sub SetSolverSheet(ws As Worksheet)
    Set oSolverSheet = ws
End Sub

Private Sub Class_Terminate()
    Set oSolverSheet = Nothing
End Sub
