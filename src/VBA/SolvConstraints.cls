VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolvConstraints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class to set/manage Solver constraints."
'@ModuleDescription "A class to set/manage Solver constraints."
'@Exposed
'@folder("SolverWrapper.Source")
' ==========================================================================
' SolverWrapper v1.1
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

'CellRef Required Variant. A reference to a cell or a range of cells that forms the left side of a constraint.
'Relation Required Integer. The arithmetic relationship between the left and right sides of the constraint.
'If you choose 4, 5, or 6, CellRef must refer to decision variable cells, and FormulaText should not be specified.

' ==========================================================================
' Public members
' ==========================================================================

'@Description("Adds a constraint to the current problem.")
Public Sub Add(ByVal cellRef As Variant, ByVal relation As SlvRelation, Optional ByVal formulaText As Variant)
Attribute Add.VB_Description = "Adds a constraint to the current problem."
    Dim ws As Worksheet
    Dim solver_num As Long
    
    Set ws = oSolverSheet
    
    'qc the inputs
    ProcessConstraintParams cellRef, relation, formulaText
    
    'find number of existing constraints plus 1
    If NameExists("solver_num", ws) Then solver_num = NameToLong("solver_num", ws) + 1
    
    If relation > 3 Then
        ws.Names.Add Name:="solver_lhs" & CStr(solver_num), RefersTo:="=" & ws.Range(cellRef).Address, Visible:=nameVisible
        ws.Names.Add Name:="solver_rel" & CStr(solver_num), RefersTo:=relation, Visible:=nameVisible
        ws.Names.Add Name:="solver_num", RefersTo:=solver_num, Visible:=nameVisible
        Select Case relation
        Case SlvRelation.slvInt
            ws.Names.Add Name:="solver_rhs" & CStr(solver_num), RefersTo:="=""integer""", Visible:=nameVisible
        Case SlvRelation.slvBin
            ws.Names.Add Name:="solver_rhs" & CStr(solver_num), RefersTo:="=""binary""", Visible:=nameVisible
        Case SlvRelation.slvAllDif
            ws.Names.Add Name:="solver_rhs" & CStr(solver_num), RefersTo:="=""AllDifferent""", Visible:=nameVisible
        End Select
    Else 'Relation <=3
        ws.Names.Add Name:="solver_lhs" & CStr(solver_num), RefersTo:="=" & ws.Range(cellRef).Address, Visible:=nameVisible
        ws.Names.Add Name:="solver_rel" & CStr(solver_num), RefersTo:=relation, Visible:=nameVisible
        ws.Names.Add Name:="solver_rhs" & CStr(solver_num), RefersTo:=formulaText, Visible:=nameVisible
        ws.Names.Add Name:="solver_num", RefersTo:=solver_num, Visible:=nameVisible
    End If
End Sub

'@Description("Adds a bounded constraint to the current problem. Equivalent to adding both cellRef>=lowBound and cellRef<=highBound")
Public Sub AddBounded(ByVal cellRef As Variant, ByVal lowBound As Variant, ByVal highBound As Variant)
Attribute AddBounded.VB_Description = "Adds a bounded constraint to the current problem. Equivalent to adding both cellRef>=lowBound and cellRef<=highBound"
    Add cellRef, slvGreaterThanEqual, lowBound
    Add cellRef, slvLessThanEqual, highBound
End Sub

'@Description("Gets the total number of constraints set for the current problem.")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the total number of constraints set for the current problem."
    If NameExists("solver_num", oSolverSheet) Then Count = NameToLong("solver_num", oSolverSheet)
End Property

'@Description("Deletes a constraint from the current problem.")
Public Sub Relax(ByVal cellRef As Variant, ByVal relation As SlvRelation, Optional ByVal formulaText As Variant)
Attribute Relax.VB_Description = "Deletes a constraint from the current problem."
    Dim ws As Worksheet
    Dim solver_num As Long
    Dim saveScreenUpdating As Long
    Dim savConstraints As New Collection
    Dim i As Long
    Dim cmatch As Long
    
    Set ws = oSolverSheet
    
    'qc the inputs
    ProcessConstraintParams cellRef, relation, formulaText
    
    'find number of existing constraints
    solver_num = NameToLong("solver_num")

    For i = 1 To solver_num
        savConstraints.Add Array(NameToString("solver_lhs" & i, ws), NameToLong("solver_rel" & i, ws), NameToString("solver_rhs" & i, ws)), CStr(i)
    Next i
    
    cmatch = 0
    For i = 1 To solver_num
        If savConstraints.Item(CStr(i))(0) = "'" & ws.Name & "'!" & cellRef Then
            If savConstraints.Item(CStr(i))(1) = relation Then
                If relation <= 3 Then
                    If IsNumeric(savConstraints.Item(CStr(i))(2)) Then
                        If savConstraints.Item(CStr(i))(2) = formulaText Then
                            cmatch = i
                            Exit For
                        End If
                    Else
                        If savConstraints.Item(CStr(i))(2) = "'" & ws.Name & "'!" & formulaText Then
                            cmatch = i
                            Exit For
                        End If
                    End If
                Else
                    cmatch = i
                    Exit For
                End If
            End If
        End If
    Next i
    If cmatch > 0 Then
        savConstraints.Remove CStr(cmatch)
        
        For i = 1 To solver_num
            ws.Names("solver_lhs" & i).Delete
            ws.Names("solver_rel" & i).Delete
            ws.Names("solver_rhs" & i).Delete
        Next i
        
        Dim key As Variant
        solver_num = 0
        For Each key In savConstraints
            cellRef = key(0)
            relation = key(1)
            formulaText = key(2)
            solver_num = solver_num + 1
            If relation > 3 Then
                ws.Names.Add Name:="solver_lhs" & CStr(solver_num), RefersTo:="=" & ws.Range(cellRef).Address, Visible:=nameVisible
                ws.Names.Add Name:="solver_rel" & CStr(solver_num), RefersTo:=relation, Visible:=nameVisible

                Select Case relation
                Case SlvRelation.slvInt
                    ws.Names.Add Name:="solver_rhs" & CStr(solver_num), RefersTo:="=""integer""", Visible:=nameVisible
                Case SlvRelation.slvBin
                    ws.Names.Add Name:="solver_rhs" & CStr(solver_num), RefersTo:="=""binary""", Visible:=nameVisible
                Case SlvRelation.slvAllDif
                    ws.Names.Add Name:="solver_rhs" & CStr(solver_num), RefersTo:="=""AllDifferent""", Visible:=nameVisible
                End Select
            Else 'Relation <=3
                If Left$(formulaText, 1) <> "=" Then formulaText = "=" & formulaText
                ws.Names.Add Name:="solver_lhs" & CStr(solver_num), RefersTo:="=" & ws.Range(cellRef).Address, Visible:=nameVisible
                ws.Names.Add Name:="solver_rel" & CStr(solver_num), RefersTo:=relation, Visible:=nameVisible
                ws.Names.Add Name:="solver_rhs" & CStr(solver_num), RefersTo:=formulaText, Visible:=nameVisible
            End If
        Next key
        ws.Names.Add Name:="solver_num", RefersTo:=solver_num, Visible:=nameVisible
    End If
End Sub

'@Description("Deletes a bounded constraint from the current problem.")
Public Sub RelaxBounded(ByVal cellRef As Variant, ByVal lowBound As Variant, ByVal highBound As Variant)
Attribute RelaxBounded.VB_Description = "Deletes a bounded constraint from the current problem."
    Relax cellRef, slvGreaterThanEqual, lowBound
    Relax cellRef, slvLessThanEqual, highBound
End Sub

'@Description("Deletes all constraints from the current problem.")
Public Sub RelaxAll()
Attribute RelaxAll.VB_Description = "Deletes all constraints from the current problem."
    Dim i As Long
    Dim ws As Worksheet
    Set ws = oSolverSheet
    For i = 1 To Me.Count
        ws.Names("solver_lhs" & i).Delete
        ws.Names("solver_rel" & i).Delete
        ws.Names("solver_rhs" & i).Delete
    Next i
    ws.Names("solver_num").value = 0
End Sub

'@Description("Changes a constraint of the current problem. If cellRef AND relation do not match an existing constraint, then use Relax and Add instead.")
Public Sub Change(ByVal cellRef As Variant, ByVal relation As SlvRelation, ByVal newFormulaText As Variant)
Attribute Change.VB_Description = "Changes a constraint of the current problem. If cellRef AND relation do not match an existing constraint, then use Relax and Add instead."
    Dim ws As Worksheet
    Dim solver_num As Long
    Dim matchCount As Long
    Dim matchNum As Long
    Dim i As Long
    
    Set ws = oSolverSheet
    
    'check that relation is in valid range
    If relation > SlvRelation.slvGreaterThanEqual Then
        Err.Raise vbObjectError + 4001, , "Please specify a less-than-or-equal, equal, or greater-than-or-equal relation."
    End If
    
    'qc the inputs
    ProcessConstraintParams cellRef, relation, newFormulaText

    'find number of existing constraints
    solver_num = NameToLong("solver_num", ws)
    
    'loop through all contraints tracking the matches
    matchCount = 0
    For i = 1 To solver_num
        If IsRangeSameRange(Application.Range(cellRef), ws.Names("solver_lhs" & i).RefersToRange) Then
            If relation = NameToLong("solver_rel" & i, ws) Then
                matchCount = matchCount + 1
                matchNum = i
            End If
        End If
    Next i
    
    'only make the change if one match was found
    If matchCount = 1 Then
        If Left$(newFormulaText, 1) <> "=" Then newFormulaText = "=" & newFormulaText
        ws.Names.Add Name:="solver_rhs" & CStr(matchNum), RefersTo:=newFormulaText, Visible:=nameVisible
    End If
End Sub

'@Description("Changes a bounded constraint of the current problem.")
Public Sub ChangeBounded(ByVal cellRef As Variant, ByVal newLowBound As Variant, ByVal newHighBound As Variant)
Attribute ChangeBounded.VB_Description = "Changes a bounded constraint of the current problem."
    Change cellRef, slvGreaterThanEqual, newLowBound
    Change cellRef, slvLessThanEqual, newHighBound
End Sub

'@Description("Returns True if all model Constraints are satisfied.")
Public Function AreSatisfied() As Boolean
Attribute AreSatisfied.VB_Description = "Returns True if all model Constraints are satisfied."
    AreSatisfied = AreConstraintsSatisfied(oSolverSheet)
End Function

' ==========================================================================
' Private members
' ==========================================================================

Private Sub ProcessConstraintParams(ByRef cellRef As Variant, ByVal relation As Long, Optional ByRef formulaText As Variant)
    Dim ws As Worksheet
    
    Set ws = oSolverSheet
    
    'check that relation is in valid range
    If relation < SlvRelation.[_First] Or relation > SlvRelation.[_Last] Then
        Err.Raise vbObjectError + 4001, , "Please specify a valid Relation between 1 and 6."
    End If
    
    'process the LHS
    If Not IsError(cellRef) Then
        cellRef = CellRefToString(cellRef, ws)
    Else
        Err.Raise vbObjectError + 4001, , "Please specify a valid cell reference"
    End If
    
    If relation > 3 Then
        'LHS MUST refer to the decision variables
        'and thus dose not count against the 100-cell constraint limit
        If Not NameExists("solver_adj", ws) Then
            Err.Raise vbObjectError + 4001, , "Cannot set constraint relations slvInt, slvBin, or slvAllDif, until after the decision variables are defined."
        End If
        If Not IsRangeInRange(ws.Range(cellRef), ws.Names("solver_adj").RefersToRange) Then
            'the int, bin, and alldifferent constraints can only be applied to decision vars
            Err.Raise vbObjectError + 4001, , "The selected Relation only applies to contraints on decision variables."
        End If
    End If
    
    If relation <= 3 Then
        'process formula text (RHS)
        If IsMissing(formulaText) Or IsError(formulaText) Then Err.Raise vbObjectError + 4001, , "Please specify a valid formula text"
        
        formulaText = ProcessFormulaText(formulaText)
        
        If CellRefHasOneArea(formulaText, ws) Then
            'cell reference consists of a single contiguous cell block
            If ws.Range(formulaText).Count > 1 Then
                'multiple cells, so must match count of cellRef
                If ws.Range(formulaText).Count <> ws.Range(cellRef).Count Then
                    Err.Raise vbObjectError + 4001, , "If cell count on the right-hand side is greater than one, then must be equal to cell count on the left-hand side."
                End If
            End If
        Else
            'then to be valid, must be either numeric or a formula text that evaluates to numeric
            'Evaluate will take a formula like "=$o2" and convert it to the actual value of cell $o2
            If Not IsNumeric(Application.Evaluate(formulaText)) Then
                Err.Raise vbObjectError + 4001, , "The formula text does not evaluate to a valid formula."
            End If
        End If
    End If
End Sub

Private Function ProcessFormulaText(ByVal formulaText As Variant) As String
    Dim tmp As String
    Select Case TypeName(formulaText)
    Case "String"
        tmp = Application.ConvertFormula(formulaText, Application.ReferenceStyle, Application.ReferenceStyle, True)
        If Application.ReferenceStyle = xlR1C1 Then tmp = Application.ConvertFormula(tmp, xlR1C1, xlA1)
        If Application.International(xlDecimalSeparator) <> "." Then tmp = Replace(tmp, Application.International(xlDecimalSeparator), ".")
    Case "Range"
        tmp = formulaText.Address
    Case Else
        tmp = formulaText
    End Select
    If Left$(tmp, 1) <> "=" Then tmp = "=" & tmp
    ProcessFormulaText = tmp
End Function

Friend Sub SetSolverSheet(ws As Worksheet)
    Set oSolverSheet = ws
End Sub

Private Sub Class_Terminate()
    Set oSolverSheet = Nothing
End Sub
