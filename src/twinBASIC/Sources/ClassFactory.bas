Attribute VB_Name = "ClassFactory"
Attribute VB_Description = "This class is used for object instantiation when referencing SolverWrapper externally from another VBA project"
'%ModuleDescription "This class is used for object instantiation when referencing SolverWrapper externally from another VBA project"
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

' From the referencing vba project use the following syntax:
'
' Dim problem as SolverWrapper.SolvProblem
' Set problem = SolverWrapper.New_SolvProblem
'
Option Explicit

'%Description("Instantiates a SolvProblem object")
Public Function New_SolvProblem() As SolvProblem
Attribute New_SolvProblem.VB_Description = "Instantiates a SolvProblem object"
    Set New_SolvProblem = New SolvProblem
End Function
