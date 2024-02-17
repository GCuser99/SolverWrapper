VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolvProblem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "The main class to define the Solver problem, including Objective, Decision Variables, and Constraints."
'%ModuleDescription "The main class to define the Solver problem, including Objective, Decision Variables, and Constraints."
'@Exposed
'@folder("SolverWrapper.Source")
' ==========================================================================
' SolverWrapper v0.1
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

Private oObjective As SolvObjective
Private oConstraints As SolvConstraints
Private oDecisionVars As SolvDecisionVars
Private oSolver As SolvSolver
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents oSolvDLL As SolvDLL
Attribute oSolvDLL.VB_VarHelpID = -1
Private oSolverSheet As Worksheet

'@VariableDescription("Fires on each solution trial iteration to allow user to perform intermediate processing.")
Public Event ShowTrial(ByVal reason As Long, ByVal trialNum As Long, stopSolver As Boolean)

'@VariableDescription("Fires on each solution trial iteration to allow user to perform intermediate processing.")
Public Event BeforeSolve(stopSolver As Boolean)

'@VariableDescription("Fires on each solution trial iteration to allow user to perform intermediate processing.")
Public Event AfterSolve(ByVal returnMsgCode As Long, ByVal trialNum As Long)

Public Enum SlvGoalType
    slvMaximize = 1
    slvMinimize = 2
    slvTargetValue = 3
    [_First] = 1
    [_Last] = 3
End Enum

Public Enum SlvShowTrial
    slvContinue = 0
    slvStop = 1
    [_First] = 0
    [_Last] = 1
End Enum

Public Enum SlvEstimates
    slvTangent = 1
    slvQuadratic = 2
    [_First] = 1
    [_Last] = 2
End Enum

Public Enum SlvDerivatives
    slvForward = 1
    slvCentral = 2
    [_First] = 1
    [_Last] = 2
End Enum

Public Enum SlvSearchOption
    slvNewton = 1
    slvConjugate = 2
    [_First] = 1
    [_Last] = 2
End Enum

Public Enum SlvRelation
    slvLessThanEqual = 1
    slvEqual = 2
    slvGreaterThanEqual = 3
    slvInt = 4
    slvBin = 5
    slvDif = 6
    [_First] = 1
    [_Last] = 6
End Enum

Public Enum SlvSolveMethod
    slvGRG_Nonlinear = 1
    slvSimplex_LP = 2
    slvEvolutionary = 3
    [_First] = 1
    [_Last] = 3
End Enum

Public Enum SlvCallbackReason
    slvShowIterations = 1
    slvMaxTimeLimit = 2
    slvMaxIterationsLimit = 3
    slvMaxSubproblemsLimit = 4
    slvMaxSolutionsLimit = 5
    [_First] = 1
    [_Last] = 5
End Enum

Private Enum SolverMode
    SolveMode = 0
    CloseMode = 1
    CancelRestoreMode = 2
    [_First] = 0
    [_Last] = 2
End Enum

' ==========================================================================
' Public members
' ==========================================================================

'%Description("Exposes the Constraints object.")
Public Property Get Constraints() As SolvConstraints
Attribute Constraints.VB_Description = "Exposes the Constraints object."
    If Not IsInitialized Then Err.Raise vbObjectError + 4001, , "Must call Initialize method first"
    Set Constraints = oConstraints
End Property

'%Description("Exposes the DecisionVars object.")
Public Property Get DecisionVars() As SolvDecisionVars
Attribute DecisionVars.VB_Description = "Exposes the DecisionVars object."
    If Not IsInitialized Then Err.Raise vbObjectError + 4001, , "Must call Initialize method first"
    Set DecisionVars = oDecisionVars
End Property

'%Description("Exposes the Objective object.")
Public Property Get Objective() As SolvObjective
Attribute Objective.VB_Description = "Exposes the Objective object."
    If Not IsInitialized Then Err.Raise vbObjectError + 4001, , "Must call Initialize method first"
    Set Objective = oObjective
End Property

'%Description("Exposes the Solver object.")
Public Function Solver() As SolvSolver
Attribute Solver.VB_Description = "Exposes the Solver object."
    If Not IsInitialized Then Err.Raise vbObjectError + 4001, , "Must call Initialize method first"
    Set Solver = oSolver
End Function

'%Description("Initializes the problem - this must be called before all other methods.")
Public Sub Initialize(ws As Worksheet, Optional ByVal resetSolver As Boolean = True)
Attribute Initialize.VB_Description = "Initializes the problem - this must be called before all other methods."
    Set oSolverSheet = ws
    
    #If TWINBASIC Then
    Set Application = ws.Application
    #End If
    
    ws.Activate
    
    If ws.ProtectContents Then
        Err.Raise vbObjectError + 4001, , "Problem cannot be initilialized because the active Worksheet is protected."
    End If
    
    InitializeSolverNames
    
    If resetSolver Then
        DeleteSolverNames ws:=ws, deleteModel:=True, deleteSolverOptions:=True, deleteStatusBarNames:=True
        InitSolver ws:=ws, initModel:=True, initSolverOptions:=True, initStatusBarNames:=True
    End If
    
    oObjective.SetSolverSheet ws
    oDecisionVars.SetSolverSheet ws
    oSolver.SetSolverSheet ws
    oConstraints.SetSolverSheet ws
    oSolver.Options.SetSolverSheet ws
End Sub

'%Description("Solves the currently defined problem. Returns the Solver result code.")
Public Function SolveIt(Optional ByVal solverMethodToUse As SlvSolveMethod = -1) As Long
Attribute SolveIt.VB_Description = "Solves the currently defined problem. Returns the Solver result code."
    If Not IsInitialized Then Err.Raise vbObjectError + 4001, , "Must call Initialize method first"
    
    If solverMethodToUse > 0 Then Me.Solver.Method = solverMethodToUse
    
    oSolvDLL.saveSolutions = Me.Solver.SaveAllTrialSolutions
    Set oSolvDLL.oSolverSheet = oSolverSheet
    oSolvDLL.enableInternalEvents = oSolver.EnableEvents
    oSolvDLL.enableUserDefinedCallback = (Me.Solver.UserCallbackMacroName <> vbNullString)
    
    SolveIt = oSolvDLL.RunSolver
End Function

'%Description("Returns the string message associated with the Solver code returned by SolveIt method.")
Public Function SolverCodeToString(ByVal code As Long) As String
Attribute SolverCodeToString.VB_Description = "Returns the string message associated with the Solver code returned by SolveIt method."
    SolverCodeToString = ReturnSolverMsg(code)
End Function

'%Description("Cleans the problem worksheet of all solver associated info.")
Public Sub CleanUp()
Attribute CleanUp.VB_Description = "Cleans the problem worksheet of all solver associated info."
    If Not IsInitialized Then Err.Raise vbObjectError + 4001, , "Must call Initialize method first"
    DeleteSolverNames ws:=oSolverSheet, deleteModel:=True, deleteSolverOptions:=True, deleteStatusBarNames:=True
End Sub

'%Description("Resets the problem objective and decision variables without affecting the solver options.")
Public Sub Reset()
Attribute Reset.VB_Description = "Resets the problem objective and decision variables without affecting the solver options."
    If Not IsInitialized Then Err.Raise vbObjectError + 4001, , "Must call Initialize method first"
    DeleteSolverNames ws:=oSolverSheet, deleteModel:=True, deleteSolverOptions:=False, deleteStatusBarNames:=True
End Sub

'%Description("Returns a Variant array containing all intermediate (trial) and final solutions, as well as a flag indicating whether each solution passed all problem contraints.")
Public Function GetSolutions(Optional ByVal sortBestToWorst As Boolean = True, Optional ByVal removeDuplicates As Boolean = True, Optional ByVal keepOnlyValid As Boolean = True) As Variant()
Attribute GetSolutions.VB_Description = "Returns a Variant array containing all intermediate (trial) and final solutions, as well as a flag indicating whether each solution passed all problem contraints."
    Dim a() As Variant, i As Long, j As Long, k As Long
    If Not oSolver.SaveAllTrialSolutions Then Err.Raise vbObjectError + 4001, , "Error in GetSolutions method  - Must first set SaveAllTrialSolutions of Solver Class to True"
    a = oSolvDLL.GetSolutions
    If keepOnlyValid Then RemoveArrayInvalids a
    If removeDuplicates Then RemoveArrayDups a
    a = TransposeArray(a)
    If sortBestToWorst Then a = FlipArray(a)
    GetSolutions = a
End Function

'%Description("Saves all intermediate and final solutions to the given Worksheet Range.")
Public Sub SaveSolutionsToRange(ByRef saveRange As Range, Optional ByVal sortBestToWorst As Boolean = True, Optional ByVal removeDuplicates As Boolean = True, Optional ByVal keepOnlyValid As Boolean = True)
Attribute SaveSolutionsToRange.VB_Description = "Saves all intermediate and final solutions to the given Worksheet Range."
    If Not oSolver.SaveAllTrialSolutions Then Err.Raise vbObjectError + 4001, , "Error in SaveSolutionsToRange method  - Must first set SaveAllTrialSolutions of Solver Class to True"
    VariantToRange GetSolutions(sortBestToWorst, removeDuplicates, keepOnlyValid), saveRange
End Sub

'%Description("Returns a reference to the current problem Worksheet")
Public Function SolverSheet() As Worksheet
Attribute SolverSheet.VB_Description = "Returns a reference to the current problem Worksheet"
    If Not IsInitialized Then Err.Raise vbObjectError + 4001, , "Must call Initialize method first"
    Set SolverSheet = oSolverSheet
End Function

' ==========================================================================
' Private members
' ==========================================================================

'The following three Event procedures respond to trial events in SolvDLL, which are not publically exposed.
'These each in turn raise public events that User can respond to.
Private Sub oSolvDLL_ShowTrialInternal(ByVal reason As Long, ByVal trialNum As Long, stopSolver As Boolean)
    RaiseEvent ShowTrial(reason, trialNum, stopSolver)
End Sub
Private Sub oSolvDLL_BeforeSolveInternal(stopSolver As Boolean)
    'chained public trial event responding to private trial event in SolvDLL, which is not publically exposed
    RaiseEvent BeforeSolve(stopSolver)
End Sub
Private Sub oSolvDLL_AfterSolveInternal(ByVal returnMsgCode As Long, ByVal trialNum As Long)
    'chained public trial event responding to private trial event in SolvDLL, which is not publically exposed
    RaiseEvent AfterSolve(returnMsgCode, trialNum)
End Sub
Private Sub oSolvDLL_ShowTrialCallbackInternal(ByVal reason As Long, ByVal trialNum As Long, stopSolver As Boolean)
    On Error Resume Next
    'user-defined callback must have the following call signature:
    'Function ShowTrial(ByVal reason As Long, ByVal trialNum As Long, oProblem As SolvProblem) As Boolean
    'where return value is a boolean indicating whether to stop or not
    stopSolver = Application.Run(Me.Solver.UserCallbackMacroName, reason, trialNum, Me)
    If Err.number <> 0 Then
        MsgBox Err.Description
        stopSolver = True
    End If
End Sub

Private Function IsInitialized() As Boolean
    If oSolverSheet Is Nothing Then IsInitialized = False Else IsInitialized = True
End Function

Public Sub InitializeSolverNames()
    Set solverNames = New Collection
    With solverNames
        'check to see if solver_old is created with multi-start?
        'entries 1-10
        .Add "solver_adj" 'solver decision var input ranges
        .Add "solver_eng" 'engine (1,2,3) (GRG, Simplex, Evolve)
        .Add "solver_num" 'number of constraints
        .Add "solver_opt" 'model obj function to optimize
        .Add "solver_lhs" 'constraint range (these are numbered eg. solver_lhs1)
        .Add "solver_rel" 'constraint rel (these are numbered eg. solver_rel1)
        .Add "solver_rhs" 'constraint value (these are numbered eg. solver_rhs1)
        .Add "solver_typ" 'Optimization Type (1-max, 2-min, 3-val of)
        .Add "solver_val" 'value for typ=3 (after reset=0)
        .Add "solver_ver" 'Solver version = 3

         'worksheet-level names for solver options
         'entries 11-30
        .Add "solver_cvg" 'convergence 0.0001
        .Add "solver_drv" 'derivatives (1 - forward, 2 - central)
        .Add "solver_est" 'Estimates (tangent-1, 2-quadratic)
        .Add "solver_itr" 'solving limit iterations MAXLONG
        '.Add "solver_lin" 'assume linear problem? (1-yes, 2-no)
        .Add "solver_mip" 'solving limits max subproblems MAXLONG
        .Add "solver_mni" 'max time wo improve in secs (30)
        .Add "solver_mrt" 'mutation rate (.075)
        .Add "solver_msl" 'multistart (1-no, 2-yes)
        .Add "solver_neg" 'make unconstrained vars non neg (1 -yes, 2 - no)
        .Add "solver_nod" 'solving limits max feasible solutions MAXLONG
        .Add "solver_nwt" '1
        .Add "solver_pre" 'constraint prec 0.000001
        .Add "solver_rbv" 'require bounds on vars (1-yes, 2-no)
        .Add "solver_rlx" 'ignore integer constraints (1-yes, 2-no)
        .Add "solver_rsd" 'random seed (0)
        .Add "solver_scl" 'auto scale (0,1)
        .Add "solver_sho" 'show inter results (1-yes, 2-no)
        .Add "solver_ssz" 'multistart sub problem size 100
        .Add "solver_tim" 'solving limit time MAXLONG
        .Add "solver_tol" 'integer optimality .01

        'workbook-level names for status bar messaging
        'entries 31-35
        .Add "msg_setup"
        .Add "msg_trial"
        .Add "msg_branch"
        .Add "msg_incumb"
        .Add "msg_setcell"
    End With
End Sub

Private Sub Class_Initialize()
    Set oConstraints = New SolvConstraints
    Set oDecisionVars = New SolvDecisionVars
    Set oObjective = New SolvObjective
    Set oSolver = New SolvSolver
    Set oSolvDLL = New SolvDLL
    Set oSolverSheet = Nothing
End Sub

Private Sub Class_Terminate()
    Set oConstraints = Nothing
    Set oDecisionVars = Nothing
    Set oObjective = Nothing
    Set oSolver = Nothing
    Set oSolvDLL = Nothing
    Set oSolverSheet = Nothing
    #If TWINBASIC Then
    Set Application = Nothing
    #End If
End Sub
