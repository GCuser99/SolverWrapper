Attribute VB_Name = "modGlobals"
'These are declares needed to run twinBASIC without registration
Option Explicit

Public Declare PtrSafe Function New_SolvProblem Lib "[Path to DLL]\SolverWrapper_win64.dll" () As Object
'Public Declare PtrSafe Function New_SolvProblem Lib "[Path to DLL]\SolverWrapper_win32.dll" () As Object

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

