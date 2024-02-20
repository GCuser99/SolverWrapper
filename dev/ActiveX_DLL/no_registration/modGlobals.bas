Attribute VB_Name = "modGlobals"
'These are declares needed to run twinBASIC DLL without registration
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
    slvAllDif = 6
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

Public Enum SlvMsgCode
    slvFoundSolution = 0
    slvConvergedOnSolution = 1
    slvCannotImproveSolution = 2
    slvMaxIterReached = 3
    slvObjectiveNotConvergent = 4
    slvCouldNotFindSolution = 5
    slvStoppedByUser = 6
    slvProblemNotLinear = 7
    slvProblemTooLarge = 8
    slvErrorInObjectiveOrConstraint = 9
    slvMaxTimeReached = 10
    slvNotEnoughMemory = 11
    slvNoDocumentation = 12
    slvErrorInModel = 13
    slvFoundIntegerSolution = 14
    slvMaxSolutionsReached = 15
    slvMaxSubProblemsReached = 16
    slvConvergedToGlobalSolution = 17
    slvAllVariablesMustBeBounded = 18
    slvBoundsConflictWithBinOrAllDif = 19
    slvBoundsAllowNoSolution = 20
    [_First] = 0
    [_Last] = 20
End Enum

