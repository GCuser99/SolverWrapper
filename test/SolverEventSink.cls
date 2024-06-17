VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolverEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class is an example of one that would be written by user to trap and process events triggered by the SolvProblem class."
'@ModuleDescription "This class is an example of one that would be written by user to trap and process events triggered by the SolvProblem class."
'@folder("SolverWrapper.Examples")

'The EnableEvents property of the SolvSolver class must be set to True to activate SolverWrapper events

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents oProblem As SolvProblem
Attribute oProblem.VB_VarHelpID = -1

Private Sub oProblem_ShowTrial(ByVal reason As Long, ByVal trialNum As Long, stopSolver As Boolean)
    'this event is triggered for any one of the reasons below
    Dim i As Long
    
    Debug.Print "Trial number: " & trialNum
    Debug.Print "Objective: " & oProblem.Objective.CellRange.value
    
    For i = 1 To oProblem.DecisionVars.Count
        Debug.Print oProblem.DecisionVars.CellRange(i).Address, oProblem.DecisionVars.CellRange(i).value
    Next i
    
    Debug.Print "Constraints Satisfied? " & oProblem.Constraints.AreSatisfied
    
    'decide whether to stop solver based on the reason for the event trigger
    Select Case reason
        Case SlvCallbackReason.slvShowIterations 'new iteration has completed or user hit esc key
            stopSolver = False
        Case SlvCallbackReason.slvMaxTimeLimit
            stopSolver = True 'if set to True then solver is stopped!
        Case SlvCallbackReason.slvMaxIterationsLimit
            stopSolver = False
        Case SlvCallbackReason.slvMaxSubproblemsLimit
            stopSolver = False
        Case SlvCallbackReason.slvMaxSolutionsLimit
            stopSolver = False
    End Select
End Sub

Private Sub oProblem_BeforeSolve(stopSolver As Boolean)
    Debug.Print "Solver started on Worksheet: " & oProblem.SolverSheet.Name
End Sub

Private Sub oProblem_AfterSolve(ByVal returnMsgCode As Long, ByVal trialNum As Long)
    Debug.Print "Solver completed with iterations= " & trialNum
    Debug.Print "Solver return code= " & oProblem.Solver.MsgCodeToString(returnMsgCode)
End Sub

Public Property Set Problem(obj As SolvProblem)
    Set oProblem = obj
End Property

Private Sub Class_Terminate()
    Set oProblem = Nothing
    Debug.Print "terminating SolvEventSink"
End Sub
