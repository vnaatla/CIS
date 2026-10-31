Attribute VB_Name = "DataFormat"
Option Explicit

Type ExportCard
    SequenceNo As String * 5
    MemberName As String * 50
    MemberIC As String * 15
    MemberPol As String * 20
    MemberNo As String * 18
    MemberPlan As String * 4
    MemberExpiry As String * 11
End Type

Type VerifyError
    NoRecord As String * 5
    ErrorRecordLine As String * 12
    ErrorPolicyNo As String * 25
    ErrorName As String * 60
    ErrorMsg As String * 90
End Type

