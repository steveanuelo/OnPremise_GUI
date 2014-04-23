VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_CoopStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rngCreatorName As String
Public rngCoopPeriod As String
Public rngStart As String

Private Sub Class_Initialize()
    rngCreatorName = "C2"
    rngCoopPeriod = "C4"
    rngStart = "B8"
End Sub
