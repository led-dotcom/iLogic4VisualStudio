Imports System
Imports System.Collections.Generic
Imports System.Math
Imports System.Windows.Forms
Imports Autodesk.iLogic.Interfaces
Imports Autodesk.iLogic.Runtime
Imports Autodesk.iLogic.Types
Imports Inventor

Namespace iLogic4VisualStudio
    Public Class Table_Rule
        Inherits RuleBase

        Public Overrides _
        Sub Main()
            'Assembly level params
            Parameter("LeftConectionDistance") = 0
            Parameter("RightConectionDistance") = 0

            '''Part level
            '''Leg params
            Parameter("leg:1", "LegDiameter") = 2
            Parameter("leg:1", "LegHeight") = 20

            'Renew the table immediately
            iLogicVb.UpdateWhenDone = True
        End Sub
    End Class
End Namespace
