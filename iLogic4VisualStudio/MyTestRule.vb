Imports System
Imports System.Collections.Generic
Imports System.Math
Imports System.Windows.Forms
Imports Autodesk.iLogic.Interfaces
Imports Autodesk.iLogic.Runtime
Imports Autodesk.iLogic.Types
Imports Inventor

Namespace iLogic4VisualStudio
    Public Class MyTestRule
        Inherits RuleBase

        Public Overrides _
        Sub Main()
            'Path to the Excel file
            Dim excelPath As String = "Query.xlsx"
            Dim sheetName As String = "Sheet1"

            'Assembly level params
            Dim choosedProject As String

            'Part level params
            Dim rack_width As Double
            Dim leg_height As Double

            choosedProject = Parameter("Project")
            GoExcel.FindRow(excelPath, sheetName, "Project", "=", choosedProject)
            rack_width = GoExcel.CurrentRowValue("Rack Width")
            leg_height = GoExcel.CurrentRowValue("Leg height")
            Parameter("Undershelf - Copy:1", "Rack_Width") = rack_width
            Parameter("TUBE - Copy:1", "Leg_height") = leg_height

            '''Assembly level params, read only now
            '''needs to update, be actual part params which showed in the form
            Parameter("Rack_Width") = rack_width
            Parameter("Leg_height") = leg_height

            'Renew the table immediately
            iLogicVb.UpdateWhenDone = True
        End Sub
    End Class
End Namespace
