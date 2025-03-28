Imports System
Imports System.Collections.Generic
Imports System.Math
Imports System.Windows.Forms
Imports Autodesk.iLogic.Interfaces
Imports Autodesk.iLogic.Runtime
Imports Autodesk.iLogic.Types
Imports Inventor

Namespace iLogic4VisualStudio
    Public Class Table_Excel_Rule
        Inherits RuleBase

        Public Overrides _
        Sub Main()
            'Path to the Excel file
            Dim excelPath As String = "Table.xlsx"
            Dim sheetName As String = "Query1"

            'Assembly level params
            Dim choosedProject As String
            Dim length As Double
            Dim width As Double

            choosedProject = Parameter("Project")
            GoExcel.FindRow(excelPath, sheetName, "Name", "=", choosedProject)
            length = GoExcel.CurrentRowValue("Length")
            width = GoExcel.CurrentRowValue("Width")
            Parameter("LeftConectionDistance") = length
            Parameter("RightConectionDistance") = width

            'Renew the table immediately
            iLogicVb.UpdateWhenDone = True
        End Sub
    End Class
End Namespace
