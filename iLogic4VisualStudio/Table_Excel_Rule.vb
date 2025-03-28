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
            Dim AbsoluteExcelPath As String = "C:\Users\led\Documents\Inventor\exercise\Parts.xlsx"
            Dim RelativeExcelPath As String = "Parts.xlsx"
            Dim sheetName As String = "Sheet1"

            Dim Project As String

            Dim length As Double
            Dim width As Double

            Project = GoExcel.CellValue(AbsoluteExcelPath, sheetName, "A2")
            GoExcel.FindRow(AbsoluteExcelPath, sheetName, "Project", "=", Project)
            length = GoExcel.CurrentRowValue("Length")
            width = GoExcel.CurrentRowValue("Width")
            Parameter("LeftConectionDistance") = length
            Parameter("RightConectionDistance") = width

            'Renew the table immediately
            iLogicVb.UpdateWhenDone = True
        End Sub
    End Class
End Namespace
