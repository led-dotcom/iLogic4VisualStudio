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
            Dim oDoc As DrawingDocument = ThisDoc.Document

            'Dim oSheetOnly As Sheet = oDoc.ActiveSheet
            'Dim oNote As DrawingNote
            'oNote = oSheetOnly.DrawingNotes.Item(1)

            'Dim oText As String = oNote.Text

            'Logger.Info(oText)




            Dim oSheets As Sheets = oDoc.Sheets

            For Each oSheet As Sheet In oSheets
                Dim iNote As DrawingNote
                For Each iNote In oSheet.DrawingNotes
                    Logger.Info(iNote.Text)

                Next


            Next



            'Dim oSheet1 As Sheet
            'oSheet1 = oSheets.Item(2)

            'Dim oNote1 As DrawingNote
            'oNote1 = oSheet1.DrawingNotes.Item(0)


            'Dim oModel As Document
            'oModel = ThisDoc.ModelDocument

            'Dim Result As DialogResult = MessageBox.Show("Turn all sketch symbol leaders on?", "Sketch symbol leaders toggle", MessageBoxButtons.YesNoCancel)
            'Dim viewlabelresult As Boolean = False`
            'If Result = DialogResult.Yes Then
            '    viewlabelresult = True
            'ElseIf Result = DialogResult.Cancel Then
            '    viewlabelresult = False
            'Else
            '    Exit Sub
            'End If



            'Dim oSketchedSymbol As SketchedSymbol
            'oSheets = oDoc.Sheets
            'oSheet = oSheets.Item(1)
            'For Each oSheet In oSheets
            '    For Each oSketchedSymbol In oSheet.SketchedSymbols
            '        oSketchedSymbol.LeaderVisible = True

            '    Next
            'Next
        End Sub
    End Class
End Namespace
