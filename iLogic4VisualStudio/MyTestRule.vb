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
            Dim oSheets As Sheets = oDoc.Sheets

            ''' Loop through all sheets and all notes
            ''' and print the note text to the logger
            ''' i is the sheet number
            ''' j is the note number
            Dim i As Integer = 1
            Dim j As Integer = 1

            For Each oSheet As Sheet In oSheets
                For Each iNote As DrawingNote In oSheet.DrawingNotes
                    Logger.Info(i & ", " & j & "," & iNote.Text)

                    ''' Search for the text "GSS"
                    ''' in the note text
                    Dim iText As String = iNote.Text
                    Dim search As String = "GSS"

                    If iText.Contains(search) Then
                        Dim oText As String = Replace(iText, search, search & "!!!!")

                        iNote.FormattedText = oText
                    End If

                    j += 1
                Next

                j = 1
                i += 1
            Next
        End Sub
    End Class
End Namespace