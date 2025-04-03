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
            ' Reset all to #430

            ' Check if this file is a drawing
            If TypeOf (ThisApplication.ActiveDocument) IsNot DrawingDocument Then
                MessageBox.Show("Drawing not active!")
                Exit Sub
            End If

            Dim searchs() As String = {"20GSS", "20 GSS", "20G.S.S", "20 G.S.S", "20gss", "20 gss", "20g.s.s", "20 g.s.s"}

            Dim oDoc As DrawingDocument = ThisDoc.Document
            Dim oSheets As Sheets = oDoc.Sheets

            ''' Loop through all sheets and all notes and print the note text to the logger
            For Each oSheet As Sheet In oSheets
                For Each iNote As DrawingNote In oSheet.DrawingNotes.GeneralNotes

                    Dim iText As String = iNote.FormattedText
                    Dim oText As String = ""

                    For Each search As String In searchs
                        If iText.Contains(search) And iText.Contains("#304") Then
                            oText = Replace(iText, "304", "430")
                            Exit For
                        ElseIf iText.Contains(search) And iText.Contains("304") Then
                            oText = Replace(iText, "304", "#430")
                            Exit For
                        ElseIf iText.Contains(search) And Not iText.Contains("430") Then
                            oText = Replace(iText, search, search & " #430")
                            Exit For
                        End If
                    Next

                    'Logger.Info("change?: " & oText)
                    If Not String.IsNullOrEmpty(oText) Then iNote.FormattedText = oText
                Next
            Next

            MsgBox("All reset to #430")
        End Sub
    End Class
End Namespace