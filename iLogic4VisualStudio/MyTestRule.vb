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
            Dim iPropertyQty As String = "9"

            ' Reset QTY as default format

            ' Check if this file is a drawing
            If Not TypeOf (ThisApplication.ActiveDocument) Is DrawingDocument Then
                MessageBox.Show("Drawing not active!")
                Exit Sub
            End If

            Dim searchStr As String = "(QTY="

            Dim oDoc As DrawingDocument = ThisDoc.Document
            Dim oSheets As Sheets = oDoc.Sheets

            ''' Loop through all sheets and all notes and print the text to the logger
            For Each oSheet As Sheet In oSheets
                For Each iNote As DrawingNote In oSheet.DrawingNotes.GeneralNotes

                    Dim iText As String = UCase(iNote.FormattedText)
                    Dim oText As String = ""

                    If iText.Contains(searchStr) Then
                        Dim subStringsArr As String() = Split(iText, "=")
                        Dim isMirrorPart As String = (subStringsArr.Length = 3)

                        Dim numsArr As String() = Split(subStringsArr(1), "X")

                        Dim qtyByUnit As String = numsArr(0)
                        'Dim units As String = numsArr(1)

                        If isMirrorPart Then
                            oText = searchStr & qtyByUnit & "X" & iPropertyQty & "=" & qtyByUnit * iPropertyQty & ")" & " L + R"
                        Else
                            oText = searchStr & qtyByUnit & "X" & iPropertyQty & "=" & qtyByUnit * iPropertyQty & ")"
                        End If

                    End If

                    'Logger.Info("change?: " & oText)
                    If Not String.IsNullOrEmpty(oText) Then iNote.FormattedText = oText
                Next
            Next

            MsgBox("All qty are reset")
        End Sub
    End Class
End Namespace