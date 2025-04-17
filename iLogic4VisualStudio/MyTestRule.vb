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

                    Dim iFormattedText As String = iNote.FormattedText
                    Dim iText As String = UCase(iNote.Text)

                    Dim oFormattedText As String = ""
                    Dim oText As String = ""

                    Logger.Info("iFormattedText: " & iFormattedText)

                    If iText.Contains(searchStr) Then
                        ''' get find string in formatted text
                        Dim leftIndex As Integer = iFormattedText.IndexOf("(")
                        Dim rightIndex As Integer = iFormattedText.IndexOf(")")

                        Dim length As Integer = rightIndex - leftIndex + 1

                        Dim findTXT As String = iFormattedText.Substring(leftIndex, length)

                        ''' get qty per unit
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

                        oFormattedText = Replace(iFormattedText, findTXT, oText)

                    End If

                    'Logger.Info("change?: " & oText)
                    If Not String.IsNullOrEmpty(oFormattedText) Then iNote.FormattedText = oFormattedText
                Next
            Next

            MsgBox("All qty are reset")
        End Sub
    End Class
End Namespace