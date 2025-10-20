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
            ShowReferences()
        End Sub

        Public Sub ShowReferences()
            ' Get the active assembly.
            Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument
            ' Get all of the referenced documents.
            Dim oRefDocs As DocumentsEnumerator = oAsmDoc.AllReferencedDocuments

            ' Iterate through the list of documents.
            Dim oRefDoc As PartDocument
            For Each oRefDoc In oRefDocs
                Logger.Info(oRefDoc.FullFileName)

                Dim newName As String = System.IO.Path.GetFileNameWithoutExtension(oRefDoc.FullFileName) & "_Copy" & System.IO.Path.GetExtension(oRefDoc.FullFileName)
                Logger.Info(newName)
                ' Copy the document to a new file in the same folder.
                oRefDoc.SaveAs(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(oRefDoc.FullFileName), newName), False)
            Next
        End Sub
    End Class
End Namespace