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
            Dim oRefDoc As Document
            For Each oRefDoc In oRefDocs
                Logger.Info(oRefDoc.FullFileName)
            Next
        End Sub
    End Class
End Namespace