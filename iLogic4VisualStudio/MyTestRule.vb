Imports System
Imports System.Collections.Generic
Imports System.Linq
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
			Dim sSourceFile As String = UseFileDialog(False, "Select Source Assembly File.")
			Dim oSourceDoc As Inventor.Document = Nothing
			If IsValidFullFileName(sSourceFile) And System.IO.File.Exists(sSourceFile) Then
				Try : oSourceDoc = ThisApplication.Documents.Open(sSourceFile, False) : Catch : End Try
			Else
				MsgBox("Empty or invalid source file specified.  Exiting routine.", vbCritical, "iLogic")
				Return
			End If
			If oSourceDoc Is Nothing Then
				MsgBox("Source Document could not be opened.  Exiting routine.", vbCritical, "iLogic")
				Return
			End If
			Dim oRefFileDescs As Inventor.FileDescriptorsEnumerator = oSourceDoc.File.ReferencedFileDescriptors
			If oRefFileDescs.Count = 0 Then
				MsgBox("No Referenced FileDescriptors Found.  Exiting routine.", vbCritical, "iLogic")
				Return
			End If
			RecurseFileDescriptors(oRefFileDescs, AddressOf ProcessFileDescriptor)
			oSourceDoc.Update2(True)
			ReplaceComponentsWithRenamedFiles(oSourceDoc)
			oSourceDoc.Update2(True)
			oSourceDoc.Save2(True)
		End Sub

		'list of 'old' file (as entry Key) and 'new' file (as entry Value)
		Dim oRenamedFiles As Dictionary(Of String, String)

		Function UseFileDialog(Optional bSave As Boolean = False, Optional sTitle As String = vbNullString, Optional sInitialDirectory As String = vbNullString, Optional sFileName As String = vbNullString) As String
			Dim oFDlg As Inventor.FileDialog : ThisApplication.CreateFileDialog(oFDlg)
			If Not String.IsNullOrEmpty(sTitle) Then oFDlg.DialogTitle = sTitle
			If String.IsNullOrEmpty(sInitialDirectory) Then
				oFDlg.InitialDirectory = ThisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath
			Else
				oFDlg.InitialDirectory = sInitialDirectory
			End If
			If Not String.IsNullOrEmpty(sFileName) Then oFDlg.FileName = sFileName
			oFDlg.Filter = "Autodesk Inventor Files (*.iam;*.dwg;*.idw;*.ipt) | *.iam;*.dwg;*.idw;*ipt | All files (*.*)|*.*"
			oFDlg.MultiSelectEnabled = False : oFDlg.OptionsEnabled = False
			oFDlg.InsertMode = False : oFDlg.CancelError = True
			Try : If bSave = True Then : oFDlg.ShowSave() : Else : oFDlg.ShowOpen() : End If : Catch : End Try
			Return oFDlg.FileName
		End Function

		Function IsValidFullFileName(sFullFileName As String) As Boolean
			If String.IsNullOrEmpty(sFullFileName) OrElse String.IsNullOrWhiteSpace(sFullFileName) Then Return False
			Dim oInvalidPathChars() As Char = System.IO.Path.GetInvalidPathChars
			Dim sPath As String = System.IO.Path.GetDirectoryName(sFullFileName)
			Dim oInvalidNameChars() As Char = System.IO.Path.GetInvalidFileNameChars
			Dim sName As String = System.IO.Path.GetFileName(sFullFileName)
			If sPath.Intersect(oInvalidPathChars).Any() OrElse sName.Intersect(oInvalidNameChars).Any() Then Return False
			Return True
		End Function

		Sub RecurseFileDescriptors(oFileDescs As FileDescriptorsEnumerator, FileDescriptorProcess As Action(Of Inventor.FileDescriptor))
			If oFileDescs Is Nothing OrElse oFileDescs.Count = 0 Then Return
			For Each oFD As Inventor.FileDescriptor In oFileDescs
				FileDescriptorProcess(oFD)
				Dim oChildFDs As FileDescriptorsEnumerator = Nothing
				If oFD.ReferencedFile IsNot Nothing Then
					Try : oChildFDs = oFD.ReferencedFile.ReferencedFileDescriptors : Catch : End Try
					If oChildFDs IsNot Nothing AndAlso oChildFDs.Count > 0 Then
						RecurseFileDescriptors(oChildFDs, FileDescriptorProcess)
					End If
				End If
			Next 'oFD
		End Sub

		Sub ProcessFileDescriptor(oFD As Inventor.FileDescriptor)
			If oFD Is Nothing Then Return
			If oFD.ReferenceMissing OrElse oFD.ReferenceDisabled Then
				Logger.Warn(vbCrLf & "Reference Missing or Disabled!" & vbCrLf &
				"Last known FullFileName is as follows:" & vbCrLf & oFD.FullFileName & vbCrLf)
				Return
			End If
			'<<< only process it if it is a part >>>
			If oFD.ReferencedFileType <> FileTypeEnum.kPartFileType Then Return
			Dim sFFN As String = oFD.FullFileName
			If sFFN = "" Then
				Logger.Warn("A referenced FileDescriptor was encountered with an 'empty' FullFileName!")
				Return
			End If
			'get file name, without path, but with the file extension
			Dim sFileName As String = System.IO.Path.GetFileName(sFFN)
			Dim sOldPrefix As String = "0980_"
			Dim sNewPrefix As String = "0970_"
			If Not sFileName.StartsWith(sOldPrefix) Then Return
			Dim sNewFileName As String = sFileName.Replace(sOldPrefix, sNewPrefix)
			Dim sNewFFN As String = sFFN.Replace(sFileName, sNewFileName)
			If System.IO.File.Exists(sNewFFN) Then
				'let user know about it
				'		MsgBox("The following file already existed:" & vbCrLf & _
				'		sNewFFN & vbCrLf & _
				'		"So, the following original file was not renamed:" & vbCrLf & sFFN, vbExclamation, "iLogic")
				Logger.Warn(vbCrLf & "The following file already existed:" & vbCrLf &
				sNewFFN & vbCrLf &
				"So, the following original file was not renamed:" & vbCrLf & sFFN)
			Else 'a file with the new name does not already exist, so rename it
				Try
					FileSystem.Rename(sFFN, sNewFFN)
					If oRenamedFiles Is Nothing Then
						oRenamedFiles = New Dictionary(Of String, String)
						oRenamedFiles.Add(sFFN, sNewFFN)
					Else
						oRenamedFiles.Add(sFFN, sNewFFN)
					End If
					'now you may have to fix the file references referring to the old file name
					oFD.ReplaceReference(sNewFFN)
					Logger.Info(vbCrLf & "The following 'original' file:" & vbCrLf &
					sFFN & vbCrLf &
					"was renamed to the following:" & vbCrLf & sNewFFN)
				Catch
					'what to do if that failed
				End Try
			End If
		End Sub

		Sub ReplaceComponentsWithRenamedFiles(oAsmDoc As AssemblyDocument)
			If (oAsmDoc Is Nothing) OrElse (TypeOf oAsmDoc Is AssemblyDocument = False) Then Return
			If oRenamedFiles Is Nothing OrElse oRenamedFiles.Count = 0 Then Return
			Dim oOccs As ComponentOccurrences = oAsmDoc.ComponentDefinition.Occurrences
			For Each oEntry As KeyValuePair(Of String, String) In oRenamedFiles
				Dim sOldFFN As String = oEntry.Key
				Dim sNewFFN As String = oEntry.Value
				Dim oRefOccs As ComponentOccurrencesEnumerator = Nothing
				Try : oRefOccs = oOccs.AllReferencedOccurrences(sOldFFN) : Catch : End Try
				If oRefOccs Is Nothing OrElse oRefOccs.Count = 0 Then Continue For
				For Each oRefOcc As ComponentOccurrence In oRefOccs
					Try : oRefOcc.Replace2(sNewFFN, True, , True) : Catch : End Try
				Next 'oRefOcc
			Next 'oEntry
		End Sub
	End Class
End Namespace