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
            Dim m_Doc As Inventor.Document = ThisDoc.Document
            'If m_Doc.DocumentType <> kAssemblyDocumentObject And m_Doc.DocumentType <> kPartDocumentObject Then
            '    MessageBox.Show("File is not a model.", "iLogic")
            '    Return 'exit rule
            'End If

            Dim m_Camera As Inventor.Camera = ThisApplication.ActiveView.Camera
            Dim m_TO As Inventor.TransientObjects = ThisApplication.TransientObjects

            Dim oFileName As String = ThisDoc.FileName(False)
            'm_Camera.Perspective = True

            m_Camera.ViewOrientationType = Inventor.ViewOrientationTypeEnum.kIsoTopLeftViewOrientation
            m_Camera.Fit()
            m_Camera.ApplyWithoutTransition()

            Dim m_CV As Inventor.View = ThisApplication.ActiveView

            m_CV.DisplayMode = Inventor.DisplayModeEnum.kShadedRendering
            ThisApplication.DisplayOptions.Show3DIndicator = False

            m_CV.Update()

            ' Create name string for the saved image
            Dim code As String = "WT"
            Dim length As Int16 = Parameter("Length")
            Dim width As Int16 = Parameter("Depth")
            Dim height As Int16 = Parameter("High")

            Dim saveName As String = code & "_" & length & "_" & width & "_" & height

            m_Camera.SaveAsBitmap("C:\Users\di\Desktop\Export\" & saveName & ".png", 1024, 768)
        End Sub
    End Class
End Namespace