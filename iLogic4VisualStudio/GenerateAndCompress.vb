Imports System
Imports System.Collections.Generic
Imports System.Math
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports Autodesk.iLogic.Interfaces
Imports Autodesk.iLogic.Runtime
Imports Autodesk.iLogic.Types
Imports Inventor

Namespace iLogic4VisualStudio
    Public Class GenerateAndCompress
        Inherits RuleBase

        Public Overrides _
        Sub Main()
            Dim lengthArray As Integer() = {24, 36, 48}
            Dim depthArray As Integer() = {24, 36, 48}
            Dim heightArray As Integer() = {24, 36, 48}

            For Each ilength As Integer In lengthArray
                For Each idepth As Integer In depthArray
                    For Each iheight As Integer In heightArray
                        ' Change variable values of the parameters
                        Parameter("Top:1", "d2") = ilength
                        Parameter("Undershelf:1", "d1") = ilength - 4
                        Parameter("Channel_H:2", "d21") = ilength - 8 - 2 * 0.0625 - 0.0625
                        Parameter("Channel_U:1", "d1") = ilength - 4.5

                        Parameter("Top:1", "d1") = idepth
                        Parameter("Undershelf:1", "d0") = idepth - 4
                        Parameter("Channel_V:1", "d46") = idepth - 0.625

                        Parameter("Leg:1", "d1") = iheight - 3

                        'Renew the table immediately
                        InventorVb.DocumentUpdate()

                        Dim m_Doc As Inventor.Document = ThisDoc.Document

                        Dim m_Camera As Inventor.Camera = ThisApplication.ActiveView.Camera
                        Dim m_TO As Inventor.TransientObjects = ThisApplication.TransientObjects

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

                        Dim saveName As String = code & "_" & ilength & "_" & idepth & "_" & iheight
                        Dim exportPath As String = "C:\Users\di\Desktop\Export\"
                        Dim tempImagePath As String = System.IO.Path.Combine(exportPath, saveName & "_temp.png")
                        Dim finalImagePath As String = System.IO.Path.Combine(exportPath, saveName & ".jpg")

                        ' First save the image as temporary PNG
                        m_Camera.SaveAsBitmap(tempImagePath, 1024, 768)

                        ' Compress and convert to JPEG
                        CompressAndSaveImage(tempImagePath, finalImagePath, 75) ' 75% quality

                        ' Clean up temporary file
                        If System.IO.File.Exists(tempImagePath) Then
                            System.IO.File.Delete(tempImagePath)
                        End If
                    Next
                Next
            Next
        End Sub

        Private Sub CompressAndSaveImage(tempPath As String, finalPath As String, quality As Long)
            Try
                ' Load the temporary image
                Using originalImage As Bitmap = New Bitmap(tempPath)
                    ' Create JPEG encoder with quality parameter
                    Dim jpegEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)
                    Dim encoderParams As New EncoderParameters(1)
                    encoderParams.Param(0) = New EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality)

                    ' Save compressed image
                    originalImage.Save(finalPath, jpegEncoder, encoderParams)
                End Using
            Catch ex As Exception
                MessageBox.Show("Error compressing image: " & ex.Message, "Compression Error")
            End Try
        End Sub

        Private Function GetEncoder(format As ImageFormat) As ImageCodecInfo
            Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageEncoders()
            For Each codec As ImageCodecInfo In codecs
                If codec.FormatID = format.Guid Then
                    Return codec
                End If
            Next
            Return Nothing
        End Function
    End Class
End Namespace