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
            ' parameters of the model
            Dim modelCode As String = "WT"

            Dim lengthArray As Integer() = {24, 36, 48, 60, 72, 84, 96, 108, 120}
            Dim depthArray As Integer() = {24, 30, 36}
            Dim heightArray As Integer() = {24, 30, 36}


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

                        ' When length is greater than 80, unit has 6 legs adjust the table bottom width
                        If ilength <= 80 Then
                            Parameter("Table Bottom:1", "d101") = 1
                        Else
                            Parameter("Table Bottom:1", "d101") = 2
                            Parameter("Table Bottom:1", "d99") = (ilength - 4) / 2
                        End If

                        ' Update the table immediately
                        InventorVb.DocumentUpdate()

                        ' Update the Camera
                        Dim m_Camera As Inventor.Camera = ThisApplication.ActiveView.Camera

                        'm_Camera.Perspective = True
                        m_Camera.ViewOrientationType = Inventor.ViewOrientationTypeEnum.kIsoTopLeftViewOrientation
                        m_Camera.Fit()
                        m_Camera.ApplyWithoutTransition()

                        ' Update the view to apply the camera settings
                        Dim m_CV As Inventor.View = ThisApplication.ActiveView

                        m_CV.DisplayMode = Inventor.DisplayModeEnum.kShadedRendering
                        ThisApplication.DisplayOptions.Show3DIndicator = False

                        m_CV.Update()

                        ' Create name string for the saved image
                        ' Example:
                        ' WT_24_30_36_BY_C
                        ' WT = modelCode
                        ' 24 = length
                        ' 30 = depth
                        ' 36 = height
                        ' BY = back splash BN = back no splash
                        ' C = casters L = legs
                        ' ES = extra shelf
                        Dim saveName As String = modelCode & "_" & ilength & "_" & idepth & "_" & iheight & "_" & "BY" & "_" & "C" & "_" & "ES" & "0"
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