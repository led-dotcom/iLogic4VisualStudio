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
            Dim modelCode As String = "FP"

            Dim lengthArray As Integer() = {24, 36, 48, 60, 72, 84, 96, 108, 117}
            Dim depthArray As Integer() = {20}
            Dim heightArray As Integer() = {24, 36, 48}

            For Each ilength As Integer In lengthArray
                For Each idepth As Integer In depthArray
                    For Each iheight As Integer In heightArray
                        ' Change variable values of the parameters

                        ' Wall
                        Parameter("WALL:1", "d0") = ilength
                        Parameter("WALL:1", "d6") = iheight

                        ' Frame
                        Parameter("Channel_U - Copy:1", "d1") = iheight - 0.75
                        Parameter("Channel_U - Copy - back:1", "d1") = iheight + 3.5
                        Parameter("Channel_U:1", "d1") = ilength - 0.625
                        Parameter("Channel_U - Copy - back:1", "d1") = iheight + 3.5
                        Parameter("Channel_U -back h:1", "d1") = ilength - 0.375

                        ' Shelves size
                        Parameter("Top:1", "d2") = ilength - 0.675

                        ' Shelves quantity and position
                        If iheight <= 24 Then
                            Parameter("d101") = 1
                            Parameter("d49") = iheight / 2
                        ElseIf iheight <= 36 Then
                            Parameter("d101") = 2
                            Parameter("d49") = 0
                            Parameter("d99") = iheight / 2
                        Else
                            Parameter("d101") = 3
                            Parameter("d49") = 0
                            Parameter("d99") = iheight / 3
                        End If

                        ' Quantity of shelf panels
                        Parameter("Table Top:1", "d119") = ilength \ 12 - 1

                        ControlUnit(modelCode, ilength, idepth, iheight)
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

        Private Sub ControlUnit(modelCode As String, ilength As Integer, idepth As Integer, iheight As Integer, Optional backSplash As String = "BN", Optional extraShelf As Integer = 0)
            ' Update the unit immediately
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
            Dim saveName As String = modelCode & "_" & ilength & "_" & idepth & "_" & iheight
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
        End Sub
    End Class
End Namespace