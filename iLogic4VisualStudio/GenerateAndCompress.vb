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
            Dim modelCode As String = "UC"

            Dim lengthArray As Integer() = {84, 96, 108, 117}
            Dim depthArray As Integer() = {24, 30}
            Dim heightArray As Integer() = {36}

            Dim backSplashArray As String() = {"BY", "BN"}

            For Each ilength As Integer In lengthArray
                For Each idepth As Integer In depthArray
                    For Each iheight As Integer In heightArray
                        For Each backSplash As String In backSplashArray
                            ' Change variable values of the parameters

                            ' Top cover
                            Parameter("Top:1", "d1") = ilength + 0.5
                            Parameter("Top:1", "d0") = idepth + 0.5

                            ' Top front Channel
                            Parameter("Channel_FT:1", "d1") = ilength + 0.5
                            Parameter("Body:1", "d387") = ilength / 2 - 3

                            ' Top cover support
                            Parameter("channel_TV - Copy:1", "d1") = ilength / 2 - 1
                            Parameter("Channel_Top:3", "d1") = ilength / 2 - 1
                            Parameter("channel_TV:1", "d1") = idepth - 1

                            ' Side panel
                            Parameter("Side_L:1", "d0") = idepth
                            Parameter("Side_L:1", "d1") = iheight

                            ' Back panel
                            Parameter("Back:1", "d1") = ilength
                            Parameter("Back:1", "d0") = iheight

                            ' Middle shelf
                            Parameter("MiddleShelf - Copy:1", "d1") = ilength / 2 - 0.5
                            Parameter("MiddleShelf - Copy:1", "d0") = idepth - 2
                            Parameter("MiddleShelf:1", "d1") = ilength / 2 - 0.5
                            Parameter("MiddleShelf:1", "d0") = idepth - 2
                            ' Middle shelf height
                            'Parameter("Body:1", "d354") = iheight / 2 - 4

                            ' Middle shelf support
                            Parameter("Channel_Middle - Copy:1", "d26") = ilength / 2 - 0.5

                            ' Middle wall
                            Parameter("MiddleWall:1", "d1") = iheight - 2
                            Parameter("MiddleWall:1", "d0") = idepth - 0.5
                            Parameter("Post II :1", "d1") = iheight - 2
                            Parameter("Body:1", "d370") = ilength / 2 - 3

                            ' Bottom shelf
                            Parameter("BottomShelf:1", "d1") = ilength + 0.5
                            Parameter("BottomShelf:1", "d0") = idepth

                            ' Bottom shelf support
                            Parameter("Chennel_Bottom:1", "d1") = ilength - 1

                            ' Leg support
                            Parameter("LegPlate_M:1", "d1") = idepth - 1
                            Parameter("Body:1", "d422") = ilength / 2
                            Parameter("MiddleLeg:1", "d57") = idepth - 4

                            ' Back splash
                            If backSplash = "BY" Then
                                Feature.IsActive("Top:1", "Flange25") = True
                            Else
                                Feature.IsActive("Top:1", "Flange25") = False
                            End If

                            ' Shelf count
                            'If extraShelf = 0 Then
                            '    'Component.IsActive({"Body:1", "Middle_Shelf:1"}) = True
                            '    Parameter("Body:1", "d354") = iheight / 2 - 4
                            '    Parameter("Body:1", "d397") = 1
                            'ElseIf extraShelf = 1 Then
                            '    'Component.IsActive({"Body:1", "Middle_Shelf:1"}) = True
                            '    Parameter("Body:1", "d354") = iheight / 3 - 2
                            '    Parameter("Body:1", "d395") = iheight / 3 - 2
                            '    Parameter("Body:1", "d397") = 2
                            'End If

                            ControlUnit(modelCode, ilength, idepth, iheight, backSplash, 0)
                        Next
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
            Dim saveName As String = modelCode & "_" & ilength & "_" & idepth & "_" & iheight & "_" & backSplash & "_" & "C" & "_" & "ES" & extraShelf
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