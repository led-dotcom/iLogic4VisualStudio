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
            ' Create a length array of 3 integers
            Dim lengthArray As Integer() = {24, 36, 48}

            For Each ilength As Integer In lengthArray
                ' Change variable values of the parameters
                Dim excelSheet = "J:\2025\25-0319 GEORGE P\0319+3 TABLE\Data\Parameter.xls"
                WriteExcel(excelSheet, "Sheet1", "B4", ilength)

                'Renew the table immediately
                'iLogicVb.UpdateWhenDone = True

                ' Use RuleParametersOutput function if you must perform an Update using DocumentUpdate.
                InventorVb.CheckParameters("")
                RuleParametersOutput()
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
                Dim length As Int16 = Parameter("Length")
                Dim width As Int16 = Parameter("Depth")
                Dim height As Int16 = Parameter("High")

                Dim saveName As String = code & "_" & length & "_" & width & "_" & height
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

        Private Sub WriteExcel(ByVal filePath As String, ByVal sheetName As String, ByVal cellAddress As String, ByVal val As String)
            Try
                GoExcel.Open(filePath, sheetName)
                GoExcel.CellValue(cellAddress) = val
                GoExcel.Save()
                GoExcel.Close()
            Catch ex As Exception
                MessageBox.Show("Error writing to Excel: " & ex.Message, "Excel Error")
            End Try
        End Sub
    End Class
End Namespace