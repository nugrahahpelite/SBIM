Imports System.Data.SqlClient
Imports Spire.Barcode
Imports System.Drawing.Drawing2D
Imports System.Drawing
Module ModGenBarcode
    Dim vsSettings As BarcodeSettings

    Dim vsIOFileStream As System.IO.FileStream
    Dim vsFileLength As Long

    Const csFileFormat = ".jpg"
    Public Sub pbuGenerateQRCode(vriFileName As String, vriDataQRCode As String, ByRef vriQRDir As String)
        'set the configuration of barcode
        vsSettings = New BarcodeSettings()
        Dim data As String = vriDataQRCode
        'Dim type As String = "Code128"
        Dim type As String = "QRCode"

        vsSettings.Data2D = data
        vsSettings.Data = vriDataQRCode

        vsSettings.Type = CType(System.Enum.Parse(GetType(BarCodeType), type), BarCodeType)
        vsSettings.HasBorder = True
        vsSettings.BorderDashStyle = CType(System.Enum.Parse(GetType(DashStyle), "Solid"), DashStyle)

        Dim fontSize As Short = 12
        Dim font As String = "Arial"

        vsSettings.TextFont = New Font(font, fontSize, FontStyle.Bold)

        Dim barHeight As Short = 15

        vsSettings.BarHeight = barHeight

        'settings.X = 1.9
        'settings.Y = 1.9

        vsSettings.ShowText = False
        vsSettings.ShowTextOnBottom = True
        vsSettings.BorderColor = Color.White

        vsSettings.ShowCheckSumChar = True

        'generate the barcode use the settings
        Dim generator As New BarCodeGenerator(vsSettings)
        Dim barcode As Image = generator.GenerateImage()

        vriQRDir = HttpContext.Current.Server.MapPath("~") & "\QRDir\"

        If Dir(vriQRDir & vriFileName) = "" Then
            barcode.Save(vriQRDir & vriFileName)
        End If
    End Sub
End Module
