Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Word
Imports QRCoder
Imports System.Drawing.Imaging
Imports SWF = System.Windows.Forms
Imports Microsoft.Office.Core
Imports ZXing
Imports System.Runtime.InteropServices.JavaScript.JSType
Imports System.Data.SqlClient
Imports System.Buffers



Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim random As New Random()
        Dim randomNumber As Long = random.Next(100000000, 999999999)
        Dim randNumber As String = ""
        randNumber = randomNumber
        randomNumber = random.Next(100000000, 999999999)
        randNumber &= randomNumber
        randomNumber = random.Next(100000000, 999999999)
        randNumber &= randomNumber
        randomNumber = random.Next(1000, 9999)
        randNumber &= randomNumber

        Dim appPath As String = SWF.Application.StartupPath
        Dim barcodeContent As String = randNumber

        Dim barcodeBitmap As Bitmap = GenerateBarcode(barcodeContent)

        ' 將條碼顯示在 PictureBox 中
        PictureBox1.Image = barcodeBitmap
        PictureBox1.Image.Save(appPath & "Output\3.jpg", Imaging.ImageFormat.Jpeg)

        Dim s0 As String = ""
        Dim s1 As String = ""
        Dim s2 As String = ""

        For x = 1 To 2 Step +1
            Dim randomAscii As Integer = random.Next(65, 91)
            Dim randomChar As Char = Convert.ToChar(randomAscii)
            s0 += randomChar
        Next

        For x = 1 To 53 Step +1
            Dim randomAscii As Integer = random.Next(48, 58)
            Dim randomChar As Char = Convert.ToChar(randomAscii)
            s1 += randomChar
        Next

        For x = 1 To 22 Step +1
            Dim randomAscii As Integer = random.Next(48, 127)
            Dim randomChar As Char = Convert.ToChar(randomAscii)
            s2 += randomChar
        Next

        Dim commodity1 As String = TextBox7.Text


        Dim qrGenerator As New QRCodeGenerator()
        Dim qrCodeData As QRCodeData = qrGenerator.CreateQrCode(s0 + s1 + s2 + "==:**********" + commodity1, QRCodeGenerator.ECCLevel.Q)
        Dim qrCode As New QRCode(qrCodeData)
        Dim qrCodeImage As Bitmap = qrCode.GetGraphic(15)

        ' 将生成的 QR 码显示在 PictureBox 中
        PictureBox1.Image = qrCodeImage
        PictureBox1.Image.Save(appPath & "Output\1.jpg", Imaging.ImageFormat.Jpeg)

        Dim qrGenerator1 As New QRCodeGenerator()
        Dim qrCodeData1 As QRCodeData = qrGenerator1.CreateQrCode(s0 + s1 + s2 + "==:**********", QRCodeGenerator.ECCLevel.Q)
        Dim qrCode1 As New QRCode(qrCodeData1)
        Dim qrCodeImage1 As Bitmap = qrCode1.GetGraphic(15)

        PictureBox1.Image = qrCodeImage1
        PictureBox1.Image.Save(appPath & "Output\2.jpg", Imaging.ImageFormat.Jpeg)
        Dim wordDoc As Microsoft.Office.Interop.Word.Document
        ' 添加新的文件


        Dim CompanyData As String = TextBox5.Text
        If TextBox5.Text = "" Then
            MsgBox(“請輸入商店名!!")
            Exit Sub
        End If
        Dim yData As String = TextBox1.Text
        Dim mData As String = TextBox2.Text
        Dim dData As String = TextBox3.Text
        Dim moneyData As String = TextBox4.Text
        If TextBox5.Text = "" Then
            MsgBox(“請輸入金額!!")
            Exit Sub
        End If
        Dim upData As String = TextBox6.Text
        Dim downData As String = TextBox8.Text
        Dim nowDateTime As DateTime = DateTime.Now
        Dim numberNo As String = ""

        If TextBox9.Text = "" Then
            For x = 1 To 8 Step +1
                Dim randomAscii As Integer = random.Next(48, 58)
                Dim randomChar As Char = Convert.ToChar(randomAscii)
                numberNo += randomChar
            Next
        Else
            numberNo = TextBox9.Text
        End If

        Dim wordApp As New Application()
        ' 顯示 Word 應用程式
        wordApp.Visible = True

        If RadioButton2.Checked Then
            wordDoc = wordApp.Documents.Open(appPath & "WordSample\autoTimes.docx")
            numberNo = "28858068"
        ElseIf RadioButton3.Checked Then
            wordDoc = wordApp.Documents.Open(appPath & "WordSample\auto全國.docx")
            numberNo = "22958907"
        Else
            wordDoc = wordApp.Documents.Open(appPath & "WordSample\auto.docx")
        End If

        Dim FileName As String = "傳統電子發票_" & nowDateTime.ToString("yyyyMMddHHmmss")

        wordApp.Run("載入auto", CompanyData, yData, mData, dData, moneyData, ComboBox1.Text, upData, downData, appPath, numberNo)

        ' 在這裡可以添加更多的操作，如格式設定、插入圖片等


        ' 保存文件，替換 "你的檔案路徑\檔案名稱.docx" 為實際的檔案路徑和名稱

        wordDoc.SaveAs2(appPath & "Output\" & FileName)

        '打印
        'Dim printers As String = ""
        'Try
        '    ' 获取系统上的所有印表機

        '    For Each printer As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
        '        printers &= printer & vbCrLf
        '    Next

        '    ' 在这里选择你要使用的印表機名称
        '    Dim selectedPrinter As String = "YourPrinterName"

        '    ' 使用 PrintOut 方法指定印表機打印文档
        '    wordDoc.PrintOut(selectedPrinter)

        '    ' 如果需要指定其他打印参数，可以使用如下方式：
        '    ' wordDoc.PrintOut(Printer:=selectedPrinter, Background:=False, Append:=False, Range:=WdPrintOutRange.wdPrintAllDocument)

        '    MessageBox.Show("文档已成功使用指定印表機打印。")
        'Catch ex As Exception
        '    MessageBox.Show("打印过程中发生错误：" & ex.Message)
        'End Try


        ' 關閉 Word 文件
        wordDoc.Close()
        wordApp.Quit()
        ' 釋放 Word 對象
        ReleaseObject(wordDoc)
        ReleaseObject(wordApp)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc)
        Marshal.ReleaseComObject(wordApp)
        wordDoc = Nothing
        wordApp = Nothing
    End Sub

    ' 釋放 COM 對象
    Private Sub ReleaseObject(obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = (DateTime.Now.ToString("yyyy")) - 1911
        TextBox2.Text = Int(DateTime.Now.ToString("MM"))
        TextBox3.Text = DateTime.Now.ToString("dd")
        ' 获取系统中所有已安装的字体
        Dim installedFonts() As FontFamily = FontFamily.Families

        ' 遍历并列出每个字体的名称
        For Each font As FontFamily In installedFonts
            ComboBox1.Items.Add(font.Name)
        Next
        RadioButton1.Checked = True

        Dim connectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\Backup\軟體開發\VB程式\WinForms發票系統\WinForms發票系統\Database1.mdf;Integrated Security=True"
        Dim connection As New SqlConnection(connectionString)
        Try
            ' 打开数据库连接
            connection.Open()

            ' 准备 SQL 查询或命令
            Dim sql As String = "SELECT company FROM company;"
            Dim command As New SqlCommand(sql, connection)

            Dim reader As SqlDataReader = command.ExecuteReader()

            ' 将查询结果添加到 ComboBox
            While reader.Read()
                ComboBox2.Items.Add(reader("company").ToString())
            End While

            'MessageBox.Show("數據添加成功！")

        Catch ex As Exception
            MessageBox.Show("發生錯誤：" & ex.Message)

        Finally
            ' 关闭数据库连接
            connection.Close()
        End Try

    End Sub

    Private Function GenerateBarcode(content As String) As Bitmap
        ' 使用 ZXing.Net 生成 Code 128 條碼
        Dim barcodeWriter As New BarcodeWriter
        barcodeWriter.Format = BarcodeFormat.CODE_128
        barcodeWriter.Options = New ZXing.Common.EncodingOptions With {
            .Width = 300, ' 條碼寬度
            .Height = 50, ' 條碼高度3
            .PureBarcode = True
        }

        Dim barcodeBitmap As Bitmap = barcodeWriter.Write(content)
        Return barcodeBitmap
    End Function

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        TextBox5.Enabled = True
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        TextBox5.Text = ""
        TextBox5.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' 设置数据库连接字符串
        Dim connectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\Backup\軟體開發\VB程式\WinForms發票系統\WinForms發票系統\Database1.mdf;Integrated Security=True"
        Dim connection As New SqlConnection(connectionString)
        If TextBox10.Text = "" Then
            MsgBox("輸入資料錯誤!!!")
            Exit Sub
        ElseIf TextBox11.Text = "" Then
            MsgBox("輸入資料錯誤!!!")
            Exit Sub
        End If
        Try
            ' 打开数据库连接
            connection.Open()

            ' 准备 SQL 查询或命令
            Dim sql As String = "INSERT INTO company (company, no) VALUES (@Value1, @Value2)"
            Dim command As New SqlCommand(sql, connection)

            ' 添加参数
            command.Parameters.AddWithValue("@Value1", TextBox10.Text)
            command.Parameters.AddWithValue("@Value2", TextBox11.Text)

            ' 执行命令
            command.ExecuteNonQuery()

            MessageBox.Show("數據添加成功！")

        Catch ex As Exception
            MessageBox.Show("發生錯誤：" & ex.Message)

        Finally
            ' 关闭数据库连接
            connection.Close()
        End Try
        TextBox10.Text = ""
        TextBox11.Text = ""
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox10.Visible = True
            TextBox11.Visible = True
            Label11.Visible = True
            Label12.Visible = True
            Button2.Visible = True
        Else
            TextBox10.Visible = False
            TextBox11.Visible = False
            Label11.Visible = False
            Label12.Visible = False
            Button2.Visible = False
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim connectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\Backup\軟體開發\VB程式\WinForms發票系統\WinForms發票系統\Database1.mdf;Integrated Security=True"
        Dim connection As New SqlConnection(connectionString)
        Try
            ' 打开数据库连接
            connection.Open()
            Dim x
            ' 准备 SQL 查询或命令

            Dim sql As String = "SELECT no FROM company WHERE company = @SearchValue;"
            Dim command As New SqlCommand(sql, connection)

            command.Parameters.AddWithValue("@SearchValue", ComboBox2.Text)

            Dim reader As SqlDataReader = command.ExecuteReader()

            ' 将查询结果添加到 ComboBox

            If reader.HasRows Then
                ' 读取第一行记录
                reader.Read()
                ' 将查询结果填充到 TextBox
                TextBox9.Text = reader("no").ToString()
                TextBox5.Text = ComboBox2.Text
            End If

            'MessageBox.Show("數據添加成功！")

        Catch ex As Exception
            MessageBox.Show("發生錯誤：" & ex.Message)

        Finally
            ' 关闭数据库连接
            connection.Close()
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim random As New Random()

        Dim 收據號碼 As String = "I-" & (TextBox1.Text + 1911) & TextBox2.Text & TextBox3.Text
        For x = 1 To 5 Step +1
            Dim randomAscii As Integer = Random.Next(48, 58)
            Dim randomChar As Char = Convert.ToChar(randomAscii)
            收據號碼 += randomChar
        Next
        Dim QRcode1 As String = 收據號碼 & (TextBox1.Text + 1911) & TextBox2.Text & TextBox3.Text
        For x = 1 To 9 Step +1
            Dim randomAscii As Integer = random.Next(48, 58)
            Dim randomChar As Char = Convert.ToChar(randomAscii)
            QRcode1 += randomChar
        Next


        Dim 運單號碼 As String = "SF14"
        For x = 1 To 11 Step +1
            Dim randomAscii As Integer = random.Next(48, 58)
            Dim randomChar As Char = Convert.ToChar(randomAscii)
            運單號碼 += randomChar
        Next
        運單號碼 += "-1P"
        If TextBox4.Text = "" Then
            MsgBox("金額錯誤!!")
            Exit Sub
        Else
            Dim 金額 As String = TextBox4.Text
        End If

        Dim yData As String = TextBox1.Text
        Dim mData As String = TextBox2.Text
        Dim dData As String = TextBox3.Text
        Dim upData As String = TextBox6.Text
        Dim downData As String = TextBox8.Text
        Dim nowDateTime As DateTime = DateTime.Now


        Dim appPath As String = SWF.Application.StartupPath
        Dim qrGenerator As New QRCodeGenerator()
        Dim qrCodeData As QRCodeData = qrGenerator.CreateQrCode(QRcode1, QRCodeGenerator.ECCLevel.Q)
        Dim qrCode As New QRCode(qrCodeData)
        Dim qrCodeImage As Bitmap = qrCode.GetGraphic(9)

        ' 将生成的 QR 码显示在 PictureBox 中
        PictureBox1.Image = qrCodeImage
        PictureBox1.Image.Save(appPath & "Output\4.jpg", Imaging.ImageFormat.Jpeg)
        Dim wordApp As New Application()
        Dim wordDoc As Microsoft.Office.Interop.Word.Document
        wordApp.Visible = True
        wordDoc = wordApp.Documents.Open(appPath & "WordSample\順豐稅金.docx")
        Dim FileName As String = "SF稅金收據_" & nowDateTime.ToString("yyyyMMddHHmmss")
        wordApp.Run("autoSF", 收據號碼, 運單號碼, yData, mData, dData, TextBox4.Text, upData, downData, appPath)
        wordDoc.SaveAs2(appPath & "Output\" & FileName)
        wordDoc.Close()
        wordApp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc)
        Marshal.ReleaseComObject(wordApp)
        wordDoc = Nothing
        wordApp = Nothing
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        Dim random As New Random()
        Dim randomNumber As Long = random.Next(100000000, 999999999)
        Dim randNumber As String = ""
        randNumber = randomNumber
        randomNumber = random.Next(100000000, 999999999)
        randNumber &= randomNumber
        randomNumber = random.Next(100000000, 999999999)
        randNumber &= randomNumber
        randomNumber = random.Next(1000, 9999)
        randNumber &= randomNumber

        Dim appPath As String = SWF.Application.StartupPath
        Dim barcodeContent As String = randNumber

        Dim barcodeBitmap As Bitmap = GenerateBarcode(barcodeContent)

        ' 將條碼顯示在 PictureBox 中
        PictureBox1.Image = barcodeBitmap
        PictureBox1.Image.Save(appPath & "Output\5.jpg", Imaging.ImageFormat.Jpeg)

        Dim s0 As String = ""
        Dim s1 As String = ""
        Dim s2 As String = ""

        For x = 1 To 2 Step +1
            Dim randomAscii As Integer = random.Next(65, 91)
            Dim randomChar As Char = Convert.ToChar(randomAscii)
            s0 += randomChar
        Next

        For x = 1 To 53 Step +1
            Dim randomAscii As Integer = random.Next(48, 58)
            Dim randomChar As Char = Convert.ToChar(randomAscii)
            s1 += randomChar
        Next

        For x = 1 To 22 Step +1
            Dim randomAscii As Integer = random.Next(48, 127)
            Dim randomChar As Char = Convert.ToChar(randomAscii)
            s2 += randomChar
        Next

        Dim yData As String = TextBox1.Text
        Dim mData As String = TextBox2.Text
        Dim dData As String = TextBox3.Text
        Dim moneyData As String = TextBox4.Text
        If TextBox5.Text = "" Then
            MsgBox(“請輸入金額!!")
            Exit Sub
        End If
        Dim upData As String = TextBox6.Text
        Dim downData As String = TextBox8.Text
        Dim nowDateTime As DateTime = DateTime.Now


        Dim wordApp As New Application()

        ' 顯示 Word 應用程式
        wordApp.Visible = True
        Dim qrGenerator As New QRCodeGenerator()
        Dim qrCodeData As QRCodeData = qrGenerator.CreateQrCode(s0 + s1 + s2 + "==:**********運費:1:" & moneyData, QRCodeGenerator.ECCLevel.Q)
        Dim qrCode As New QRCode(qrCodeData)
        Dim qrCodeImage As Bitmap = qrCode.GetGraphic(15)

        ' 将生成的 QR 码显示在 PictureBox 中
        PictureBox1.Image = qrCodeImage
        PictureBox1.Image.Save(appPath & "Output\6.jpg", Imaging.ImageFormat.Jpeg)

        Dim qrGenerator1 As New QRCodeGenerator()
        Dim qrCodeData1 As QRCodeData = qrGenerator1.CreateQrCode("*******其他費用:1:00", QRCodeGenerator.ECCLevel.Q)
        Dim qrCode1 As New QRCode(qrCodeData1)
        Dim qrCodeImage1 As Bitmap = qrCode1.GetGraphic(7)

        PictureBox1.Image = qrCodeImage1
        PictureBox1.Image.Save(appPath & "Output\7.jpg", Imaging.ImageFormat.Jpeg)
        Dim wordDoc As Microsoft.Office.Interop.Word.Document
        ' 添加新的文件

        wordDoc = wordApp.Documents.Open(appPath & "WordSample\auto順豐發票.docx")

        Dim FileName As String = "SF電子發票_" & nowDateTime.ToString("yyyyMMddHHmmss")

        wordApp.Run("SFinvoiceAuto", yData, mData, dData, moneyData, upData, downData, appPath)

        ' 在這裡可以添加更多的操作，如格式設定、插入圖片等


        ' 保存文件，替換 "你的檔案路徑\檔案名稱.docx" 為實際的檔案路徑和名稱

        wordDoc.SaveAs2(appPath & "Output\" & FileName)

        '打印
        'Dim printers As String = ""
        'Try
        '    ' 获取系统上的所有印表機

        '    For Each printer As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
        '        printers &= printer & vbCrLf
        '    Next

        '    ' 在这里选择你要使用的印表機名称
        '    Dim selectedPrinter As String = "YourPrinterName"

        '    ' 使用 PrintOut 方法指定印表機打印文档
        '    wordDoc.PrintOut(selectedPrinter)

        '    ' 如果需要指定其他打印参数，可以使用如下方式：
        '    ' wordDoc.PrintOut(Printer:=selectedPrinter, Background:=False, Append:=False, Range:=WdPrintOutRange.wdPrintAllDocument)

        '    MessageBox.Show("文档已成功使用指定印表機打印。")
        'Catch ex As Exception
        '    MessageBox.Show("打印过程中发生错误：" & ex.Message)
        'End Try


        ' 關閉 Word 文件
        wordDoc.Close()
        wordApp.Quit()
        ' 釋放 Word 對象
        ReleaseObject(wordDoc)
        ReleaseObject(wordApp)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc)
        Marshal.ReleaseComObject(wordApp)
        wordDoc = Nothing
        wordApp = Nothing
    End Sub
End Class
