Imports System.IO
Imports System.IO.Compression
Imports System.Xml.Linq
Imports System.Threading.Tasks

Public Class Form1

    Private _files As New List(Of String)
    Private _totalPivotCaches As Integer = 0
    Private _totalPivotTables As Integer = 0

    ' -------------------------------------------------
    ' DOSYA / KLASÖR SEÇ
    ' -------------------------------------------------
    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click

        _files.Clear()

        Using dlg As New OpenFileDialog()
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
            dlg.Multiselect = True

            If dlg.ShowDialog() = DialogResult.OK Then
                _files.AddRange(dlg.FileNames)
            End If
        End Using

        If _files.Count = 0 Then
            Using fbd As New FolderBrowserDialog()
                If fbd.ShowDialog() = DialogResult.OK Then
                    _files.AddRange(Directory.GetFiles(fbd.SelectedPath, "*.xlsx"))
                End If
            End Using
        End If

        lblStatus.Text = _files.Count & " dosya seçildi"
        lblResult.Text = ""

    End Sub

    ' -------------------------------------------------
    ' BAŞLAT
    ' -------------------------------------------------
    Private Async Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        If _files.Count = 0 Then
            MessageBox.Show("Önce dosya veya klasör seçin.")
            Return
        End If

        btnStart.Enabled = False
        progressBar1.Value = 0
        progressBar1.Maximum = _files.Count
        progressBar1.Step = 1

        _totalPivotCaches = 0
        _totalPivotTables = 0

        Await Task.Run(Sub()

                           For Each file In _files
                               Try
                                   ProcessXlsx(file)
                               Catch
                                   ' Hatalı dosya atlanır
                               End Try

                               Invoke(Sub()
                                          progressBar1.PerformStep()
                                          lblStatus.Text = System.IO.Path.GetFileName(file)
                                      End Sub)
                           Next

                       End Sub)

        lblStatus.Text = "Tamamlandı"
        lblResult.Text = "Pivot Cache: " & _totalPivotCaches &
                         "   |   Pivot Table: " & _totalPivotTables

        btnStart.Enabled = True
        MessageBox.Show("Tüm dosyalar işlendi.")

    End Sub

    ' -------------------------------------------------
    ' XLSX İŞLE (YEDEK + FIX)
    ' -------------------------------------------------
    Private Sub ProcessXlsx(xlsxPath As String)

        ' ----- YEDEK KLASÖRÜ -----
        Dim backupFolder As String =
            System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(xlsxPath),
                "backup"
            )

        If Not Directory.Exists(backupFolder) Then
            Directory.CreateDirectory(backupFolder)
        End If

        Dim backupName As String =
            System.IO.Path.GetFileNameWithoutExtension(xlsxPath) &
            "_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") &
            ".xlsx"

        Dim backupPath As String =
            System.IO.Path.Combine(backupFolder, backupName)

        File.Copy(xlsxPath, backupPath, True)

        ' ----- TEMP KLASÖR -----
        Dim tempDir As String =
            System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "pivotfix_" & Guid.NewGuid().ToString()
            )

        Directory.CreateDirectory(tempDir)

        ' ----- ZIP AÇ -----
        ZipFile.ExtractToDirectory(xlsxPath, tempDir)

        _totalPivotCaches += FixPivotCaches(tempDir)
        _totalPivotTables += FixPivotTables(tempDir)

        ' ----- ZIP GERİ OLUŞTUR -----
        File.Delete(xlsxPath)

        ZipFile.CreateFromDirectory(
            tempDir,
            xlsxPath,
            CompressionLevel.Optimal,
            False
        )

        Directory.Delete(tempDir, True)

    End Sub

    ' -------------------------------------------------
    ' SAVE SOURCE DATA = OFF
    ' -------------------------------------------------
    Private Function FixPivotCaches(root As String) As Integer

        Dim count As Integer = 0

        Dim folderPath As String =
            System.IO.Path.Combine(root, "xl", "pivotCache")

        If Not Directory.Exists(folderPath) Then Return 0

        For Each file In Directory.GetFiles(folderPath, "pivotCacheDefinition*.xml")

            Dim doc As XDocument = XDocument.Load(file)

            If doc.Root IsNot Nothing Then
                doc.Root.SetAttributeValue("saveData", "0")
                count += 1
            End If

            doc.Save(file)

        Next

        Return count

    End Function

    ' -------------------------------------------------
    ' REFRESH ON OPEN = ON
    ' -------------------------------------------------
    Private Function FixPivotTables(root As String) As Integer

        Dim count As Integer = 0

        Dim folderPath As String =
            System.IO.Path.Combine(root, "xl", "pivotTables")

        If Not Directory.Exists(folderPath) Then Return 0

        For Each file In Directory.GetFiles(folderPath, "pivotTableDefinition*.xml")

            Dim doc As XDocument = XDocument.Load(file)

            If doc.Root IsNot Nothing Then

                Dim attr = doc.Root.Attribute("refreshOnLoad")

                If attr IsNot Nothing Then
                    attr.Value = "1"
                Else
                    doc.Root.Add(New XAttribute("refreshOnLoad", "1"))
                End If

                count += 1
            End If

            doc.Save(file)

        Next

        Return count

    End Function

End Class
