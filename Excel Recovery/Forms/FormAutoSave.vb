Option Explicit On

Imports System.IO
Imports System.Drawing
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Imports Microsoft.Win32

Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormAutoSave

#Region "Fields"

    Protected _bBusy As Boolean
    Protected _bStop As Boolean

#End Region

#Region "Events"

    Public Sub New()

        InitializeComponent()

    End Sub

    Private Sub FormAutoSave_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        gridFiles.ColumnHeadersDefaultCellStyle.BackColor = Color.Black
        gridFiles.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        gridFiles.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font(gridFiles.Font, FontStyle.Bold)

        gridFiles.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        gridFiles.ColumnHeadersHeight = 30
        gridFiles.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised

    End Sub

    Private Sub FormAutoSave_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        If _bBusy = True Then

            e.Cancel = True

        End If

    End Sub

    Private Sub btnOpenFile_Click(sender As System.Object, e As System.EventArgs) Handles btnOpenFile.Click

        OpenFile()

    End Sub

    Private Sub btnSaveAs_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveAs.Click

        SaveFile()

    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click

    End Sub

    Private Sub gridFiles_CellDoubleClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridFiles.CellDoubleClick

        Try

            If e.RowIndex >= 0 Then

                OpenFile()

            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click

        Dim dlgFolderBrowser As FolderBrowserDialog = New FolderBrowserDialog()

        If dlgFolderBrowser.ShowDialog() = DialogResult.OK Then

            txtCustomLocation.Text = dlgFolderBrowser.SelectedPath

        End If

    End Sub

    Private Sub chkbSearchBackup_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkbSearchBackup.CheckedChanged

        Handle_CheckBox_CheckedChanged()

    End Sub

    Private Sub chkbSearchXAR_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkbSearchXAR.CheckedChanged

        Handle_CheckBox_CheckedChanged()

    End Sub

    Private Sub chkbSearchXLK_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkbSearchXLK.CheckedChanged

        Handle_CheckBox_CheckedChanged()

    End Sub

    Private Sub btnStart_Click(sender As System.Object, e As System.EventArgs) Handles btnStart.Click

        If _bBusy = True Then
            _bStop = True
        Else
            StartSearch()
        End If

    End Sub

    Private Sub chkbTextSearch_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkbTextSearch.CheckedChanged

        txtTextSearch.Enabled = chkbTextSearch.Checked

    End Sub

#End Region

#Region "Implementaion"

    Private Sub StartSearch()

        _bBusy = True
        _bStop = False

        Try

            Dim bUserLocation As Boolean = chkbSearchXAR.Checked Or chkbSearchXLK.Checked

            If chkbSearchBackup.Checked = True Or bUserLocation = True Then

                Dim bContinue As Boolean = True

                If bUserLocation = True Then

                    If String.IsNullOrEmpty(txtCustomLocation.Text) = True Or Directory.Exists(txtCustomLocation.Text) = False Then
                        MessageBox.Show("Please specify a valid location to search for auto-saved files", FormMain.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        bContinue = False
                    End If

                End If

                If bContinue = True Then

                    Dim nSteps As Integer = 0

                    gridFiles.Rows.Clear()

                    '--------------------------------------------------------------

                    Dim arrPaths As ArrayList = Nothing

                    If chkbSearchBackup.Checked = True Then
                        arrPaths = GetExcelBackupPaths()
                        nSteps = nSteps + arrPaths.Count
                    End If

                    '--------------------------------------------------------------

                    If chkbSearchXAR.Checked = True Then
                        nSteps = nSteps + 1
                    End If

                    If chkbSearchXLK.Checked = True Then
                        nSteps = nSteps + 1
                    End If

                    '--------------------------------------------------------------

                    UpdateControls(True, nSteps)

                    '--------------------------------------------------------------

                    If chkbSearchBackup.Checked Then
                        For Each strPath As String In arrPaths

                            If chkbSearchBackup.Checked = True Then

                                FindFiles_TMP(strPath)

                            End If

                            progress.PerformStep()

                            System.Windows.Forms.Application.DoEvents()

                        Next

                    End If

                    '--------------------------------------------------------------

                    System.Windows.Forms.Application.DoEvents()

                    If _bStop = False Then
                        If chkbSearchXAR.Checked = True Then
                            FindFiles_Custom(txtCustomLocation.Text, "*.XAR")
                        End If

                        progress.PerformStep()

                    End If

                    System.Windows.Forms.Application.DoEvents()

                    If _bStop = False Then
                        If chkbSearchXLK.Checked = True Then
                            FindFiles_Custom(txtCustomLocation.Text, "*.XLK")
                        End If

                        progress.PerformStep()

                    End If

                    System.Windows.Forms.Application.DoEvents()

                    '--------------------------------------------------------------

                    UpdateControls(False, -1)

                    Me.Invalidate()
                    Me.Update()

                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        _bBusy = False
        _bStop = False

    End Sub

    Private Sub FindFiles_TMP(strPath As String)

        Try

            Dim nRow As Integer
            Dim row As DataGridViewRow

            If Directory.Exists(strPath) Then

                Dim di As DirectoryInfo = New DirectoryInfo(strPath)
                Dim arrFileInfo As FileInfo() = di.GetFiles("~DF*.TMP")

                Dim strDateTime As String

                For Each fi As FileInfo In arrFileInfo

                    System.Windows.Forms.Application.DoEvents()

                    If _bStop = False Then

                        If IsMatch(fi.FullName) = True Then
                            strDateTime = fi.LastWriteTime.ToShortDateString() + "   " + fi.LastWriteTime.ToLongTimeString()

                            nRow = gridFiles.Rows.Add(fi.Name, strDateTime, fi.DirectoryName)
                            If (nRow <> -1) Then

                                row = gridFiles.Rows(nRow)
                                row.Tag = fi.FullName

                            End If

                        End If

                    Else

                        Exit For

                    End If

                Next

            End If

        Catch
        End Try

    End Sub

    Private Sub FindFiles_Custom(strPath As String, strPattern As String)

        Try

            Dim nRow As Integer
            Dim row As DataGridViewRow

            If Directory.Exists(strPath) Then

                Dim di As DirectoryInfo = New DirectoryInfo(strPath)

                Dim arrDirInfo As DirectoryInfo() = di.GetDirectories()

                For Each diSub In arrDirInfo

                    System.Windows.Forms.Application.DoEvents()

                    If _bStop = False Then

                        FindFiles_Custom(diSub.FullName, strPattern)

                    Else

                        Exit For

                    End If

                Next

                '---------------------------------------------------------------------------

                Dim arrFileInfo As FileInfo() = di.GetFiles(strPattern)

                Dim strDateTime As String

                For Each fi As FileInfo In arrFileInfo

                    System.Windows.Forms.Application.DoEvents()

                    If _bStop = False Then

                        If IsMatch(fi.FullName) = True Then

                            strDateTime = fi.LastWriteTime.ToShortDateString() + "   " + fi.LastWriteTime.ToLongTimeString()

                            nRow = gridFiles.Rows.Add(fi.Name, strDateTime, fi.DirectoryName)
                            If (nRow <> -1) Then

                                row = gridFiles.Rows(nRow)
                                row.Tag = fi.FullName

                            End If

                        End If

                    Else

                        Exit For

                    End If

                Next

            End If

        Catch
        End Try

    End Sub

    Private Function IsMatch(strPath As String) As Boolean

        Dim bResult As Boolean = True

        Try

            If chkbTextSearch.Checked = True Then

                Dim strText As String = txtTextSearch.Text

                If String.IsNullOrEmpty(strText) = False Then

                    Dim strContent = File.ReadAllText(strPath)

                    If String.IsNullOrEmpty(strContent) = False Then

                        If strContent.IndexOf(strText, StringComparison.InvariantCultureIgnoreCase) = -1 Then

                            bResult = False

                        End If

                    Else

                        bResult = False

                    End If

                End If

            End If

        Catch
        End Try

        Return bResult

    End Function

    Private Sub OpenFile()

        Try

            If gridFiles.SelectedRows.Count = 1 Then

                Dim row As DataGridViewRow = gridFiles.SelectedRows(0)
                Dim strFilePath = CType(row.Tag, String)

                Dim objExcel As New Excel.Application
                objExcel.Visible = True

                objExcel.Workbooks.Open(strFilePath)

                Try
                    objExcel.ActiveWindow.WindowState = XlWindowState.xlMaximized
                    objExcel.ActiveWindow.Activate()
                Catch
                End Try

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub SaveFile()

        Try

            If gridFiles.SelectedRows.Count = 1 Then

                Dim row As DataGridViewRow = gridFiles.SelectedRows(0)
                Dim strFilePath = CType(row.Tag, String)

                Dim dlgSaveFile As New SaveFileDialog()
                dlgSaveFile.Filter = "Excel Workbooks (*.xlsx; *.xls)|*.xlsx;*.xls|All files (*.*)|*.*||"
                dlgSaveFile.FileName = "Recovered Workbook"

                If dlgSaveFile.ShowDialog() = DialogResult.OK Then

                    File.Copy(strFilePath, dlgSaveFile.FileName)

                    MessageBox.Show("File saved successfully", FormMain.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub UpdateControls(bBusy As Boolean, nSteps As Integer)

        Try

            If bBusy = True Then

                btnOpenFile.Enabled = False
                btnSaveAs.Enabled = False
                btnCancel.Enabled = False

                progress.Visible = True

                progress.Value = 1
                progress.Maximum = nSteps + 1

                btnStart.Text = "Stop"

                '-----------------------------------------------------------

                chkbSearchBackup.Enabled = False
                chkbSearchXAR.Enabled = False
                chkbSearchXLK.Enabled = False

                lblCustomLocation.Enabled = False
                txtCustomLocation.Enabled = False
                btnBrowse.Enabled = False

            Else

                btnOpenFile.Enabled = True
                btnSaveAs.Enabled = True
                btnCancel.Enabled = True
                btnStart.Enabled = True

                progress.Visible = False

                btnStart.Text = "Start"

                '-----------------------------------------------------------

                chkbSearchBackup.Enabled = True
                chkbSearchXAR.Enabled = True
                chkbSearchXLK.Enabled = True

                lblCustomLocation.Enabled = True
                txtCustomLocation.Enabled = True
                btnBrowse.Enabled = True

            End If

            System.Windows.Forms.Application.DoEvents()

        Catch
        End Try

    End Sub

    Private Function GetExcelBackupPaths() As ArrayList

        Dim arrPaths As New ArrayList()

        Dim strDefault_WindowsXP As String = "C:\Documents and Settings\" + Environment.UserName + "\Local Settings\Temp"
        Dim strDefault_WindowsXP2 As String = "C:\Documents and Settings\" + Environment.UserName + "\Application Data\Microsoft\Excel"
        Dim strDefault_WindowsVista As String = "C:\Users\" + Environment.UserName + "\AppData\Local\Temp"
        Dim strDefault_Windows7 As String = "C:\Users\" + Environment.UserName + "\AppData\Roaming\Microsoft\Excel"

        Dim strUserModifiedValue As String = ""

        '------------------------------------------------------------------------------------------

        Try

            Dim key As RegistryKey

            '------------------------------------------------------------------------------------------

            key = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Office\11.0\Excel\Options", False)
            If key IsNot Nothing Then

                strUserModifiedValue = key.GetValue("AutoRecoverPath", "")
                If String.IsNullOrEmpty(strUserModifiedValue) = False Then
                    arrPaths.Add(strUserModifiedValue)
                End If

            End If

            '------------------------------------------------------------------------------------------

            key = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Office\12.0\Excel\Options", False)
            If key IsNot Nothing Then

                strUserModifiedValue = key.GetValue("AutoRecoverPath", "")
                If String.IsNullOrEmpty(strUserModifiedValue) = False Then
                    arrPaths.Add(strUserModifiedValue)
                End If

            End If

            '------------------------------------------------------------------------------------------

            key = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Office\14.0\Excel\Options", False)
            If key IsNot Nothing Then

                strUserModifiedValue = key.GetValue("AutoRecoverPath", "")
                If String.IsNullOrEmpty(strUserModifiedValue) = False Then
                    arrPaths.Add(strUserModifiedValue)
                End If

            End If

        Catch
        End Try

        '------------------------------------------------------------------------------------------

        If Environment.OSVersion.Version.Major >= 6 Then

            If Environment.OSVersion.Version.Minor = 1 Then

                arrPaths.Add(strDefault_Windows7)

            Else

                arrPaths.Add(strDefault_WindowsVista)

            End If

        Else

            arrPaths.Add(strDefault_WindowsXP)
            arrPaths.Add(strDefault_WindowsXP2)

        End If

        '------------------------------------------------------------------------------------------

        Return arrPaths

    End Function

    Private Sub Handle_CheckBox_CheckedChanged()

        Dim bUserLocation As Boolean = chkbSearchXAR.Checked Or chkbSearchXLK.Checked

        lblCustomLocation.Enabled = bUserLocation
        txtCustomLocation.Enabled = bUserLocation
        btnBrowse.Enabled = bUserLocation

        If chkbSearchBackup.Checked = False And bUserLocation = False Then

            chkbSearchBackup.Checked = True

        End If

    End Sub

#End Region

End Class
