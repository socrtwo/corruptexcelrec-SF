Option Explicit On

Imports System.IO
Imports System.Drawing
Imports System.Windows
Imports System.Reflection
Imports System.Collections
Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormMain

#Region "DIMs"

    Dim filename As String
    Dim counterVariable As Integer
    Dim previousVersionCounterVariable As Integer
    Dim saveShadowPath As String
    Dim sFileShadowPath As String
    Dim sFileShadowName As String
    Dim sFileShadowSize As String
    Dim sFileShadowPathDate As String
    Dim selectedsFileShadowPathDate As String
    Dim selectedsFileShadowPathSize As String
    Dim selectedPreviousVersion As String
    Dim pathToComboBoxSelectedFile As String
    Dim shadowLinkFolderName As New List(Of String)
    Dim nonErrorShadowPathList As New List(Of String)
    Dim comboBoxIndex As Integer = 0
    Dim matchCount As Integer = 0
    Dim comboBoxChoiceIndex As Integer = 0
    Dim pathToComboBoxSelectedFileSize As Integer = 0
    Dim preVersionHashTable As New Hashtable


#End Region

#Region "Functions"

    Public Function InitProgressForm(ByVal nSteps As Integer) As FormProgress

        Dim formProgress As New FormProgress()

        formProgress.TopMost = True

        'formProgress.progress.Maximum = nSteps

        formProgress.Show()

        System.Windows.Forms.Application.DoEvents()

        Return formProgress

    End Function

    Public Function PerformProgressStep(ByVal formProgress As FormProgress) As Boolean

        System.Windows.Forms.Application.DoEvents()

        If (formProgress._bStop = False) Then

            'formProgress.progress.PerformStep()

            System.Windows.Forms.Application.DoEvents()

            Return True

        Else

            Return False

        End If

    End Function

    Public Function SaveTextToFile(ByVal strData As String, _
     ByVal FullPath As String, _
       Optional ByVal ErrInfo As String = "") As Boolean


        Dim bAns As Boolean = False
        Dim objReader As StreamWriter
        Try


            objReader = New StreamWriter(FullPath)
            objReader.Write(strData)
            objReader.Close()
            bAns = True
        Catch Ex As Exception
            ErrInfo = Ex.Message

        End Try
        Return bAns
    End Function
    Public Function DelFromRight(ByVal sChars As String, ByVal sLine As String) As String

        'Removes unwanted characters from right of given string
        ' EXAMPLE
        '  MsgBox DelFromRight(" TEST", "THIS IS A TEST")
        'displays "THIS IS A"

        sLine = ReverseString(sLine)
        sChars = ReverseString(sChars)
        sLine = DelFromLeft(sChars, sLine)
        DelFromRight = ReverseString(sLine)

        Exit Function


    End Function

    Public Function DelFromLeft(ByVal sChars As String, _
            ByVal sLine As String) As String

        ' Removes unwanted characters from left of given string
        '  EXAMPLE
        '      MsgBox DelFromLeft("THIS", "THIS IS A TEST")
        '        displays  "IS A TEST"


        Dim iCount As Integer
        Dim sChar As String

        DelFromLeft = ""
        ' Remove unwanted characters to left of folder name
        If InStr(sLine, sChars) > 0 Then
            For iCount = 1 To Len(sChars)
                ' Retrieve character from start string to 
                'look for in folder string (sLine)
                sChar = Mid$(sChars, iCount, 1)
                ' Remove all characters to left of found string
                sLine = Mid$(sLine, InStr(sLine, sChar) + 1)

            Next iCount
        End If
        DelFromLeft = sLine
        Exit Function

    End Function

    Public Function ReverseString(ByVal InputString As String) _
      As String

        'If you have vb6, you can use
        'StrReverse instead of this function

        Dim lLen As Long, lCtr As Long
        Dim sChar As String
        Dim sAns As String = ""

        lLen = Len(InputString)
        For lCtr = lLen To 1 Step -1
            sChar = Mid(InputString, lCtr, 1)
            sAns = sAns & sChar
        Next

        ReverseString = sAns

    End Function
    Private Function ReadExeFromResources(ByVal filename As String) As Byte()
        Dim CurrentAssembly As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim Resource As String = String.Empty
        Dim ArrResources As String() = CurrentAssembly.GetManifestResourceNames()
        For Each Resource In ArrResources
            If Resource.IndexOf(filename) > -1 Then Exit For
        Next
        Dim ResourceStream As IO.Stream = CurrentAssembly.GetManifestResourceStream(Resource)
        If ResourceStream Is Nothing Then
            Return Nothing
        End If
        Dim ResourcesBuffer(CInt(ResourceStream.Length) - 1) As Byte
        ResourceStream.Read(ResourcesBuffer, 0, ResourcesBuffer.Length)
        ResourceStream.Close()
        Return ResourcesBuffer
    End Function


#End Region

#Region "Fields"

    Protected Shared _strAppName As String = "Excel Recovery"

#End Region

#Region "Properties"

    Public Shared Property AppName As String
        Get
            Return _strAppName
        End Get
        Set(value As String)

        End Set
    End Property

#End Region

#Region "Events"

    Private Sub FormMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        System.Windows.Forms.Application.EnableVisualStyles()

    End Sub

    Private Sub picMinimizeBox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub picXToCloseBox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

        Using myProcess As Process = New Process

            Dim x As Integer
            Dim myStreamWriter As StreamWriter = Nothing

            Try

                For x = 0 To shadowLinkFolderName.Count - 1

                    'Starts a command line in the background and removes the 
                    'temporary folders mapped to the restore point snapshots.

                    myProcess.StartInfo.FileName = "cmd.exe"
                    myProcess.StartInfo.UseShellExecute = False
                    myProcess.StartInfo.RedirectStandardInput = True
                    myProcess.StartInfo.RedirectStandardOutput = True
                    myProcess.StartInfo.CreateNoWindow = True
                    myProcess.Start()
                    myStreamWriter = myProcess.StandardInput
                    myStreamWriter.WriteLine("rmdir " & shadowLinkFolderName(x))
                    myStreamWriter.Flush()
                    myStreamWriter.Close()
                    myStreamWriter = Nothing
                    myProcess.WaitForExit()
                    myProcess.Close()

                Next

                Me.Close()

                TerminateProcess(myProcess.ProcessName, myProcess.Id)

            Catch ex As Exception

                MessageBox.Show(ex.Message)

            End Try

        End Using

    End Sub

    Private Sub picFolderFileChooserPreviousVersionSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Impl_FolderFileChooserPreviousVersionSearch()

    End Sub

    Private Sub picPreviousVersionFileChoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Impl_PreviousVersionFileChoice()

    End Sub

    Private Sub picSaveFileFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Impl_SaveFileFolder()

    End Sub

    Private Sub picSaveAsSYLK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox10.Click

        Impl_SaveAsSYLK()

    End Sub

    Private Sub picSaveAsHTML_Click(sender As System.Object, e As System.EventArgs) Handles PictureBox3.Click

        Impl_SaveAsHTML()

    End Sub

    Private Sub pickManualCalculations_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click

        Impl_ManualCalculations()

    End Sub

    Private Sub picExternalReferenceRecoveryMethod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click

        Impl_ExternalReferenceRecoveryMethod()

    End Sub

    Private Sub picOpenInSafeMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click

        Impl_OpenInSafeMode()

    End Sub

    Private Sub picExtractDataFromChart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox16.Click

        Impl_ExtractDataFromChart()

    End Sub

    Private Sub picOpenInWordPad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox11.Click

        Impl_OpenInWordPad()

    End Sub

    Private Sub picOpenInWord_Click(sender As Object, e As EventArgs) Handles PictureBox17.Click

        Impl_OpenInWord()

    End Sub

    Private Sub picNonMSExtractDataI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox13.Click

        Impl_NonMSExtractDataI(PathTb.Text)

    End Sub

    Private Sub picNonMSExtractDataII_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox14.Click

        Impl_NonMSExtractDataII(PathTb.Text)

    End Sub

    Private Sub picZipRepair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.Click

        Impl_ZipRepairTry(PathTb.Text)

    End Sub

    Private Sub picXMLRepair_Click(sender As Object, e As EventArgs) Handles PictureBox18.Click

        Impl_XMLRepair(PathTb.Text)

    End Sub

    Private Sub pickPowerPointViewer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click

        Impl_PowerPointViewer()

    End Sub

    Private Sub picOpenAndRepair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click

        Impl_OpenAndRepair()

    End Sub

    Private Sub picOpenAndExtractData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click

        Impl_OpenAndExtractData()

    End Sub

    Private Sub picRestoreFromAutoBackup_Click(sender As System.Object, e As System.EventArgs) Handles picRestoreFromAutoBackup.Click

        Impl_RestoreFromBackup()

    End Sub

    Private Sub picOpenRecoveryToolboxforWordURL_Click(sender As System.Object, e As System.EventArgs) Handles picOpenRecoveryToolboxforWordURL.Click

        Impl_Recovery_Toolbox()

    End Sub

    Private Sub picOpenKernelWordRecoveryURL_Click(sender As System.Object, e As System.EventArgs) Handles picOpenKernelWordRecoveryURL.Click

        Impl_Kernel_Excel_recovery()

    End Sub

    Private Sub picExcelFix_Click(sender As System.Object, e As System.EventArgs) Handles PictureBox12.Click

        Impl_ExcelFix()

    End Sub

    Private Sub picOpenOnlineWordRepairUrl_Click(sender As System.Object, e As System.EventArgs) Handles picOpenOnlineWordRepairUrl.Click

        Impl_OpenOnlineWordRepairUrl()

    End Sub

    Private Sub picOpenPayPalDonateUrl_Click(sender As System.Object, e As System.EventArgs) Handles picOpenPayPalDonateUrl.Click

        Impl_OpenPayPalDonateUrl()

    End Sub

#Region "Cool Move Mouse Handler for External References"

    Private allowCoolMove As Boolean = False

    Private dx, dy As Integer

    'I used this two integers as I could use the function new POint due to the Import of the Excel

    Private Sub FormMain_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown

        allowCoolMove = True
        dx = Cursor.Position.X - Me.Location.X '// get coordinates.
        dy = Cursor.Position.Y - Me.Location.Y '// get coordinates.

    End Sub

    Private Sub FormMain_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove

        If allowCoolMove = True Then

            Me.Location = New System.Drawing.Point(Cursor.Position.X - dx, Cursor.Position.Y - dy) '// set coordinates.

        End If

    End Sub

    Private Sub FormMain_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp

        allowCoolMove = False

    End Sub

#End Region

#End Region

#Region "Implementation"

    Private Sub Impl_FolderFileChooserPreviousVersionSearch()

        Using objProcess As Process = New Process

            Using myProcess As Process = New Process

                Using fileCompare As Process = New Process

                    Dim OFD As New OpenFileDialog
                    Dim matches As MatchCollection
                    Dim objProcessReader As StreamReader = Nothing
                    Dim objProcessErrorReader As StreamReader = Nothing
                    Dim myStreamWriter As StreamWriter = Nothing
                    Dim fileCompareWriter As StreamWriter = Nothing
                    Dim fileCompareReader As StreamReader = Nothing
                    Dim fileCompareBoolean As Boolean
                    Dim fileCompBooleanError As Boolean
                    Dim objProcessOut As String
                    Dim driveLetter As String
                    Dim sFile As String
                    Dim fileCompareOut As String
                    Dim previoussFileShadowPath As String
                    Dim prevsFileShadowPathDate As String
                    Dim sFileShadowSize As Long
                    Dim prevsFileShadowPathSize As Long

                    Try

                        MsgBox("Note, some of the buttons won't work if there are visible " _
                               & "or invisible Excel instances running in task manager. Before " _
                               & "hitting the OK button on this message, please be sure to hit " _
                               & "Ctl-Alt-Delete, start Task Manager and end all instances of " _
                               & """EXCEL.EXE"" or ""EXCEL.EXE * 32"".", MsgBoxStyle.Exclamation)

                        With OFD

                            .ShowDialog()
                            filename = .FileName
                            PathTb.Text = .FileName

                        End With

                        sFile = PathTb.Text
                        shadowLinkFolderName.Clear()
                        nonErrorShadowPathList.Clear()
                        ComboBox1.Items.Clear()

                        MsgBox("Please wait while Excel Recovery searches for previous " _
                               & "versions of your file.", MsgBoxStyle.Information)

                        'Find out the number of vss shadow snapshots (restore 
                        'points). All shadows apparently have a linkable path 
                        '\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy#,
                        'where # is a simple one or two or three digit integer.

                        objProcess.StartInfo.UseShellExecute = False
                        objProcess.StartInfo.CreateNoWindow = True
                        objProcess.StartInfo.RedirectStandardOutput = True
                        objProcess.StartInfo.RedirectStandardError = True
                        objProcess.StartInfo.FileName() = "vssadmin"
                        objProcess.StartInfo.Arguments() = "List Shadows"
                        objProcess.Start()

                        objProcessReader = objProcess.StandardOutput
                        objProcessOut = objProcessReader.ReadToEnd
                        'MsgBox(objProcessOut)
                        'objProcessErrorReader = objProcess.StandardError
                        'objProcessError = objProcessErrorReader.ReadToEnd
                        objProcess.WaitForExit()
                        objProcess.Close()

                        ' Call Regex.Matches method.

                        driveLetter = sFile.Substring(0, 2)

                        matches = Regex.Matches(objProcessOut, _
                        "\\\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy[0-9]+")
                        counterVariable = 0
                        matchCount = matches.Count

                        If matchCount = 0 Then

                            MsgBox("There are no saved Restore Points for your machine, so no previous versions " _
                                   & "of your file are avilable. To enable previous versions of your file, you " _
                                   & "must turn on System Protection for the drive your file is stored on. You " _
                                   & "turn on System Protection in the System App of the Control Panel.")

                            Exit Sub

                        Else

                            ' Loop over matches.

                            For Each m As Match In matches

                                'MsgBox(m.ToString)

                                shadowLinkFolderName.Add(driveLetter & "\" & DelFromLeft( _
                                    "\\?\GLOBALROOT\Device\HarddiskVolume", (m.ToString())))
                                sFileShadowPath = (shadowLinkFolderName(counterVariable) & DelFromLeft( _
                                    driveLetter, sFile))

                                'Here I create temporary folders off the C: 
                                'drive which are mapped to each snapshot.

                                myProcess.StartInfo.FileName = "cmd.exe"
                                myProcess.StartInfo.UseShellExecute = False
                                myProcess.StartInfo.RedirectStandardInput = True
                                myProcess.StartInfo.RedirectStandardOutput = True
                                myProcess.StartInfo.CreateNoWindow = True
                                myProcess.Start()
                                myStreamWriter = myProcess.StandardInput
                                myStreamWriter.WriteLine("mklink /d " & _
                                (shadowLinkFolderName(counterVariable).ToString) _
                                & " " & (m.ToString()) & "\")
                                myStreamWriter.Flush()
                                myStreamWriter.Close()
                                myProcess.WaitForExit()
                                myProcess.Close()

                                Dim sFileShadowPathInfo As New FileInfo(sFileShadowPath)

                                'MsgBox(sFileShadowPath)
                                'MsgBox(sFileShadowPathInfo.Exists.ToString)

                                'Here I check if the file exists in the filing system of 
                                'the shadow image to which I've just created a shortcut. 
                                'If it does not, I delete the just created shortcut.

                                If sFileShadowPathInfo.Exists = False Then

                                    myProcess.StartInfo.FileName = "cmd.exe"
                                    myProcess.StartInfo.UseShellExecute = False
                                    myProcess.StartInfo.RedirectStandardInput = True
                                    myProcess.StartInfo.RedirectStandardOutput = True
                                    myProcess.StartInfo.CreateNoWindow = True
                                    myProcess.Start()
                                    myStreamWriter = myProcess.StandardInput
                                    myStreamWriter.WriteLine("rmdir " & shadowLinkFolderName(counterVariable))
                                    myStreamWriter.Flush()
                                    myStreamWriter.Close()
                                    myStreamWriter = Nothing
                                    myProcess.WaitForExit()
                                    myProcess.Close()

                                    counterVariable = counterVariable + 1

                                    Continue For

                                Else

                                    'Here I compare our recovery target file against the shadow 
                                    'copies. One shadow file copy is compared for each iteration 
                                    'of the loop. If the string "no difference encountered is found" 
                                    'then I know this shadow copy of the file is not worth looking 
                                    'at, as it is the same as the recovery target. Addditonally if 
                                    'the file compare error returns "FC: cannot open", then I end the
                                    'match iteration of the loop to and go to the next one.

                                    fileCompare.StartInfo.FileName = "cmd.exe"
                                    fileCompare.StartInfo.UseShellExecute = False
                                    fileCompare.StartInfo.RedirectStandardInput = True
                                    fileCompare.StartInfo.RedirectStandardOutput = True
                                    fileCompare.StartInfo.CreateNoWindow = True
                                    fileCompare.Start()

                                    fileCompareWriter = fileCompare.StandardInput
                                    fileCompareWriter.WriteLine("fc """ & sFile & """ """ _
                                                    & sFileShadowPath & """")
                                    fileCompareWriter.Flush()
                                    fileCompareWriter.Close()
                                    fileCompareReader = fileCompare.StandardOutput
                                    fileCompareOut = fileCompareReader.ReadToEnd
                                    fileCompareReader.Close()
                                    fileCompare.WaitForExit()
                                    fileCompare.Close()

                                    fileCompareBoolean = fileCompareOut.Contains("no differences encountered").ToString
                                    fileCompBooleanError = fileCompareOut.Contains("FC: cannot open").ToString

                                    'MsgBox(fileCompareBoolean)
                                    'MsgBox(fileCompBooleanError)

                                    If fileCompBooleanError = "True" Then

                                        counterVariable = counterVariable + 1

                                        Continue For

                                    End If

                                    If fileCompareBoolean = "True" Then

                                        counterVariable = counterVariable + 1

                                        Continue For

                                    End If

                                    'Here I take a positive result of a file difference between
                                    'the target and the shadow copy, and I write it out to a combo 
                                    'box on the form, so it can be chosen. I also only keep the 
                                    'first instance of a different shadow file as the others are 
                                    'identical. I distinguish if they are the same by date and size.

                                    sFileShadowPathDate = sFileShadowPathInfo.LastWriteTime
                                    sFileShadowSize = sFileShadowPathInfo.Length
                                    sFileShadowName = sFileShadowPathInfo.Name

                                    If ComboBox1.Items.Count = 0 Then

                                        ComboBox1.Items.Add("File Name: " _
                                        & sFileShadowName & " Last Modified: " & sFileShadowPathDate _
                                        & " Size in Bytes: " & sFileShadowSize)
                                        nonErrorShadowPathList.Add(sFileShadowPath)
                                        preVersionHashTable.Add(sFileShadowSize, sFileShadowPath)

                                        counterVariable = counterVariable + 1

                                        Continue For

                                    End If

                                    previousVersionCounterVariable = ComboBox1.Items.Count - 1
                                    previoussFileShadowPath = nonErrorShadowPathList(previousVersionCounterVariable)

                                    Dim prevsFileShadowPathInfo As New FileInfo(previoussFileShadowPath)

                                    prevsFileShadowPathSize = prevsFileShadowPathInfo.Length
                                    prevsFileShadowPathDate = prevsFileShadowPathInfo.LastWriteTime

                                    If String.Equals(sFileShadowPathDate, prevsFileShadowPathDate) _
                                        And Long.Equals(sFileShadowSize, prevsFileShadowPathSize) Then

                                        counterVariable = counterVariable + 1

                                        Continue For

                                    Else

                                        ComboBox1.Items.Add("File Name: " _
                                        & sFileShadowName & " Last Modified: " & sFileShadowPathDate _
                                        & " Size in Bytes: " & sFileShadowSize)
                                        nonErrorShadowPathList.Add(sFileShadowPath)
                                        preVersionHashTable.Add(sFileShadowSize, sFileShadowPath)
                                        counterVariable = counterVariable + 1

                                        Continue For

                                    End If

                                End If

                            Next m

                            MsgBox("Processing has finished and should have returned previous " _
                                   & "versions, if they exist.", MsgBoxStyle.Information)

                        End If

                    Catch ex As Exception

                        MessageBox.Show(ex.Message)

                    End Try

                End Using

            End Using

        End Using

    End Sub

    Private Sub Impl_PreviousVersionFileChoice()

        Try

            'Extract the size of the selected previous version file.

            comboBoxChoiceIndex = ComboBox1.SelectedIndex
            pathToComboBoxSelectedFile = nonErrorShadowPathList(comboBoxChoiceIndex)

            MsgBox("ComboBox1.selectedindex: " & comboBoxChoiceIndex)
            MsgBox(pathToComboBoxSelectedFile)


            'Use hash table set up in the Sub Impl_FolderFileChooserPreviousVersionSearch()
            'and retrieve the full path of the recovered file by presenting its date as key.

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub Impl_SaveFileFolder()

        Dim saveFileDialog1 As New SaveFileDialog()

        Try

            'Save the file recovered. 

            saveFileDialog1.Filter = "Excel 97-2003 Format Files (*.xls)|*.xls|" _
                & "Excel 2007 - 2013 Format Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
            saveFileDialog1.AddExtension = True
            saveFileDialog1.InitialDirectory = sFileShadowPath
            saveFileDialog1.DefaultExt = "xls"
            saveFileDialog1.FilterIndex = 1
            saveFileDialog1.RestoreDirectory = True
            saveFileDialog1.SupportMultiDottedExtensions = True
            saveFileDialog1.AutoUpgradeEnabled = True
            saveFileDialog1.OverwritePrompt = True

            If saveFileDialog1.ShowDialog() = DialogResult.OK Then

                Dim pathToComboBoxSelectedFileInfo As New FileInfo(pathToComboBoxSelectedFile)

                MsgBox(saveFileDialog1.FileName)
                saveShadowPath = saveFileDialog1.FileName
                pathToComboBoxSelectedFileSize = pathToComboBoxSelectedFileInfo.Length / 1024
                MsgBox(pathToComboBoxSelectedFileSize)

                'Write to message box of successful saving and location of recovered version.

                If System.IO.File.Exists(pathToComboBoxSelectedFile) = True Then

                    System.IO.File.Copy(pathToComboBoxSelectedFile, saveShadowPath, True)

                    MsgBox(("The Previous version of " & pathToComboBoxSelectedFileInfo.Name _
                                  & " last modified on " _
                                  & pathToComboBoxSelectedFileInfo.LastWriteTime _
                                  & ". The file is " & pathToComboBoxSelectedFileSize _
                                  & " KB in size" & " was saved to a new location: " & _
                                  saveShadowPath) & ".", MsgBoxStyle.Information)

                Else

                    MsgBox("Can't connect to previous version file.", MsgBoxStyle.Exclamation)

                End If

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub Impl_SaveAsSYLK()

        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing
        Dim oWSheet As Worksheet = Nothing
        Dim sFile As String = PathTb.Text
        Dim sFileName As String = Nothing
        Dim sDirName As String = Nothing
        Dim sFileSylkName As String = Nothing
        Dim TargetKey As RegistryKey

        Try

            sFile = PathTb.Text
            sFileName = Path.GetFileNameWithoutExtension(sFile)
            sDirName = Path.GetDirectoryName(sFile)
            sFileSylkName = sDirName & "\" & sFileName & ".slk"

            TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

            'Look to see if Excel is installed.

            If TargetKey Is Nothing Then

                MsgBox("The save to SYLK format method requires Excel," _
                       & " however it does not appear to be installed.", MsgBoxStyle.Exclamation)

                Exit Sub

            Else

                'Key is found

                TargetKey.Close()

                sFile = PathTb.Text
                sFileName = Path.GetFileNameWithoutExtension(sFile)
                sDirName = Path.GetDirectoryName(sFile)
                sFileSylkName = sDirName & "\" & sFileName & ".slk"

                'Start Excel and open the workbook. Then save to SYLK format.

                oExcel.Visible = False
                oBooks = oExcel.Workbooks
                oBook = oBooks.Open(Filename:=sFile)
                oWSheet = oBook.ActiveSheet()
                oWSheet.SaveAs(Filename:=sFileSylkName, FileFormat:=Excel.XlFileFormat.xlSYLK)

                'Open SYLK formatted file.

                If File.Exists(sFileSylkName) Then

                    oBook.Close()
                    oExcel.Visible = True
                    oBook = oBooks.Open(Filename:=sFileSylkName)

                Else

                    MessageBox.Show("Failed to create " & sFileSylkName)

                End If

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            oWSheet.Delete()
            oBook.Close()
            oBooks.Close()

            ' Excel will stick around after Quit if it is not under user 
            ' control and there are outstanding references. When Excel is 
            ' started or attached programmatically and 
            ' Application.Visible = false, Application.UserControl is false. 
            ' The UserControl property can be explicitly set to True which 
            ' should force the application to terminate when Quit is called, 
            ' regardless of outstanding references.

            oExcel.UserControl = True
            oExcel.Quit()

            If Not oWSheet Is Nothing Then

                Marshal.FinalReleaseComObject(oWSheet)
                oWSheet = Nothing

            End If

            If Not oBook Is Nothing Then

                Marshal.FinalReleaseComObject(oBook)
                oBook = Nothing

            End If

            If Not oBooks Is Nothing Then

                Marshal.FinalReleaseComObject(oBooks)
                oBooks = Nothing

            End If

            If Not oExcel Is Nothing Then

                Marshal.FinalReleaseComObject(oExcel)
                oExcel = Nothing

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Sub

    Private Sub Impl_SaveAsHTML()

        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing
        Dim oWSheet As Worksheet = Nothing
        Dim sFile As String
        Dim sFileName As String
        Dim sDirName As String
        Dim sFileHTMLName As String
        Dim TargetKey As RegistryKey

        Try

            sFile = PathTb.Text
            sFileName = Path.GetFileNameWithoutExtension(sFile)
            sDirName = Path.GetDirectoryName(sFile)
            sFileHTMLName = sDirName & "\" & sFileName & ".html"
            TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

            'Look to see if Excel is installed.

            If TargetKey Is Nothing Then

                MsgBox("The save to SYLK format method requires Excel, however it " _
                       & "does not appear to be installed.", MsgBoxStyle.Exclamation)

                Exit Sub

            Else

                'Key is found

                TargetKey.Close()

                'Start Excel and open the workbook. Then save to HTML format.

                oExcel.Visible = False
                oBooks = oExcel.Workbooks
                oBook = oBooks.Open(Filename:=sFile)
                oWSheet = oBook.ActiveSheet()
                oWSheet.SaveAs(Filename:=sFileHTMLName, FileFormat:=Excel.XlFileFormat.xlHtml)

                'Open HTML formatted file.

                If File.Exists(sFileHTMLName) Then

                    oBook.Close()
                    oExcel.Visible = True
                    oBook = oBooks.Open(Filename:=sFileHTMLName)

                Else

                    MessageBox.Show("Failed to create " & sFileHTMLName)

                End If

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            oWSheet.Delete()
            oBook.Close()
            oBooks.Close()

            ' Excel will stick around after Quit if it is not under user 
            ' control and there are outstanding references. When Excel is 
            ' started or attached programmatically and 
            ' Application.Visible = false, Application.UserControl is false. 
            ' The UserControl property can be explicitly set to True which 
            ' should force the application to terminate when Quit is called, 
            ' regardless of outstanding references.

            oExcel.UserControl = True
            oExcel.Quit()

            If Not oWSheet Is Nothing Then

                Marshal.FinalReleaseComObject(oWSheet)
                oWSheet = Nothing

            End If

            If Not oBook Is Nothing Then

                Marshal.FinalReleaseComObject(oBook)
                oBook = Nothing

            End If

            If Not oBooks Is Nothing Then

                Marshal.FinalReleaseComObject(oBooks)
                oBooks = Nothing

            End If

            If Not oExcel Is Nothing Then

                Marshal.FinalReleaseComObject(oExcel)
                oExcel = Nothing

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Sub

    Private Sub Impl_ManualCalculations()

        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing
        Dim sFile As String
        Dim TargetKey As RegistryKey

        Try

            sFile = PathTb.Text
            TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

            'Make sure Excel is installed.

            If TargetKey Is Nothing Then

                MsgBox("Opening the Excel file with calculations set to manual method " _
                       & "requires Excel, however it does not appear to be installed.", MsgBoxStyle.Exclamation)

                Exit Sub

            Else

                'Excel Key is found.

                TargetKey.Close()

                oExcel.Visible = True

                'Open chosen target Excel file with calculations set to manual.

                oBooks = oExcel.Workbooks
                oBook = oBooks.Open(sFile)
                oExcel.Calculation = Excel.XlCalculation.xlCalculationManual

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            oBook.Close()
            oBooks.Close()

            ' Excel will stick around after Quit if it is not under user 
            ' control and there are outstanding references. When Excel is 
            ' started or attached programmatically and 
            ' Application.Visible = false, Application.UserControl is false. 
            ' The UserControl property can be explicitly set to True which 
            ' should force the application to terminate when Quit is called, 
            ' regardless of outstanding references.

            oExcel.UserControl = True
            oExcel.Quit()

            If Not oBook Is Nothing Then

                Marshal.FinalReleaseComObject(oBook)
                oBook = Nothing

            End If

            If Not oBooks Is Nothing Then

                Marshal.FinalReleaseComObject(oBooks)
                oBooks = Nothing

            End If

            If Not oExcel Is Nothing Then

                Marshal.FinalReleaseComObject(oExcel)
                oExcel = Nothing

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Sub

    Private Sub Impl_ExternalReferenceRecoveryMethod()

        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing
        Dim oWSheet As Worksheet = Nothing
        Dim rRange As Range = Nothing
        Dim sFile As String
        Dim TargetKey As RegistryKey

        Try

            sFile = PathTb.Text
            TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

            'Make sure Excel is installed.

            If TargetKey Is Nothing Then

                MsgBox("The External References method requires Excel, however it " _
                       & "does not appear to be installed.", MsgBoxStyle.Exclamation)

                Exit Sub

            Else

                'Excel key is found, so Excel is installed.

                TargetKey.Close()

                'Reset calculations to automatic.

                oBook = oExcel.Workbooks.Add
                oExcel.Visible = True
                oBook.Activate()
                oExcel.Calculation = Excel.XlCalculation.xlCalculationAutomatic

                'In cell A1 put in the External References formula.
                'Hopefully selection box appears allowing selection
                'of target worksheet for recovery.

                oExcel.Range("A1").Value = "=" & "'" & sFile & "'" & "!A1"

                'Msgbox is displayed with capability of selecting range 
                'from corrupt file to be displayed in new healthy sheet.

                oExcel.DisplayAlerts = False
                rRange = oExcel.InputBox(Prompt:= _
                    "Please select a range similar in size to your corrupt data " _
                    & "that you wish to recover.", Title:="SPECIFY RANGE", Type:=8)
                oExcel.DisplayAlerts = True

                If rRange Is Nothing Then

                    Exit Sub

                Else

                    oExcel.Range("A1").Copy()
                    oBook.ActiveSheet.Paste(rRange)

                End If

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            rRange.Delete()
            oWSheet.Delete()
            oBook.Close()
            oBooks.Close()

            ' Excel will stick around after Quit if it is not under user 
            ' control and there are outstanding references. When Excel is 
            ' started or attached programmatically and 
            ' Application.Visible = false, Application.UserControl is false. 
            ' The UserControl property can be explicitly set to True which 
            ' should force the application to terminate when Quit is called, 
            ' regardless of outstanding references.

            oExcel.UserControl = True
            oExcel.Quit()

            If Not rRange Is Nothing Then

                Marshal.FinalReleaseComObject(rRange)
                rRange = Nothing

            End If

            If Not oWSheet Is Nothing Then

                Marshal.FinalReleaseComObject(oWSheet)
                oWSheet = Nothing

            End If

            If Not oBook Is Nothing Then

                Marshal.FinalReleaseComObject(oBook)
                oBook = Nothing

            End If

            If Not oBooks Is Nothing Then

                Marshal.FinalReleaseComObject(oBooks)
                oBooks = Nothing

            End If

            If Not oExcel Is Nothing Then

                Marshal.FinalReleaseComObject(oExcel)
                oExcel = Nothing

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Sub

    Private Sub Impl_OpenInSafeMode()

        'Get path of current version of Excel.

        Dim excelPath As String
        Dim sFile As String
        Dim excelPathAndExcel As String = Nothing

        Using cmdLaunchExcel As Process = New Process

            Try

                excelPath = Registry.GetValue( _
                    "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", _
                    "Path", "Key does not exist")

                'TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

                'Make sure Excel exists for this method.

                If excelPath = "Key does not exist" Then

                    MsgBox("The Open in Safe Mode method requires Excel, " _
                           & "however it does not appear to be installed.", MsgBoxStyle.Exclamation)

                    Exit Sub

                Else

                    'Key is found so Excel exists. Open selected file in safe mode.

                    sFile = PathTb.Text
                    excelPathAndExcel = excelPath & "excel.exe"
                    MsgBox("""" & excelPathAndExcel _
                        & """ /s " & """" & sFile & """")
                    cmdLaunchExcel.StartInfo.UseShellExecute = True
                    cmdLaunchExcel.StartInfo.CreateNoWindow = False
                    cmdLaunchExcel.StartInfo.FileName() = """" & excelPathAndExcel & """"
                    cmdLaunchExcel.StartInfo.Arguments() = "/s " & """" & sFile & """"
                    cmdLaunchExcel.Start()

                End If

            Catch ex As Exception

                MessageBox.Show(ex.Message)

            End Try

        End Using

    End Sub

    Private Sub Impl_ExtractDataFromChart()

        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing
        Dim oWSheet As Worksheet = Nothing
        Dim oChart As Chart = Nothing
        Dim sFile As String
        Dim TargetKey As RegistryKey
        Dim NumberOfRows As Integer
        Dim X As Object
        Dim Counter As Integer = 2

        Try

            sFile = PathTb.Text
            TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

            If TargetKey Is Nothing Then
                MsgBox("The Macro Graph Data Recovery method requires Excel, " _
                       & "however it does not appear to be installed.", MsgBoxStyle.Exclamation)
                Exit Sub

            Else

                TargetKey.Close()

                'key is found

                oExcel.Visible = True
                oBooks = oExcel.Workbooks
                oBook = oBooks.Open(Filename:=sFile)
                oWSheet = oBook.Worksheets.Add()
                oWSheet.Name = "ChartData"
                oWSheet.Activate()
                MsgBox("Select the chart you wish to extract data from.", MsgBoxStyle.Information)
                oChart = oBook.ActiveChart

                ' Calculate the number of rows of data. 

                NumberOfRows = UBound(oChart.SeriesCollection(1).Values)
                oWSheet.Cells(1, 1) = "X Values"

                ' Write x-axis values to worksheet. 

                With oWSheet
                    .Range(.Cells(2, 1), _
                    .Cells(NumberOfRows + 1, 1)).Value = _
                    oExcel.WorksheetFunction.Transpose(oChart.SeriesCollection(1).XValues)
                End With

                ' Loop through all series in the chart   
                ' and write their values to the worksheet. 

                For Each X In oChart.SeriesCollection

                    oWSheet.Cells(1, Counter) = X.Name

                    With oWSheet

                        .Range(.Cells(2, Counter), _
                        .Cells(NumberOfRows + 1, Counter)).Value = _
                        oExcel.WorksheetFunction.Transpose(X.Values)

                    End With

                    Counter = Counter + 1

                Next

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub Impl_OpenInWordPad()

        Dim sFile As String
        Dim regVersion As Microsoft.Win32.RegistryKey = Nothing
        Dim proc As New Process

        Try

            'Make sure WordPad exists, then open with Wordpad. 
            'It will show machine code only apparently.

            sFile = PathTb.Text
            regVersion = Microsoft.Win32.Registry.LocalMachine.OpenSubKey( _
                "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wordpad.exe", False)

            If regVersion IsNot Nothing Then

                sFile = PathTb.Text

                With proc.StartInfo

                    .FileName = regVersion.GetValue("").ToString
                    .Arguments = """" & sFile & """"

                End With

                proc.Start()

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub Impl_OpenInWord()

        Dim sFile As String
        Dim regVersion As Microsoft.Win32.RegistryKey = Nothing
        Dim proc As New Process

        Try

            'Make sure WordPad exists, then open with Wordpad. 
            'It will show machine code only apparently.

            sFile = PathTb.Text
            regVersion = Microsoft.Win32.Registry.LocalMachine.OpenSubKey( _
                "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe", False)

            If regVersion IsNot Nothing Then

                sFile = PathTb.Text

                With proc.StartInfo

                    .FileName = regVersion.GetValue("").ToString
                    .Arguments = """" & sFile & """"

                End With

                proc.Start()

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Public Sub Impl_NonMSExtractDataI(sFile As String)

        Dim formProgress As FormProgress = InitProgressForm(8)
        Dim extractCMD As New Process()
        Dim myPercent As Char
        Dim myZipCommand As String
        Dim strAddinPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
        Dim strTempFileName As String
        Dim strTempFilePath As String
        Dim sFileText As String
        Dim doctotextOutput As String
        Dim sErr As String
        Dim bAns As String
        Dim oExcel As New Excel.Application
        Dim oBooks As Workbooks = Nothing
        Dim oBook As Workbook = Nothing

        Try

            sFile = PathTb.Text
            myPercent = Chr(37)
            myZipCommand = """no-frills.exe " & myPercent _
                           & "a " & myPercent & "d " & myPercent & "f"""

            strTempFileName = "ExcelRecoveryAddin.tmp"
            strTempFilePath = Path.Combine(strAddinPath, strTempFileName)

            If File.Exists(strTempFilePath) Then

                File.Delete(strTempFilePath)

            End If

            If PerformProgressStep(formProgress) Then

                File.Copy(sFile, strTempFilePath)

                If PerformProgressStep(formProgress) Then

                    extractCMD.StartInfo.FileName = Path.Combine(strAddinPath, "doctotext.exe")
                    extractCMD.StartInfo.Arguments = "--fix-xml --unzip-cmd=" & myZipCommand & _
                        " """ & strTempFileName & """"
                    extractCMD.StartInfo.UseShellExecute = False
                    extractCMD.StartInfo.RedirectStandardOutput = True
                    extractCMD.StartInfo.CreateNoWindow = True
                    extractCMD.StartInfo.WorkingDirectory = strAddinPath
                    extractCMD.Start()

                    If PerformProgressStep(formProgress) Then

                        sFileText = sFile & ".txt"
                        doctotextOutput = extractCMD.StandardOutput.ReadToEnd()

                        If PerformProgressStep(formProgress) Then

                            sErr = ""

                            'Save to different file

                            bAns = SaveTextToFile(doctotextOutput, sFileText, sErr)

                            If bAns Then

                                If PerformProgressStep(formProgress) Then

                                    MsgBox("Please note: all successfully extracted worksheets " _
                                           & "will appear in just one text file opening in " _
                                           & "Excel. The data from each worksheet will appear " _
                                           & "underneath the previous one.", MsgBoxStyle.Information)

                                    oExcel.Visible = True

                                    If PerformProgressStep(formProgress) Then

                                        oBooks = oExcel.Workbooks
                                        oBook = oBooks.Open(sFileText)

                                        If oBook IsNot Nothing Then

                                            Try

                                                Me.WindowState = FormWindowState.Minimized

                                                oBook.Activate()

                                            Catch

                                            End Try

                                        End If

                                        PerformProgressStep(formProgress)

                                        If Not oBook Is Nothing Then

                                            Marshal.FinalReleaseComObject(oBook)
                                            oBook = Nothing

                                        End If

                                        If Not oBooks Is Nothing Then

                                            Marshal.FinalReleaseComObject(oBooks)
                                            oBooks = Nothing

                                        End If

                                        If Not oExcel Is Nothing Then

                                            Marshal.FinalReleaseComObject(oExcel)
                                            oExcel = Nothing

                                        End If

                                        GC.Collect()
                                        GC.WaitForPendingFinalizers()

                                        ' GC needs to be called twice in order to get the Finalizers called  
                                        ' - the first time in, it simply makes a list of what is to be  
                                        ' finalized, the second time in, it actually the finalizing. Only  
                                        ' then will the object do its automatic ReleaseComObject. 

                                        GC.Collect()
                                        GC.WaitForPendingFinalizers()

                                        PerformProgressStep(formProgress)

                                    Else

                                        MsgBox("Error extracting file: " & sErr, MsgBoxStyle.Exclamation)

                                    End If

                                    extractCMD.Close()

                                    File.Delete(strTempFilePath)

                                End If

                            End If

                        End If

                    End If

                End If

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            oBook.Close()
            oBooks.Close()

            ' Excel will stick around after Quit if it is not under user 
            ' control and there are outstanding references. When Excel is 
            ' started or attached programmatically and 
            ' Application.Visible = false, Application.UserControl is false. 
            ' The UserControl property can be explicitly set to True which 
            ' should force the application to terminate when Quit is called, 
            ' regardless of outstanding references.

            oExcel.UserControl = True
            oExcel.Quit()

            If Not oBook Is Nothing Then

                Marshal.FinalReleaseComObject(oBook)
                oBook = Nothing

            End If

            If Not oBooks Is Nothing Then

                Marshal.FinalReleaseComObject(oBooks)
                oBooks = Nothing

            End If

            If Not oExcel Is Nothing Then

                Marshal.FinalReleaseComObject(oExcel)
                oExcel = Nothing

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

        CleanupProgressForm(formProgress)

    End Sub

    Private Sub Impl_NonMSExtractDataII(sFile As String)

        Dim sFileExtension As String = Path.GetExtension(sFile)
        Dim OFlD As New FolderBrowserDialog
        Dim coffecCMD As New Process()
        Dim sFileInfo As New FileInfo(sFile)
        Dim sFileName As String
        Dim sFilePath As String

        Try

            'This extract uses a command line recovery app that only works 
            'with xlsx files. Will start up command line in background.

            sFile = PathTb.Text
            sFileExtension = Path.GetExtension(sFile)

            If sFileExtension = ".xls" Then

                MsgBox("This data extraction only works with xlsx files.", MsgBoxStyle.Exclamation)

                Exit Sub

            Else

                MsgBox("If successful, each Worksheet will be saved as separate CSV in " _
                   & "the same directory as your corrupt file.", MsgBoxStyle.Information)

                coffecCMD.StartInfo.FileName = "coffec.exe"
                coffecCMD.StartInfo.Arguments = "-t """ & sFile & """"
                coffecCMD.StartInfo.UseShellExecute = True
                coffecCMD.StartInfo.CreateNoWindow = True
                coffecCMD.Start()
                coffecCMD.WaitForExit()
                coffecCMD.Close()

                'Results are written to CSV file. 
                'Directory saved in is opened in Explorer.

                sFileName = sFileInfo.Name
                sFilePath = DelFromRight(sFileName, sFile)

                Process.Start("explorer.exe", sFilePath)

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Sub

    Public Sub Impl_ZipRepairTry(sFile As String)

        Dim formProgress As FormProgress = Nothing
        Dim repairZip As New Process()
        Dim sFileZip As String
        Dim sFileExtension As String
        Dim sFileName As String
        Dim zipRepairedsFileName As String
        Dim sFileBasePath As String
        Dim zipRepairedFullPathFileName As String
        Dim strFullPath As String
        Dim repairZipReader As StreamReader = Nothing
        Dim repairZipCompOut As String
        Dim zipRepairedFullPathXlsxName As String
        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing

        Try

            sFile = PathTb.Text
            sFileExtension = Path.GetExtension(sFile)

            If sFileExtension = ".xls" Then

                MsgBox("Zip repair is only useful for files with .xslx " _
                       & "extensions and format.", MsgBoxStyle.Exclamation)

                Exit Sub

            End If

            formProgress = InitProgressForm(6)
            sFileZip = sFile & ".zip"
            sFileName = Path.GetFileName(sFile)
            zipRepairedsFileName = "zipRepaired" & sFileName & ".zip"
            sFileBasePath = DelFromRight(sFileName, sFile)
            zipRepairedFullPathFileName = sFileBasePath & zipRepairedsFileName

            If File.Exists(sFileZip) Then

                File.Delete(sFileZip)

            End If

            FileCopy(sFile, sFileZip)

            If PerformProgressStep(formProgress) Then

                strFullPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), _
                                           "zip.exe")
                repairZip.StartInfo.FileName = strFullPath
                repairZip.StartInfo.Arguments = "-FF """ & sFileZip & """ --out " _
                    & Chr(34) & zipRepairedFullPathFileName & Chr(34)
                repairZip.StartInfo.UseShellExecute = False
                repairZip.StartInfo.RedirectStandardOutput = True
                repairZip.StartInfo.CreateNoWindow = True
                repairZip.Start()

                If PerformProgressStep(formProgress) Then

                    repairZipReader = repairZip.StandardOutput
                    repairZipCompOut = repairZipReader.ReadToEnd

                    If PerformProgressStep(formProgress) Then

                        repairZipReader.Close()
                        repairZip.WaitForExit()
                        repairZip.Close()

                        If PerformProgressStep(formProgress) Then

                            zipRepairedFullPathXlsxName = DelFromRight(".zip", zipRepairedFullPathFileName)
                            oBooks = Nothing
                            oBook = Nothing

                            If File.Exists(zipRepairedFullPathXlsxName) Then

                                File.Delete(zipRepairedFullPathXlsxName)

                            End If

                            Rename(zipRepairedFullPathFileName, zipRepairedFullPathXlsxName)
                            oExcel.Visible = True

                            If PerformProgressStep(formProgress) Then

                                oBooks = oExcel.Workbooks
                                oBook = oBooks.Open(Filename:=zipRepairedFullPathXlsxName)

                                Me.WindowState = FormWindowState.Minimized

                                Try

                                    oBook.Activate()

                                Catch

                                End Try

                                PerformProgressStep(formProgress)

                            End If

                        End If

                    End If

                End If

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            oBook.Close()
            oBooks.Close()

            ' Excel will stick around after Quit if it is not under user 
            ' control and there are outstanding references. When Excel is 
            ' started or attached programmatically and 
            ' Application.Visible = false, Application.UserControl is false. 
            ' The UserControl property can be explicitly set to True which 
            ' should force the application to terminate when Quit is called, 
            ' regardless of outstanding references.

            oExcel.UserControl = True
            oExcel.Quit()

            If Not oBook Is Nothing Then

                Marshal.FinalReleaseComObject(oBook)
                oBook = Nothing

            End If

            If Not oBooks Is Nothing Then

                Marshal.FinalReleaseComObject(oBooks)
                oBooks = Nothing

            End If

            If Not oExcel Is Nothing Then

                Marshal.FinalReleaseComObject(oExcel)
                oExcel = Nothing

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

        CleanupProgressForm(formProgress)

    End Sub

    Private Sub Impl_XMLRepair(sFile As String)

        Dim repairZip As New Process()
        Dim sFileZip As String
        Dim sFileExtension As String
        Dim sFileName As String
        Dim zipRepairedsFileName As String
        Dim sFileBasePath As String
        Dim zipRepairedFullPathFileName As String
        Dim strFullPath As String
        Dim sevenZipFullPath As String
        Dim repairZipReader As StreamReader = Nothing
        Dim zipRepairedFullPathXlsx As String
        Dim zipRepairedFullPathXlsxName As String
        Dim extractedRepairedZipOutputDirectory As String
        Dim extractedZipDirectorySpacesRemoved As String
        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing

        Try

            'First we repair the zip as with the previous button.

            sFile = PathTb.Text
            sFileExtension = Path.GetExtension(sFile)

            If sFileExtension = ".xls" Then

                MsgBox("Zip repair is only useful for files with .xslx " _
                       & "extensions and format.", MsgBoxStyle.Exclamation)

                Exit Sub

            End If

            sFileZip = sFile & ".zip"
            sFileName = Path.GetFileName(sFile)
            zipRepairedsFileName = "zipRepaired_" & sFileName & ".zip"
            sFileBasePath = DelFromRight(sFileName, sFile)
            zipRepairedFullPathFileName = sFileBasePath & zipRepairedsFileName

            If File.Exists(sFileZip) Then

                File.Delete(sFileZip)

            End If

            FileCopy(sFile, sFileZip)

            strFullPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), _
                                       "zip.exe")
            MsgBox("The repair zip full command is: """ & strFullPath & """ -FF """ & sFileZip & """ --out """ _
                & zipRepairedFullPathFileName & """")
            repairZip.StartInfo.FileName = """" & strFullPath & """"
            repairZip.StartInfo.Arguments = "-FF """ & sFileZip & """ --out """ _
                & zipRepairedFullPathFileName & """"
            repairZip.StartInfo.UseShellExecute = False
            repairZip.StartInfo.RedirectStandardOutput = True
            repairZip.StartInfo.CreateNoWindow = True
            repairZip.Start()
            repairZip.WaitForExit()
            repairZip.Close()

            MsgBox("Check for zipRepairedFullPathFileName: " & zipRepairedFullPathFileName)

            'Now we extract the repaired file.

            zipRepairedFullPathXlsx = DelFromRight(".zip", _
                                        zipRepairedFullPathFileName)

            Dim zipRepairedFullPathXlsxNameInfo As New FileInfo(zipRepairedFullPathXlsx)

            zipRepairedFullPathXlsxName = zipRepairedFullPathXlsxNameInfo.Name
            extractedRepairedZipOutputDirectory = DelFromRight(".xlsx", _
                                        zipRepairedFullPathXlsxName)
            extractedZipDirectorySpacesRemoved = extractedRepairedZipOutputDirectory.Replace(" ", "_")

            sevenZipFullPath = _
                Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), _
                "7z.exe")
            MsgBox("""" & sevenZipFullPath & """" & " x """ & zipRepairedFullPathFileName _
                   & """ -o" & extractedZipDirectorySpacesRemoved)
            repairZip.StartInfo.FileName = """" & sevenZipFullPath & """"
            repairZip.StartInfo.Arguments = "x """ & zipRepairedFullPathFileName & """ -o" _
                & extractedZipDirectorySpacesRemoved
            repairZip.StartInfo.UseShellExecute = False
            repairZip.StartInfo.CreateNoWindow = True
            repairZip.Start()
            repairZip.WaitForExit()
            repairZip.Close()










            oBooks = Nothing
            oBook = Nothing

            If File.Exists(zipRepairedFullPathXlsx) Then

                File.Delete(zipRepairedFullPathXlsx)

            End If

            File.Copy(zipRepairedFullPathFileName, zipRepairedFullPathXlsx)
            oExcel.Visible = True
            oBooks = oExcel.Workbooks
            oBook = oBooks.Open(Filename:=zipRepairedFullPathXlsx)

            Me.WindowState = FormWindowState.Minimized

            Try

                oBook.Activate()

            Catch

            End Try

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            oBook.Close()
            oBooks.Close()

            ' Excel will stick around after Quit if it is not under user 
            ' control and there are outstanding references. When Excel is 
            ' started or attached programmatically and 
            ' Application.Visible = false, Application.UserControl is false. 
            ' The UserControl property can be explicitly set to True which 
            ' should force the application to terminate when Quit is called, 
            ' regardless of outstanding references.

            oExcel.UserControl = True
            oExcel.Quit()

            If Not oBook Is Nothing Then

                Marshal.FinalReleaseComObject(oBook)
                oBook = Nothing

            End If

            If Not oBooks Is Nothing Then

                Marshal.FinalReleaseComObject(oBooks)
                oBooks = Nothing

            End If

            If Not oExcel Is Nothing Then

                Marshal.FinalReleaseComObject(oExcel)
                oExcel = Nothing

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

        CleanupProgressForm(formProgress)

    End Sub

    Private Sub Impl_PowerPointViewer()

        Dim sFile As String = PathTb.Text
        Dim TargetKey As RegistryKey
        Dim myExcelViewerPath As String
        Dim shellID As Integer

        Try

            TargetKey = Registry.LocalMachine.OpenSubKey( _
            "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\" _
            & "S-1-5-18\Components\744F90AE0017B424981389D9B3E9A1AB")

            'Make sure the most recent version of Excel viewer is installed.

            If TargetKey Is Nothing Then

                MsgBox("You have not downloaded the most recent Microsoft Excel Viewer. " _
                    & "Please install after downloading and try clicking this button again.", MsgBoxStyle.Exclamation)

                System.Diagnostics.Process.Start("http://www.microsoft.com/download/en/details.aspx?id=10")

            Else

                'From the path in the registry, launch Excel viewer with target file's path.

                myExcelViewerPath = My.Computer.Registry.GetValue(TargetKey.ToString, _
                                    "00002159F30090400000000000F01FEC", Nothing)
                shellID = Shell(myExcelViewerPath & " /s /r " & """" & sFile & """", AppWinStyle.NormalFocus)

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub Impl_OpenAndRepair()

        Dim TargetKey As RegistryKey
        Dim sFile As String
        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing

        Try

            TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

            'Make sure Excel exists.

            If TargetKey Is Nothing Then

                MsgBox("The Open and Repair method requires Excel, however it " _
                       & "does not appear to be installed.", MsgBoxStyle.Exclamation)

                Exit Sub

            Else

                'Key is found, so Excel exists.

                TargetKey.Close()
                sFile = PathTb.Text

                'Start Excel and open the workbook in Open and Repair mode.

                If Path.GetExtension(sFile) = ".xls" Then

                    oExcel.Visible = True
                    oBooks = oExcel.Workbooks
                    oBook = oBooks.Open(Filename:=sFile, CorruptLoad:=XlCorruptLoad.xlRepairFile)

                    MsgBox("Excel completed file level validation and repair. Some parts of this " _
                    & "workbook may have been repaired or discarded.", MsgBoxStyle.Information)

                Else

                    oExcel.Visible = True
                    oBooks = oExcel.Workbooks
                    oBook = oBooks.Open(Filename:=sFile, CorruptLoad:=XlCorruptLoad.xlRepairFile)

                End If

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            oBook.Close()
            oBooks.Close()

            ' Excel will stick around after Quit if it is not under user 
            ' control and there are outstanding references. When Excel is 
            ' started or attached programmatically and 
            ' Application.Visible = false, Application.UserControl is false. 
            ' The UserControl property can be explicitly set to True which 
            ' should force the application to terminate when Quit is called, 
            ' regardless of outstanding references.

            oExcel.UserControl = True
            oExcel.Quit()

            If Not oBook Is Nothing Then

                Marshal.FinalReleaseComObject(oBook)
                oBook = Nothing

            End If

            If Not oBooks Is Nothing Then

                Marshal.FinalReleaseComObject(oBooks)
                oBooks = Nothing

            End If

            If Not oExcel Is Nothing Then

                Marshal.FinalReleaseComObject(oExcel)
                oExcel = Nothing

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Sub

    Private Sub Impl_OpenAndExtractData()

        Dim TargetKey As RegistryKey
        Dim sFile As String
        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing

        Try

            TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

            'Make sure Excel is installed

            If TargetKey Is Nothing Then

                MsgBox("The Open and Extract Data method requires Excel, " _
                       & "however it does not appear to be installed.", MsgBoxStyle.Exclamation)

                Exit Sub

            Else

                'Key is found so Excel is installed.

                TargetKey.Close()

                sFile = PathTb.Text

                'Start Excel and open the workbook in open and extract data mode.

                oExcel.Visible = True
                oBooks = oExcel.Workbooks
                oBook = oBooks.Open(Filename:=sFile, CorruptLoad:=XlCorruptLoad.xlExtractData)

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            oBook.Close()
            oBooks.Close()

            ' Excel will stick around after Quit if it is not under user 
            ' control and there are outstanding references. When Excel is 
            ' started or attached programmatically and 
            ' Application.Visible = false, Application.UserControl is false. 
            ' The UserControl property can be explicitly set to True which 
            ' should force the application to terminate when Quit is called, 
            ' regardless of outstanding references.

            oExcel.UserControl = True
            oExcel.Quit()

            If Not oBook Is Nothing Then

                Marshal.FinalReleaseComObject(oBook)
                oBook = Nothing

            End If

            If Not oBooks Is Nothing Then
                oBooks.Close()
                Marshal.FinalReleaseComObject(oBooks)
                oBooks = Nothing

            End If

            If Not oExcel Is Nothing Then

                Marshal.FinalReleaseComObject(oExcel)
                oExcel = Nothing

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually the finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Sub

    Private Sub Impl_RestoreFromBackup()

        Try

            Me.Visible = False

            Dim formAutoSave As New FormAutoSave()

            formAutoSave.ShowDialog()

            Me.Visible = True

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub Impl_OpenPayPalDonateUrl()

        Dim proc As New Process()

        Try

            proc.StartInfo.FileName = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=3SJY6GVWUL65S"
            proc.StartInfo.Arguments = ""
            proc.StartInfo.UseShellExecute = True
            proc.Start()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub Impl_OpenOnlineWordRepairUrl()

        Dim proc As New Process()

        Try

            MsgBox("Until May 1, 2013, for a free $39 value file repair attempt, " _
       & "first go through Demo recovery with your corrupt file on the Online Office " _
       & "Recovery site that is about to open. After recovery, scroll down past ""Demo " _
       & "Results"" and enter in the coupon code ""S2SERVICES"" in the field above the " _
       & """Submit Code"" button at the end of the ""Full Results"" section. Use all " _
       & "caps for the code but don't include the quotes.", MsgBoxStyle.Information)

            proc.StartInfo.FileName = "https://online.officerecovery.com/"
            proc.StartInfo.Arguments = ""
            proc.StartInfo.UseShellExecute = True
            proc.Start()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Public Sub Impl_Recovery_Toolbox()

        Dim proc As New Process()

        Try

            proc.StartInfo.FileName = "http://www.plimus.com/jsp/redirect.jsp?contractId=2264308&referrer=socrtwo"
            proc.StartInfo.Arguments = ""
            proc.StartInfo.UseShellExecute = True
            proc.Start()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Public Sub Impl_Kernel_Excel_recovery()

        Dim proc As New Process()

        Try

            proc.StartInfo.FileName = "http://esd.element5.com/product.html?productid=300136778&stylefrom=300338460"
            proc.StartInfo.Arguments = ""
            proc.StartInfo.UseShellExecute = True
            proc.Start()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Public Sub Impl_ExcelFix()

        Dim proc As New Process()

        Try

            proc.StartInfo.FileName = "http://www.cimaware.com/info/info.php?lang=en&id=622&path=excelfix.html"
            proc.StartInfo.Arguments = ""
            proc.StartInfo.UseShellExecute = True
            proc.Start()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Public Sub CleanupProgressForm(ByVal formProgress As FormProgress)

        formProgress.Close()

    End Sub

    Private Sub TerminateProcess(ByVal procName As String, ByVal procId As Integer)

        Try

            For Each pr As Process In Process.GetProcesses()

                If (pr.ProcessName = procName.ToString().ToUpper() AndAlso pr.Id = procId) Then

                    pr.Kill()

                End If

            Next

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

        End Try

    End Sub

    Private Sub unusedChartDataRecoveryMacro(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Workbook = Nothing
        Dim oWSheet As Worksheet = Nothing
        Dim oChart As Chart = Nothing
        Dim sFile As String
        Dim TargetKey As RegistryKey
        Dim NumberOfRows As Integer
        Dim X As Object
        Dim Counter As Integer = 2

        Try

            sFile = PathTb.Text
            TargetKey = Registry.ClassesRoot.OpenSubKey("excel.application")

            If TargetKey Is Nothing Then
                MsgBox("The Macro Graph Data Recovery method requires Excel, " _
                       & "however it does not appear to be installed.", MsgBoxStyle.Exclamation)
                Exit Sub

            Else

                TargetKey.Close()

                'key is found

                oExcel.Visible = True
                oBooks = oExcel.Workbooks
                oBook = oBooks.Open(Filename:=sFile)
                oWSheet = oBook.Worksheets.Add()
                oWSheet.Name = "ChartData"
                oWSheet.Activate()
                MsgBox("Select the chart you wish to extract data from.", MsgBoxStyle.Information)
                oChart = oBook.ActiveChart

                ' Calculate the number of rows of data. 

                NumberOfRows = UBound(oChart.SeriesCollection(1).Values)
                oWSheet.Cells(1, 1) = "X Values"

                ' Write x-axis values to worksheet. 

                With oWSheet
                    .Range(.Cells(2, 1), _
                    .Cells(NumberOfRows + 1, 1)).Value = _
                    oExcel.WorksheetFunction.Transpose(oChart.SeriesCollection(1).XValues)
                End With

                ' Loop through all series in the chart   
                ' and write their values to the worksheet. 

                For Each X In oChart.SeriesCollection

                    oWSheet.Cells(1, Counter) = X.Name

                    With oWSheet

                        .Range(.Cells(2, Counter), _
                        .Cells(NumberOfRows + 1, Counter)).Value = _
                        oExcel.WorksheetFunction.Transpose(X.Values)

                    End With

                    Counter = Counter + 1

                Next

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

#End Region

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click

    End Sub
    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click

    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click

    End Sub

End Class
