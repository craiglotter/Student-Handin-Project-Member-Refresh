Imports System
Imports System.IO
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing
Imports System.Threading
Imports System.Windows.Forms


Public Class Main_Screen

    Private application_busy_exiting As Boolean = False
    Private busyworking As Boolean = False

    Private lastinputline As String = ""
    Private lastinputlineFull As String = ""
    Private inputlines As Long = 0
    Private statusmessage As String = ""
    Private statusresults As String = ""
    Private highestPercentageReached As Integer = 0
    Private inputlinesprecount As Long = 0
    Private datelaunched As Date = Now()
    Private pretestdone As Boolean = False
    
    Private primary_PercentComplete As Integer = 0

    Private secondary_inputlinesprecount As Long = 0
    Private secondary_inputlines As Long = 0
    Private secondary_inputlinesTotal As Long = 0
    Private secondary_lastinputline As String = ""
    Private secondary_highestPercentageReached As Integer = 0
    Private secondary_PercentComplete As Integer = 0
    Private secondary_lastinputlineFull As String = ""

    Private secondary_newinputlines As Long = 0
    Private secondary_newinputlinesTotal As Long = 0
    Private secondary_oldinputlines As Long = 0
    Private secondary_oldinputlinesTotal As Long = 0

    Private TimesExecuted As Integer = 0

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                If FullErrors_Checkbox.Checked = True Then
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString
                Else
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.Message.ToString
                End If
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Display_Message1.Dispose()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
                filewriter.Dispose()
                filewriter = Nothing
                ex = Nothing
                identifier_msg = Nothing
            End If
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub


    Private Sub Activity_Handler(ByVal Message As String)
        Try
            Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            dir = Nothing
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & Message)
            filewriter.Flush()
            filewriter.Close()
            filewriter.Dispose()
            filewriter = Nothing
            Message = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Activity_Logger")
        End Try
    End Sub

    Private Sub Status_Handler(ByVal Message As String)
        Try
            Status_Textbox.Text = Message.ToUpper
            Message = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Status_Handler")
        End Try
    End Sub


    Private Sub Status_Results_Handler(ByVal Message As String)
        Try
            If Message.Length > 0 Then
                If (Message.Length + StatusResults_RichtextBox.Text.Length) > StatusResults_RichtextBox.MaxLength Then
                    StatusResults_RichtextBox.Clear()
                End If
                StatusResults_RichtextBox.AppendText(Message & vbCrLf)
                StatusResults_RichtextBox.Focus()
                StatusResults_RichtextBox.Select(StatusResults_RichtextBox.Text.Length - 1, 0)
                StatusResults_RichtextBox.ScrollToCaret()
            End If
            Message = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Status_Results_Handler")
        End Try
    End Sub


    Private Function File_Exists(ByVal file_path As String) As Boolean
        Dim result As Boolean = False
        Try
            If Not file_path = "" And Not file_path Is Nothing Then
                Dim dinfo As FileInfo = New FileInfo(file_path)
                If dinfo.Exists = False Then
                    result = False
                Else
                    result = True
                End If
                dinfo = Nothing
            End If
            file_path = Nothing
        Catch ex As Exception
            Error_Handler(ex, "File_Exists")
        End Try
        Return result
    End Function

    Private Function Directory_Exists(ByVal directory_path As String) As Boolean
        Dim result As Boolean = False
        Try
            If Not directory_path = "" And Not directory_path Is Nothing Then
                Dim dinfo As DirectoryInfo = New DirectoryInfo(directory_path)
                If dinfo.Exists = False Then
                    result = False
                Else
                    result = True
                End If
                dinfo = Nothing
            End If
            directory_path = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Directory_Exists")
        End Try
        Return result
    End Function

    Private Sub main_screen_formclosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Timer1.Stop()
            Timer1.Dispose()
            NotifyIcon1.Dispose()

        Catch ex As Exception
            Error_Handler(ex, "main_screen_formclosing")
        End Try
    End Sub

    Private Sub main_screen_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DoubleClick
        Try
            hide_main_application()
        Catch ex As Exception
            Error_Handler(ex, "main_screen_dblclick")
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Label1.Text = My.Application.Info.Version.Major & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & My.Application.Info.Version.Revision
            TimesExecuted = 0
            'If Not My.Settings.InputTargetFolder_Textbox = Nothing Then
            InputTargetFolder_Textbox.Text = My.Settings.InputTargetFolder_Textbox
            InputTargetFolder_Textbox.Select(0, 0)
            'End If
            'If Not My.Settings.FullErrors_Checkbox = Nothing Then
            Select Case My.Settings.FullErrors_Checkbox
                Case True
                    FullErrors_Checkbox.Checked = True
                    Exit Select
                Case False
                    FullErrors_Checkbox.Checked = False
                    Exit Select
                Case Else
                    FullErrors_Checkbox.Checked = True
                    Exit Select
            End Select
            'End If
            'MsgBox(My.Settings.WindowsServer)
            ' If Not My.Settings.WindowsServer. = Nothing Then
            'MsgBox("here")
            Select Case My.Settings.WindowsServer
                Case True
                    'MsgBox("1")
                    RadioButton1.Checked = True
                    RadioButton2.Checked = False
                    Exit Select
                Case False
                    'MsgBox("2")
                    RadioButton1.Checked = False
                    RadioButton2.Checked = True
                    Exit Select
                Case Else
                    ' MsgBox("3")
                    RadioButton1.Checked = True
                    RadioButton2.Checked = False
                    Exit Select
            End Select
            ' End If
            'MsgBox(RadioButton1.Checked)

            inputseconds_label.Text = "86400"
            Timer1.Interval = 1000
            Timer1.Start()

            If InputTargetFolder_Textbox.Text.Length > 0 Then
                StartWorker()
            End If
            sender = Nothing
            e = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Main_Screen_Load")
        End Try
    End Sub

    Private Sub Main_Screen_Close(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            ' Cancel the asynchronous operation.
            Me.BackgroundWorker1.CancelAsync()
            ' Disable the Cancel button.
            cancelAsyncButton.Enabled = False
            BackgroundWorker1.Dispose()

            NotifyIcon1.Dispose()

            My.Settings.FullErrors_Checkbox = FullErrors_Checkbox.Checked
            My.Settings.InputTargetFolder_Textbox = InputTargetFolder_Textbox.Text
            My.Settings.WindowsServer = RadioButton1.Checked
            My.Settings.Save()


        Catch ex As Exception
            Error_Handler(ex, "Main_Screen_Close")
        End Try

    End Sub




    Private Sub FullErrors_Checkbox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FullErrors_Checkbox.CheckedChanged
        Status_Handler("Error Level Reporting Altered")
    End Sub

    Private Sub Browse_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Browse_Button.Click
        Status_Handler("Selecting Root Hand-in Folder")
        Try
            FolderBrowserDialog1.Description = "Select root hand-in folder:"
            If Directory_Exists(InputTargetFolder_Textbox.Text) = True Then
                FolderBrowserDialog1.SelectedPath = InputTargetFolder_Textbox.Text
            End If

            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog
            If result = Windows.Forms.DialogResult.OK Then
                InputTargetFolder_Textbox.Text = FolderBrowserDialog1.SelectedPath
            End If

        Catch ex As Exception
            Error_Handler(ex, "Browse_Button_Click")
        End Try
        Status_Handler("Root Hand-in Folder Selected")
    End Sub




    Private Sub startAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles startAsyncButton.Click
        If InputTargetFolder_Textbox.Text.Length > 0 Then
            StartWorker()
        End If
    End Sub


    Private Sub StartWorker()
        Try
            If busyworking = False Then

                StatusResults_RichtextBox.Clear()
                busyworking = True
                statusmessage = "Initializing Application for Operation Launch"
                Status_Handler(statusmessage)

                ' Reset the text in the result label.

                inputlines_label.Text = [String].Empty
                lastinputline_label.Text = [String].Empty
                datelaunched_label.Text = [String].Empty
                inputlines2_label.Text = [String].Empty
                inputlines2Total_label.Text = [String].Empty
                lastinputline2_label.Text = [String].Empty

                secondary_inputlinesprecount = 0
                secondary_inputlines = 0
                secondary_inputlinesTotal = 0
                secondary_newinputlines = 0
                secondary_newinputlinesTotal = 0
                secondary_oldinputlines = 0
                secondary_oldinputlinesTotal = 0

                secondary_lastinputline = ""
                secondary_lastinputlineFull = ""
                secondary_highestPercentageReached = 0
                secondary_PercentComplete = 0

                inputlines = 0
                lastinputline = ""
                lastinputlineFull = ""
                statusmessage = ""
                highestPercentageReached = 0
                inputlinesprecount = 0
                datelaunched = Now()
                pretestdone = False


                Controls_Enabler("run")


                ' Start the asynchronous operation.
                Me.LinkLabel1.Visible = True
                BackgroundWorker1.RunWorkerAsync(InputTargetFolder_Textbox.Text)
            End If
        Catch ex As Exception
            Error_Handler(ex, "StartWorker")
        End Try
    End Sub 'startAsyncButton_Click




    Private Sub cancelAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelAsyncButton.Click

        ' Cancel the asynchronous operation.
        Me.BackgroundWorker1.CancelAsync()

        ' Disable the Cancel button.
        cancelAsyncButton.Enabled = False
        sender = Nothing
        e = Nothing
    End Sub 'cancelAsyncButton_Click

    ' This event handler is where the actual work is done.
    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            ' Get the BackgroundWorker object that raised this event.
            Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
            ' Assign the result of the computation
            ' to the Result property of the DoWorkEventArgs
            ' object. This is will be available to the 
            ' RunWorkerCompleted eventhandler.
            e.Result = MainWorkerFunction(worker, e)
            Try
                worker.Dispose()
                sender = Nothing
                e = Nothing
            Catch ex As Exception
                Error_Handler(ex, "DoWork Worker Dispose")
            End Try
        Catch ex As Exception
            Error_Handler(ex, "backgroundWorker1_DoWork")
        End Try
    End Sub 'backgroundWorker1_DoWork

    ' This event handler deals with the results of the
    ' background operation.
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        busyworking = False
        ' First, handle the case where an exception was thrown.
        If Not (e.Error Is Nothing) Then
            Error_Handler(e.Error, "backgroundWorker1_RunWorkerCompleted")
        ElseIf e.Cancelled Then
            ' Next, handle the case where the user canceled the 
            ' operation.
            ' Note that due to a race condition in 
            ' the DoWork event handler, the Cancelled
            ' flag may not have been set, even though
            ' CancelAsync was called.
            Me.ProgressBar1.Value = 0
            inputlines_label.Text = "Cancelled"
            lastinputline_label.Text = "Cancelled"
            inputlines2_label.Text = "Cancelled"
            inputlines2Total_label.Text = "Cancelled"
            lastinputline2_label.Text = "Cancelled"
            Me.ToolTip1.SetToolTip(lastinputline_label, "Cancelled")
            Me.ToolTip1.SetToolTip(lastinputline2_label, "Cancelled")
            statusmessage = "Operation Cancelled"
            Status_Results_Handler("-- Operation Cancelled --")
        Else
            ' Finally, handle the case where the operation succeeded.

            Status_Handler(e.Result)

            Me.ProgressBar2.Value = 100
            Me.inputlines2_label.Text = secondary_inputlines & " (of " & secondary_inputlinesprecount & ") - " & secondary_newinputlines & " New; " & secondary_oldinputlines & " Removed"
            Me.inputlines2Total_label.Text = secondary_inputlinesTotal & " - " & secondary_newinputlinesTotal & " New; " & secondary_oldinputlinesTotal & " Removed"
            Me.lastinputline2_label.Text = secondary_lastinputline
            Me.ToolTip1.SetToolTip(lastinputline2_label, secondary_lastinputlineFull)


            Me.ProgressBar1.Value = 100
            Me.inputlines_label.Text = inputlines & " (of " & inputlinesprecount & ")"
            Me.lastinputline_label.Text = lastinputline
            Me.ToolTip1.SetToolTip(lastinputline_label, lastinputlineFull)
            Me.datelaunched_label.Text = Format(datelaunched, "yyyy/MM/dd HH:mm:ss") & " - " & Format(Now, "yyyy/MM/dd HH:mm:ss") & " (" & Now.Subtract(Me.datelaunched).TotalSeconds() & " s)"
            Me.LinkLabel1.Visible = True
            statusmessage = "Operation Completed"
            Status_Results_Handler("-- Operation Completed --")
        End If

        Status_Handler(statusmessage)
        Controls_Enabler("stop")
        sender = Nothing
        e = Nothing

        TimesExecuted = TimesExecuted + 1
        If TimesExecuted = 11 Then
            Shell((Application.StartupPath & "\" & "Application_Loader.exe").Replace("\\", "\"), AppWinStyle.NormalFocus, False, -1)
            Me.Close()
        End If
        Label9.Text = (CInt(Label9.Text) - 1).ToString
    End Sub 'backgroundWorker1_RunWorkerCompleted

    Private Sub Controls_Enabler(ByVal action As String)
        Select Case action.ToLower
            Case "run"
                Me.InputTargetFolder_Textbox.Enabled = False
                Me.Browse_Button.Enabled = False
                Me.startAsyncButton.Enabled = False
                Me.LinkLabel1.Enabled = False
                ' Disable the Cancel button.
                Me.cancelAsyncButton.Enabled = True
                Exit Select
            Case "stop"
                Me.InputTargetFolder_Textbox.Enabled = True
                Me.Browse_Button.Enabled = True
                Me.startAsyncButton.Enabled = True
                Me.LinkLabel1.Enabled = True
                ' Disable the Cancel button.
                Me.cancelAsyncButton.Enabled = False
                Exit Select
            Case Else
                Me.InputTargetFolder_Textbox.Enabled = False
                Me.Browse_Button.Enabled = False
                Me.startAsyncButton.Enabled = False
                Me.LinkLabel1.Enabled = False
                ' Disable the Cancel button.
                Me.cancelAsyncButton.Enabled = True
                Exit Select
        End Select
        action = Nothing

    End Sub

    ' This event handler updates the progress bar.
    Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        Me.ProgressBar2.Value = secondary_PercentComplete
        inputlines2_label.Text = secondary_inputlines & " (of " & secondary_inputlinesprecount & ") - " & secondary_newinputlines & " New; " & secondary_oldinputlines & " Removed"
        inputlines2Total_label.Text = secondary_inputlinesTotal & " - " & secondary_newinputlinesTotal & " New; " & secondary_oldinputlinesTotal & " Removed"
        lastinputline2_label.Text = secondary_lastinputline
        Me.ToolTip1.SetToolTip(lastinputline2_label, secondary_lastinputlineFull)

        Me.ProgressBar1.Value = e.ProgressPercentage
        inputlines_label.Text = inputlines & " (of " & inputlinesprecount & ")"
        lastinputline_label.Text = lastinputline
        Me.ToolTip1.SetToolTip(lastinputline_label, lastinputlineFull)

        datelaunched_label.Text = Format(datelaunched, "yyyy/MM/dd HH:mm:ss") & " - " & Format(Now, "yyyy/MM/dd HH:mm:ss") & " (" & Now.Subtract(Me.datelaunched).TotalSeconds() & " s)"
        If statusresults.Length > 0 Then
            Status_Results_Handler(statusresults.Trim)
        End If
        If statusmessage.Length > 0 Then
            Status_Handler(statusmessage)
        End If
        statusresults = ""
        statusmessage = ""
        sender = Nothing
        e = Nothing
    End Sub

    ' This is the method that does the actual work. 
    Function MainWorkerFunction(ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs) As String
        Dim result As String = ""
        Try

            ' Abort the operation if the user has canceled.
            ' Note that a call to CancelAsync may have set 
            ' CancellationPending to true just after the
            ' last invocation of this method exits, so this 
            ' code will not have the opportunity to set the 
            ' DoWorkEventArgs.Cancel flag to true. This means
            ' that RunWorkerCompletedEventArgs.Cancelled will
            ' not be set to true in your RunWorkerCompleted
            ' event handler. This is a race condition.
            If worker.CancellationPending Then
                e.Cancel = True
            End If

            secondary_inputlinesprecount = 0
            secondary_inputlines = 0
            secondary_inputlinesTotal = 0
            secondary_lastinputline = ""
            secondary_lastinputlineFull = ""
            secondary_highestPercentageReached = 0
            secondary_newinputlinesTotal = 0
            secondary_newinputlines = 0
            secondary_oldinputlines = 0
            secondary_oldinputlinesTotal = 0

            If Me.pretestdone = False Then
                statusmessage = "Calculating Operation Parameters"
                primary_PercentComplete = 0
                worker.ReportProgress(0)
                PreCount_Function()
                Me.pretestdone = True

            End If

            If worker.CancellationPending Then
                e.Cancel = True
            End If

            statusmessage = "Beginning Operation"
            primary_PercentComplete = 0
            worker.ReportProgress(0)




            inputlines = 0
            If Directory_Exists(InputTargetFolder_Textbox.Text) = True Then
                Dim rootdir As DirectoryInfo = New DirectoryInfo(InputTargetFolder_Textbox.Text)
                Dim basedir As DirectoryInfo

                For Each basedir In rootdir.GetDirectories()

                    If worker.CancellationPending Then
                        e.Cancel = True
                        Exit For
                    End If

                    If basedir.Name.Length = 4 And IsNumeric(basedir.Name) = True Then
                        If Integer.Parse(basedir.Name) > Year(Now) - 1 Then
                            If basedir.Exists = True Then
                                Dim departmentdir As DirectoryInfo
                                For Each departmentdir In basedir.GetDirectories
                                    If worker.CancellationPending Then
                                        e.Cancel = True
                                        Exit For
                                    End If
                                    If departmentdir.Exists = True Then
                                        Dim coursedir As DirectoryInfo
                                        For Each coursedir In departmentdir.GetDirectories
                                            If worker.CancellationPending Then
                                                e.Cancel = True
                                                Exit For
                                            End If
                                            If coursedir.Exists = True Then
                                                Dim projectdir As DirectoryInfo
                                                For Each projectdir In coursedir.GetDirectories
                                                    If worker.CancellationPending Then
                                                        e.Cancel = True
                                                        Exit For
                                                    End If
                                                    lastinputline = projectdir.Name
                                                    lastinputlineFull = projectdir.FullName
                                                    Dim percentComplete As Integer
                                                    percentComplete = 0
                                                    If inputlinesprecount > 0 Then
                                                        percentComplete = CSng(inputlines) / CSng(inputlinesprecount) * 100
                                                    Else
                                                        percentComplete = 100
                                                    End If
                                                    primary_PercentComplete = percentComplete
                                                    If percentComplete > highestPercentageReached Then
                                                        highestPercentageReached = percentComplete
                                                        statusmessage = "Checking Group Members"
                                                        worker.ReportProgress(percentComplete)
                                                    End If

                                                    If projectdir.Exists = True Then

                                                        Dim dowork As Boolean = False
                                                        Dim worktodo As String = "remove"
                                                        Dim accessgate As String = "close"

                                                        If File_Exists(projectdir.FullName & "\access.ini") = True Then
                                                            Dim filereader As StreamReader = New StreamReader(projectdir.FullName & "\access.ini")
                                                            Dim lineread As String = ""
                                                            Dim opentime As String = ""
                                                            Dim closetime As String = ""
                                                            While Not filereader.Peek = -1
                                                                lineread = filereader.ReadLine
                                                                If lineread.ToLower.StartsWith("open:") Then
                                                                    opentime = lineread.Replace("open:", "")
                                                                End If
                                                                If lineread.ToLower.StartsWith("close:") Then
                                                                    closetime = lineread.Replace("close:", "")
                                                                End If
                                                            End While
                                                            filereader.Close()
                                                            filereader = Nothing
                                                            Dim opendate, closedate As Date
                                                            If opentime.Length = 12 And closetime.Length = 12 Then
                                                                opendate = New Date(Integer.Parse(opentime.Substring(0, 4)), Integer.Parse(opentime.Substring(4, 2)), Integer.Parse(opentime.Substring(6, 2)), Integer.Parse(opentime.Substring(8, 2)), Integer.Parse(opentime.Substring(10, 2)), 0, 0)
                                                                closedate = New Date(Integer.Parse(closetime.Substring(0, 4)), Integer.Parse(closetime.Substring(4, 2)), Integer.Parse(closetime.Substring(6, 2)), Integer.Parse(closetime.Substring(8, 2)), Integer.Parse(closetime.Substring(10, 2)), 0, 0)
                                                                If Now > closedate Then
                                                                    'closed
                                                                    accessgate = "closed"
                                                                    Activity_Handler(coursedir.Name & " - " & projectdir.FullName & ": Access Gate is " & accessgate)
                                                                    statusresults = statusresults & (coursedir.Name & " - " & projectdir.Name & ": Access Gate is " & accessgate).Trim.Trim.Replace(vbCrLf, " ") & vbCrLf
                                                                Else
                                                                    'open
                                                                    accessgate = "open"
                                                                    Activity_Handler(coursedir.Name & " - " & projectdir.FullName & ": Access Gate is " & accessgate)
                                                                    statusresults = statusresults & (coursedir.Name & " - " & projectdir.Name & ": Access Gate is " & accessgate).Trim.Trim.Replace(vbCrLf, " ") & vbCrLf
                                                                End If
                                                            End If
                                                        End If

                                                        dowork = True


                                                        secondary_inputlinesprecount = 0
                                                        secondary_inputlines = 0
                                                        secondary_newinputlines = 0
                                                        secondary_oldinputlines = 0

                                                        secondary_lastinputline = ""
                                                        secondary_highestPercentageReached = 0

                                                        Dim apptorun As String = ""
                                                        apptorun = """" & (Application.StartupPath & "\Commerce Courses Student List Extractor.exe").Replace("\\", "\") & """"
                                                        Dim students As String() = ApplicationLauncher(apptorun, """" & coursedir.Name & """").Split(vbCrLf)
                                                        If students.Length > 0 Then



                                                            Dim studentdir As DirectoryInfo

                                                            secondary_inputlinesprecount = students.Length - 1
                                                            Dim newdetected As Boolean = False
                                                            For Each str As String In students
                                                                If worker.CancellationPending Then
                                                                    e.Cancel = True
                                                                    Exit For
                                                                End If
                                                                Dim student As String = str.Trim.ToUpper.Replace(coursedir.Name.ToUpper & "_G ", "")
                                                                If student.StartsWith("STATUS") = True Then
                                                                    ' Report progress as a percentage of the total task.
                                                                    If secondary_inputlinesprecount > 0 Then
                                                                        secondary_PercentComplete = CSng(secondary_inputlines) / CSng(secondary_inputlinesprecount) * 100
                                                                    Else
                                                                        secondary_PercentComplete = 100
                                                                    End If

                                                                    If secondary_PercentComplete > secondary_highestPercentageReached Then
                                                                        secondary_highestPercentageReached = secondary_PercentComplete
                                                                        statusmessage = "Checking Group Members"
                                                                        worker.ReportProgress(primary_PercentComplete)
                                                                    End If
                                                                    Exit For
                                                                End If

                                                                studentdir = New DirectoryInfo((projectdir.FullName & "\" & student).Replace("\\", "\"))
                                                                If studentdir.Exists = False Then
                                                                    'foldercount = foldercount + 1


                                                                    newdetected = True
                                                                    Activity_Handler("New Student Detected: Creating Folder for " & student & " (" & coursedir.Name & " - " & projectdir.Name & ")")
                                                                    statusresults = statusresults & "New Student Detected: " & student & vbCrLf
                                                                    studentdir.Create()
                                                                    secondary_newinputlinesTotal = secondary_newinputlinesTotal + 1
                                                                    secondary_newinputlines = secondary_newinputlines + 1
                                                                    'do work
                                                                    If dowork = True Then
                                                                        If RadioButton2.Checked = True Then
                                                                            If File_Exists((Application.StartupPath & "\SETTRUST.exe").Replace("\\", "\")) = True Then
                                                                                apptorun = ""
                                                                                Dim apptorunArgs As String = ""
                                                                                apptorun = """" & (Application.StartupPath & "\SETTRUST.exe").Replace("\\", "\") & """"
                                                                                If accessgate = "open" Then
                                                                                    apptorunArgs = """" & studentdir.FullName & """ " & "[RWECMF]" & " " & "." & studentdir.Name & ".Students.com.main.uct"

                                                                                Else
                                                                                    apptorunArgs = """" & studentdir.FullName & """ " & "[RF]" & " " & "." & studentdir.Name & ".Students.com.main.uct"
                                                                                End If
                                                                                Dim appresult As String = ApplicationLauncher(apptorun, apptorunArgs)
                                                                                'statusresults = statusresults & appresult.Trim.Trim.Replace(vbCrLf, " ") & vbCrLf
                                                                                Activity_Handler(appresult.Trim.Trim.Replace(vbCrLf, " "))
                                                                            End If
                                                                        End If
                                                                        
                                                                    End If

                                                                End If

                                                                secondary_inputlines = secondary_inputlines + 1
                                                                secondary_inputlinesTotal = secondary_inputlinesTotal + 1
                                                                secondary_lastinputline = studentdir.Name
                                                                secondary_lastinputlineFull = studentdir.FullName
                                                                secondary_PercentComplete = 0
                                                                ' Report progress as a percentage of the total task.
                                                                If secondary_inputlinesprecount > 0 Then
                                                                    secondary_PercentComplete = CSng(secondary_inputlines) / CSng(secondary_inputlinesprecount) * 100
                                                                Else
                                                                    secondary_PercentComplete = 100
                                                                End If

                                                                If secondary_PercentComplete > secondary_highestPercentageReached Then
                                                                    secondary_highestPercentageReached = secondary_PercentComplete
                                                                    statusmessage = "Checking Group Members"
                                                                    worker.ReportProgress(primary_PercentComplete)
                                                                End If

                                                            Next

                                                            If newdetected = False Then
                                                                Activity_Handler("No New Students Detected")
                                                                statusresults = statusresults & "No New Students Detected" & vbCrLf
                                                            End If

                                                            secondary_PercentComplete = 0
                                                            ' Report progress as a percentage of the total task.
                                                            If secondary_inputlinesprecount > 0 Then
                                                                secondary_PercentComplete = CSng(secondary_inputlines) / CSng(secondary_inputlinesprecount) * 100
                                                            Else
                                                                secondary_PercentComplete = 100
                                                            End If

                                                            If secondary_PercentComplete > secondary_highestPercentageReached Then
                                                                secondary_highestPercentageReached = secondary_PercentComplete
                                                                statusmessage = "Checking Group Members"
                                                                worker.ReportProgress(primary_PercentComplete)
                                                            End If
                                                            studentdir = Nothing
                                                        End If
                                                        Dim studentexists As Boolean = False
                                                        Dim currentstudentdirs As DirectoryInfo
                                                        Dim olddetected As Boolean = False
                                                        For Each currentstudentdirs In projectdir.GetDirectories()
                                                            If worker.CancellationPending Then
                                                                e.Cancel = True
                                                                Exit For
                                                            End If
                                                            studentexists = False
                                                            For Each str As String In students
                                                                If worker.CancellationPending Then
                                                                    e.Cancel = True
                                                                    Exit For
                                                                End If
                                                                Dim student As String = str.Trim.ToUpper.Replace(coursedir.Name.ToUpper & "_G ", "")
                                                                If student = currentstudentdirs.Name Then
                                                                    studentexists = True
                                                                    Exit For
                                                                End If
                                                            Next
                                                            If studentexists = False Then
                                                                olddetected = True
                                                                If File_Exists((Application.StartupPath & "\SETTRUST.exe").Replace("\\", "\")) = True Then
                                                                    If RadioButton2.Checked = True Then
                                                                        apptorun = ""
                                                                        Dim apptorunArgs As String = ""
                                                                        apptorun = """" & (Application.StartupPath & "\SETTRUST.exe").Replace("\\", "\") & """"
                                                                        apptorunArgs = """" & currentstudentdirs.FullName & """ " & "[RF]" & " " & "." & currentstudentdirs.Name & ".Students.com.main.uct"
                                                                        Dim appresult As String = ApplicationLauncher(apptorun, apptorunArgs)
                                                                        statusresults = statusresults & "Deregistered Student: " & currentstudentdirs.Name & vbCrLf
                                                                    End If
                                                                    Activity_Handler("Deregistered Student Detected: Removing Folder for " & currentstudentdirs.Name & " (" & coursedir.Name & " - " & projectdir.Name & ")")
                                                                    'Activity_Handler(appresult.Trim.Trim.Replace(vbCrLf, " "))
                                                                    secondary_oldinputlinesTotal = secondary_oldinputlinesTotal + 1
                                                                    secondary_oldinputlines = secondary_oldinputlines + 1
                                                                    'msgbox (rootdir.FullName & "\Deregistered Student Bin\" & basedir.Name & "\" & coursedir.Name & "\" & projectdir.Name & "\" & currentstudentdirs.Name)
                                                                    My.Computer.FileSystem.MoveDirectory(currentstudentdirs.FullName, rootdir.FullName & "\Deregistered Student Bin\" & basedir.Name & "\" & coursedir.Name & "\" & projectdir.Name & "\" & currentstudentdirs.Name, True)
                                                                End If

                                                            End If
                                                        Next
                                                        currentstudentdirs = Nothing
                                                        students = Nothing
                                                        If olddetected = False Then
                                                            Activity_Handler("No Deregistered Students Detected")
                                                            statusresults = statusresults & "No Deregistered Students Detected" & vbCrLf
                                                        End If
                                                    End If




                                                    lastinputline = projectdir.Name
                                                    lastinputlineFull = projectdir.FullName
                                                    inputlines = inputlines + 1

                                                    ' Report progress as a percentage of the total task.
                                                    percentComplete = 0
                                                    If inputlinesprecount > 0 Then
                                                        percentComplete = CSng(inputlines) / CSng(inputlinesprecount) * 100
                                                    Else
                                                        percentComplete = 100
                                                    End If
                                                    primary_PercentComplete = percentComplete
                                                    If percentComplete > highestPercentageReached Then
                                                        highestPercentageReached = percentComplete
                                                        statusmessage = "Checking Group Members"
                                                        worker.ReportProgress(percentComplete)
                                                    End If

                                                Next
                                                projectdir = Nothing
                                            End If
                                        Next
                                        coursedir = Nothing
                                    End If
                                Next
                                departmentdir = Nothing
                            End If
                        End If
                    End If
                Next
                basedir = Nothing
                rootdir = Nothing
            Else
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "Connection to the Root Directory cannot be established."
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
            End If



        Catch ex As Exception
            Error_Handler(ex, "MainWorkerFunction")
        End Try

        Try
            worker.Dispose()
            worker = Nothing
            e = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Dispose Worker")
        End Try

        Return result

    End Function

    Private Sub PreCount_Function()
        Try
            inputlinesprecount = 0

            If Directory_Exists(InputTargetFolder_Textbox.Text) = True Then
                Dim rootdir As DirectoryInfo = New DirectoryInfo(InputTargetFolder_Textbox.Text)
                Dim basedir As DirectoryInfo
                For Each basedir In rootdir.GetDirectories()
                    If basedir.Name.Length = 4 And IsNumeric(basedir.Name) = True Then
                        If Integer.Parse(basedir.Name) > Year(Now) - 1 Then
                            If basedir.Exists = True Then
                                Dim departmentdir As DirectoryInfo
                                For Each departmentdir In basedir.GetDirectories
                                    If departmentdir.Exists = True Then
                                        Dim coursedir As DirectoryInfo
                                        For Each coursedir In departmentdir.GetDirectories
                                            If coursedir.Exists = True Then
                                                Dim projectdir As DirectoryInfo
                                                For Each projectdir In coursedir.GetDirectories
                                                    inputlinesprecount = inputlinesprecount + 1
                                                Next
                                                projectdir = Nothing
                                            End If
                                        Next
                                        coursedir = Nothing
                                    End If
                                Next
                                departmentdir = Nothing
                            End If
                        End If
                    End If
                Next
                basedir = Nothing
                rootdir = Nothing
            End If

        Catch ex As Exception
            Error_Handler(ex, "PreCount_Function")
        End Try
    End Sub






    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            If File_Exists((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt") = True Then
                Dim systemDirectory As String
                systemDirectory = System.Environment.SystemDirectory
                Dim finfo As FileInfo = New FileInfo((systemDirectory & "\notepad.exe").Replace("\\", "\"))
                If finfo.Exists = True Then
                    Dim apptorun As String
                    apptorun = """" & (systemDirectory & "\notepad.exe").Replace("\\", "\") & """ """ & (Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt" & """"
                    Dim procID As Integer = Shell(apptorun, AppWinStyle.NormalFocus, False)
                End If
                finfo = Nothing
            End If
        Catch ex As Exception
            Error_Handler(ex, "LinkLabel1_LinkClicked")
        End Try
    End Sub










    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False


            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter


            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            'myProcess.StartInfo.FileName = AppToRun

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If





            sIn.Close()
            sIn.Dispose()
            sOut.Close()
            sOut.Dispose()
            sErr.Close()
            sErr.Dispose()
            myProcess.Close()
            myProcess.Dispose()
            AppToRun = Nothing

        Catch ex As Exception
            Error_Handler(ex, "DosShellCommand")
        End Try

        Return s
    End Function

    Private Function ApplicationLauncher(ByVal AppToRun As String, ByVal apptorunArgs As String) As String
        Dim s As String = ""
        Try

            Dim myProcess As Process = New Process


            myProcess.StartInfo.UseShellExecute = False


            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter


            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            myProcess.StartInfo.FileName = AppToRun
            myProcess.StartInfo.Arguments = apptorunArgs

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            sIn.Close()
            sIn.Dispose()
            sOut.Close()
            sOut.Dispose()
            sErr.Close()
            sErr.Dispose()
            myProcess.Close()
            myProcess.Dispose()
            AppToRun = Nothing
            apptorunArgs = Nothing
        Catch ex As Exception
            Error_Handler(ex, "ApplicationLauncher")
            s = ""
        End Try
        Return s
    End Function


    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        StatusResults_RichtextBox.Clear()
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            inputseconds_label.Text = (Integer.Parse(inputseconds_label.Text) - 1).ToString
            If Integer.Parse(inputseconds_label.Text) = 0 Then
                'do action
                StartWorker()
                inputseconds_label.Text = "86400"
            End If
            sender = Nothing
            e = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Timer1_Tick")
        End Try
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        If Not Me.WindowState = FormWindowState.Normal Then
            show_main_application()
        Else
            hide_main_application()
        End If

    End Sub

    Private Sub ExitApplicationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitApplicationToolStripMenuItem.Click
        Try
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "ExitApplicationToolStripMenuItem_Click")
        End Try
    End Sub

    Private Sub DisplayApplicationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DisplayApplicationToolStripMenuItem.Click
        Try
            show_main_application()
        Catch ex As Exception
            Error_Handler(ex, "DisplayApplicationToolStripMenuItem_Click")
        End Try
    End Sub

    Private Sub show_main_application()
        Try
            Me.Opacity = 1
            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal
            'Me.WindowState = FormWindowState.Normal
            'Me.Visible = True
            'Me.Opacity = 100
        Catch ex As Exception
            Error_Handler(ex, "show_main_application")
        End Try
    End Sub

    Private Sub hide_main_application()
        Try
            Me.WindowState = FormWindowState.Minimized
            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
            'Me.WindowState = FormWindowState.Minimized
            'Me.Visible = False
            'Me.Opacity = 0
        Catch ex As Exception
            Error_Handler(ex, "hide_main_application")
        End Try
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Try
            hide_main_application()
        Catch ex As Exception
            Error_Handler(ex, "ToolStripMenuItem2_Click")
        End Try
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        hide_main_application()
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        Status_Handler("Server Type Set")
    End Sub
End Class
