'-Project Title:The SCALA Project
'-Project Version: v5.0
'-Programmer: Jerome de las Alas
'-Date Started: April 11 2013
'-Last Update: May 26 2014

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.IO.Compression

Public Class Form1
    Dim currPath As String = ""
    Dim myPath() As String
    Dim activeTab As Integer = 1
    Dim scanlogPresent As Integer = 2
    Dim buildlogPresent As Integer = 2
    Dim fileSPresent As Integer = 2

    'enable drag function for button
    Private Sub Button5_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Button5.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    'drag and drop function of button
    Private Sub Button5_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Button5.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            myPath = e.Data.GetData(DataFormats.FileDrop)


            currPath = myPath(0)

            'test if the dragged file is a zip file
            If currPath.Contains(".zip") Then
                'extract the zip file
                Dim zipPath As String = currPath
                Dim extractPath As String = Replace(currPath, ".zip", "")

                If (Not System.IO.Directory.Exists(extractPath)) Then
                    System.IO.Directory.CreateDirectory(extractPath)
                End If

                Using archive As ZipArchive = ZipFile.OpenRead(zipPath)
                    For Each entry As ZipArchiveEntry In archive.Entries
                        ToolStripStatusLabel3.Text = "Extracting " + entry.FullName
                        If System.IO.File.Exists(Path.Combine(extractPath, entry.FullName)) Then
                            System.IO.File.Delete(Path.Combine(extractPath, entry.FullName))
                        End If
                        entry.ExtractToFile(Path.Combine(extractPath, entry.FullName))
                    Next
                End Using
                currPath = extractPath
            End If

            ToolStripStatusLabel3.Text = currPath


            ''check if the folder contains the log files
            ''call checker function
            If (CheckFolderValidity(currPath) = 0) Then
                MsgBox("Invalid Directory", MsgBoxStyle.OkOnly, "ERROR")
            Else
                ''case 1 - 3 disable scan if file does not exist
                Call AnalyzeData()
            End If
        End If
    End Sub

    Private Function CheckFolderValidity(ByVal FilePath As String) As Integer
        Dim buildflag As Boolean = False
        Dim scanflag As Boolean = False
        Dim fileSflag As Boolean = False

        If Directory.Exists(FilePath) Then
            'directory is valid
        Else
            MsgBox("Invalid Folder or Directory!", MsgBoxStyle.OkOnly, "Invalid Directory")
            Return 0
        End If

        If File.Exists(FilePath + "\build.log") Then
            'directory contains build.log
            buildflag = True
            buildlogPresent = 1
        Else
            buildflag = False
            buildlogPresent = 0
            MsgBox("build.log file not found in folder", MsgBoxStyle.OkOnly, "WARNING")
        End If

        If File.Exists(FilePath + "\scan.log") Then
            'directory contains scan.log
            scanlogPresent = 1
            scanflag = True
        Else
            scanlogPresent = 0
            scanflag = False
            MsgBox("scan.log file not found in folder", MsgBoxStyle.OkOnly, "WARNING")
        End If

        If File.Exists(FilePath + "\filesScanned.txt") Then
            'directory contains filesScanned.txt
            fileSflag = True
            fileSPresent = 1
        Else
            fileSPresent = 0
        End If


        If buildflag = True And scanflag = False Then
            'only buildlog exists
            Return 1
        ElseIf buildflag = False And scanflag = True Then
            'only scanlog exists
            Return 2
        ElseIf buildflag = True And scanflag = True Then
            'both files exist
            Return 3
        Else
            'no readable files present
            Return 0
        End If


    End Function


    '----------TAB#1----------'
    'BASIC STATS TAB

    'function to analyze data dropped to application
    Private Sub AnalyzeData()

        Dim f1 As Boolean = False, f2 As Boolean = False, f3 As Boolean = False, f4 As Boolean = False
        Dim filesProcessedCtr As Integer = 0
        Dim buildtimestampstring As String = ""
        Dim scantimestampstring As String = ""

        ProgressBar1.Value = 0
        ProgressBar1.Maximum = 100
        ProgressBar1.Step = 20

        ListBox1.Items.Clear()
        'ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
        ListView1.Items.Clear()
        ListView2.Items.Clear()
        ListView3.Items.Clear()
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        CheckBox6.Checked = False
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox5.Text = ""
        TextBox2.Text = ""
        TextBox1.Text = ""
        ProgressBar1.PerformStep()

        If (scanlogPresent = 1) Then
            Using reader As New StreamReader(currPath + "\scan.log")
                Dim flag As Integer = 0
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    If line.Contains("vulns") Then
                        TextBox3.Text = line
                        f1 = True
                    End If
                    If line.Contains("Lines of Code:") Then
                        TextBox4.Text = line
                        f2 = True
                    End If
                    If line.Contains("Timers: ") Then
                        TextBox7.Text = line
                        f3 = True
                    End If
                    If line.Contains("WorkQueue.size: ") Then
                        TextBox10.Text = line
                        TextBox1.Text = line
                        f4 = True
                    End If
                    If f1 = True And f2 = True And f3 = True And f4 = True Then
                        Exit While
                    End If
                End While
            End Using
        End If
        ProgressBar1.PerformStep()

        If (buildlogPresent = 1) Then
            ListView1.Items.Clear()
            Using reader As New StreamReader(currPath + "\build.log")
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    Dim lineLookAhead As String


                    If line.Contains("Translating ") Then
                        ListBox4.Items.Add(line)
                    End If
                    If line.Contains("Fidelity") Then
                        ListBox5.Items.Add(line)
                        CheckBox6.Checked = True
                    End If

                    If line.Contains("Processing") Then
                        filesProcessedCtr = filesProcessedCtr + 1
                    End If

                    If (line.Contains(" WARNING ") Or line.Contains(" SEVERE ")) Then
                        'process the line to eliminate unwanted warnings
                        lineLookAhead = reader.ReadLine()
                        Call uniqErr(line, 0, lineLookAhead)

                        'missing dependencies via 1219 and 1202
                        If (line.Contains("WARNING 1219") Or line.Contains("WARNING 1202")) Then
                            'process the line to generate unknown symbols
                            Dim MatchObj As Match = Regex.Match(lineLookAhead, "^.*'")
                            If MatchObj.Success = True Then
                                'filter out same entries
                                Call UnknownSymbol(line, lineLookAhead)
                            End If
                        End If

                        'missing dependencies via 1237
                        ''PROCESS WARNING 1237 missing dependencies error
                        If (line.Contains("WARNING 1237")) Then
                            'line contains warning 1237
                            line = reader.ReadLine()
                            While Not (line.Contains("[20"))
                                ListBox1.Items.Add(line)
                                line = reader.ReadLine()
                            End While
                        End If
                    End If

                    If line = "Args:" Then
                        lineLookAhead = reader.ReadLine()
                        TextBox12.Text = lineLookAhead
                        'ViewExcludes script
                        Dim regexPatternEx As String = """-exclude"", ""([^""]*)"""
                        Dim MatchObjEx As Match = Regex.Match(lineLookAhead, regexPatternEx)
                        Dim excludeEntry As String = ""

                        For Each MatchObjEx In Regex.Matches(lineLookAhead, regexPatternEx)
                            excludeEntry = Regex.Replace(MatchObjEx.Value, """-exclude"", """, "")
                            excludeEntry = Regex.Replace(excludeEntry, """", "")
                            ListBox3.Items.Add(excludeEntry)
                        Next
                    End If

                    If line.StartsWith("[20") Then
                        buildtimestampstring = line
                    End If

                End While
            End Using
        End If
        TextBox5.Text = buildtimestampstring
        ProgressBar1.PerformStep()

        If (scanlogPresent = 1) Then
            ListView2.Items.Clear()
            Using reader As New StreamReader(currPath + "\scan.log")
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    Dim lineLookAhead As String
                    If (line.Contains(" WARNING ") Or line.Contains(" SEVERE ")) Then
                        'process the line to eliminate unwanted warnings
                        lineLookAhead = reader.ReadLine()
                        Call uniqErr(line, 1, lineLookAhead)
                    End If
                    If line = "Args:" Then
                        lineLookAhead = reader.ReadLine()
                        TextBox11.Text = lineLookAhead
                    End If
                    If line.StartsWith("[20") Then
                        scantimestampstring = line
                    End If
                End While
            End Using
        End If
        TextBox2.Text = scantimestampstring
        ProgressBar1.PerformStep()

        Dim fileCounter As Integer = 0
        Dim maxdepth As Integer = 0
        Dim curdepth As Integer = 0
        If (fileSPresent = 1) Then
            ListView3.Items.Clear()
            Using reader As New StreamReader(currPath + "\filesScanned.txt")
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()


                    fileCounter = fileCounter + 1
                    curdepth = line.Split("\").Length - 1
                    If curdepth > maxdepth Then
                        maxdepth = curdepth
                    End If

                    'get the characters after the last \
                    'check if the characters contain .
                    'if yes, get the characters after the .
                    'if no, get all characters

                    Dim curfileName As String = ""
                    curfileName = line.Substring(line.LastIndexOf("\") + 1)
                    Call addFileType(curfileName)
                End While
            End Using
        End If
        TextBox8.Text = fileCounter
        TextBox9.Text = maxdepth
        ProgressBar1.PerformStep()


        If ListBox1.Items.Count > 0 Then
            CheckBox4.Checked = True
        End If
        If ListBox3.Items.Count > 0 Then
            CheckBox5.Checked = True
        End If
        Dim filesProcessedCount As String = "Total Number of Files Translated by SCA: " + Convert.ToString(filesProcessedCtr)
        ListBox5.Items.Add(filesProcessedCount)

    End Sub

    'UniqErr script
    Private Sub uniqErr(ByVal line As String, ByVal stage As Integer, ByVal lineLookAhead As String)

        Dim regexPattern As String = "(WARNING|SEVERE) \d+"
        Dim sample As String
        Dim count As Integer
        Dim MatchObj As Match = Regex.Match(line, regexPattern)

        If MatchObj.Success = True Then
            If stage = 0 Then
                'build stage
                Dim item1 As ListViewItem = ListView1.FindItemWithText(MatchObj.Value)
                If (item1 IsNot Nothing) Then
                    'add to item count
                    sample = ListView1.Items(item1.Index).SubItems(1).Text
                    count = Convert.ToInt32(sample)
                    count = count + 1
                    ListView1.Items(item1.Index).SubItems(1).Text = count
                Else
                    'add entry
                    ListView1.Items.Add(MatchObj.Value)
                    ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(1)
                    ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(lineLookAhead)
                End If
            End If

            If stage = 1 Then
                'scan stage
                Dim item1 As ListViewItem = ListView2.FindItemWithText(MatchObj.Value)
                If (item1 IsNot Nothing) Then
                    'add to item count
                    sample = ListView2.Items(item1.Index).SubItems(1).Text
                    count = Convert.ToInt32(sample)
                    count = count + 1
                    ListView2.Items(item1.Index).SubItems(1).Text = count
                Else
                    'add entry
                    ListView2.Items.Add(MatchObj.Value)
                    ListView2.Items(ListView2.Items.Count - 1).SubItems.Add(1)
                    ListView2.Items(ListView2.Items.Count - 1).SubItems.Add(lineLookAhead)
                End If
            End If
        End If
    End Sub

    'display output of AnalyzeData function on-change
    Private Sub ListView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        Dim currSelectedPeek As ListView.SelectedListViewItemCollection = Me.ListView1.SelectedItems
        Dim item As ListViewItem
        Dim peek As String = ""

        For Each item In currSelectedPeek
            peek = item.SubItems(2).Text
        Next

        TextBox6.Text = peek
    End Sub
    Private Sub ListView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView2.SelectedIndexChanged
        Dim currSelectedPeek As ListView.SelectedListViewItemCollection = Me.ListView2.SelectedItems
        Dim item As ListViewItem
        Dim peek As String = ""

        For Each item In currSelectedPeek
            peek = item.SubItems(2).Text
        Next

        TextBox6.Text = peek
    End Sub


    '----------TAB#2----------'
    'DEPENDENCIES TAB
    'populate the dependencies tab
    Private Sub UnknownSymbol(ByVal line As String, ByVal lineLookAhead As String)
        'ListBox1.Items.Add(MatchObj)
        Dim regexPattern As String = "'([^']*)'"
        Dim MatchObj As Match = Regex.Match(lineLookAhead, regexPattern)
        Dim dependencyEntry As String = ""
        'excludeEntry = Regex.Replace(MatchObjEx.Value, """-exclude"", """, "")

        If MatchObj.Success = True Then
            'build stage
            dependencyEntry = Regex.Replace(MatchObj.Value, "'", "")
            Dim item1 As Integer = -1
            'item1 = ListBox1.FindStringExact(MatchObj.Value)
            item1 = ListBox1.FindStringExact(dependencyEntry)
            If (item1 <> -1) Then
                'do nothing
            Else
                'add entry
                'ListBox1.Items.Add(MatchObj.Value)
                ListBox1.Items.Add(dependencyEntry)
            End If
        End If
    End Sub

    '----------TAB#3----------'
    'EXCLUDES TAB

    '----------TAB#4----------'
    'SEARCH TAB
    'DISCONTINUED FUNCTIONALITIY
    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ListBox2.Items.Clear()
    '    Dim searchString As String
    '    searchString = TextBox5.Text
    '    'If RadioButton2.Checked = True Then
    '    'Dim pattern As String = "[~#%&*{}/()<>?|\-^]"
    '    'searchString = Regex.Replace(searchString, pattern, "")
    '    'End If
    '    'validate of search box text is a whitespace
    '    If ((String.IsNullOrEmpty(searchString)) Or (searchString = " ") Or (searchString = ControlChars.Tab) Or (searchString = ControlChars.CrLf) Or (searchString = ControlChars.NewLine)) Then
    '        MsgBox("Search field was left blank or the search string is not allowed!", MsgBoxStyle.OkOnly, "Search Field Error")
    '        Exit Sub
    '    End If

    '    If ((RadioButton2.Checked = True) And (searchString = "*.")) Then
    '        MsgBox("Invalid Regex  String!", MsgBoxStyle.OkOnly, "Search Field Error")
    '        Exit Sub
    '    End If

    '    If CheckBox1.Checked = True Then 'and file exists @file-manifest
    '        'file-manifest should exist
    '    End If

    '    If CheckBox2.Checked = True Then 'and file exists
    '        ListBox2.Items.Add("[---------------------BUILD LOG RESULTS---------------------]")
    '        If buildlogPresent = 1 Then
    '            Using reader As New StreamReader(currPath + "\build.log")
    '                Dim flag As Integer = 0
    '                While Not reader.EndOfStream
    '                    Dim line As String = reader.ReadLine()
    '                    If (line.Contains(searchString) And RadioButton1.Checked = True) Then
    '                        ListBox2.Items.Add(line)
    '                    End If
    '                    Dim MatchObj As Match = Regex.Match(line, searchString)
    '                    If ((RadioButton2.Checked = True) And (MatchObj.Success = True)) Then
    '                        ListBox2.Items.Add(line)
    '                    End If
    '                End While
    '            End Using
    '        End If
    '    End If

    '    If CheckBox3.Checked = True Then 'and file exists
    '        ListBox2.Items.Add("[---------------------SCAN LOG RESULTS---------------------]")
    '        If scanlogPresent = 1 Then
    '            Using reader As New StreamReader(currPath + "\scan.log")
    '                Dim flag As Integer = 0
    '                While Not reader.EndOfStream
    '                    Dim line As String = reader.ReadLine()
    '                    If line.Contains(searchString) Then
    '                        ListBox2.Items.Add(line)
    '                    End If
    '                End While
    '            End Using
    '        End If
    '    End If
    'End Sub
    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Call cleanUpTab2()
    'End Sub
    'Private Sub cleanUpTab2()
    '    TextBox5.Text = ""
    '    ListBox2.Items.Clear()
    'End Sub

    '----------TAB#5----------'
    'FILE MANIFEST ANALYSIS TAB
    'file manifest analysis function not yet implemented
    'addFileType script
    Private Sub addFileType(ByVal curFile As String)

        Dim sample As String
        Dim count As Integer
        Dim fileExtension As String

        If (curFile.Contains(".")) Then
            fileExtension = curFile.Substring(curFile.LastIndexOf(".") + 1)
        Else
            fileExtension = curFile
        End If

        Dim item1 As ListViewItem = ListView3.FindItemWithText(fileExtension)
        If (item1 IsNot Nothing) Then
            'add to item count
            sample = ListView3.Items(item1.Index).SubItems(1).Text
            count = Convert.ToInt32(sample)
            count = count + 1
            ListView3.Items(item1.Index).SubItems(1).Text = count
        Else
            'add entry
            ListView3.Items.Add(fileExtension)
            ListView3.Items(ListView3.Items.Count - 1).SubItems.Add(1)
        End If
    End Sub

    Private Sub openFileS_Click(sender As Object, e As EventArgs) Handles openFileS.Click
        'open filesScanned.txt
    End Sub

    '----------TAB#6----------'
    'COMPARE ARGUMENTS TAB
    'enable drag function for button (compare tab)
    Private Sub Button6_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Button6.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    'drag and drop function of button  (compare tab)
    Private Sub Button6_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Button6.DragDrop
        Dim currPath1 As String = ""
        Dim myPath1() As String = {"", ""}
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            myPath1 = e.Data.GetData(DataFormats.FileDrop)
            currPath1 = myPath1(0)

            'test if the dragged file is a zip file
            If currPath1.Contains(".zip") Then
                'extract the zip file
                Dim zipPath As String = currPath1
                Dim extractPath As String = Replace(currPath1, ".zip", "")

                If (Not System.IO.Directory.Exists(extractPath)) Then
                    System.IO.Directory.CreateDirectory(extractPath)
                End If

                Using archive As ZipArchive = ZipFile.OpenRead(zipPath)
                    For Each entry As ZipArchiveEntry In archive.Entries
                        If System.IO.File.Exists(Path.Combine(extractPath, entry.FullName)) Then
                            System.IO.File.Delete(Path.Combine(extractPath, entry.FullName))
                        End If
                            entry.ExtractToFile(Path.Combine(extractPath, entry.FullName))
                    Next
                End Using
                currPath1 = extractPath
            End If

            ''check if the folder contains the log files
            ''call checker function
            If (CheckFolderValidity(currPath1) = 0) Then
                MsgBox("Invalid Directory", MsgBoxStyle.OkOnly, "ERROR")
            Else
                ''Run function to get arguments here
                If buildlogPresent = 1 Then
                    Using reader As New StreamReader(currPath1 + "\build.log")
                        While Not reader.EndOfStream
                            Dim line As String = reader.ReadLine()
                            Dim lineLookAhead As String

                            If line = "Args:" Then
                                lineLookAhead = reader.ReadLine()
                                TextBox14.Text = lineLookAhead
                            End If
                        End While
                    End Using
                End If
                If scanlogPresent = 1 Then
                    Using reader As New StreamReader(currPath1 + "\scan.log")
                        While Not reader.EndOfStream
                            Dim line As String = reader.ReadLine()
                            Dim lineLookAhead As String

                            If line = "Args:" Then
                                lineLookAhead = reader.ReadLine()
                                TextBox13.Text = lineLookAhead
                            End If
                        End While
                    End Using
                End If
            End If
        End If
    End Sub

    '----------TAB#7----------'
    'TIME STAMPS TAB
    'gets the latest time stamp for both build and scan logs
    'populate the time stamps tab


    '----------TAB SELECTION FUNCTION----------'
    'SWITCH TABS
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            activeTab = 1
        ElseIf TabControl1.SelectedIndex = 1 Then
            activeTab = 2
        ElseIf TabControl1.SelectedIndex = 2 Then
            activeTab = 3
        ElseIf TabControl1.SelectedIndex = 3 Then
            activeTab = 4
        ElseIf TabControl1.SelectedIndex = 4 Then
            activeTab = 5
        ElseIf TabControl1.SelectedIndex = 5 Then
            activeTab = 6
        Else
            activeTab = 1
        End If
    End Sub

    '----------SAVE TO FILE----------'
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'save data into text file
        If (CheckFolderValidity(currPath) = 0) Then
            MsgBox("The current directory does NOT contain SCA log files. Check if you dragged the proper folder to the application.", MsgBoxStyle.OkOnly, "ERROR!")
            Exit Sub
        End If

        If activeTab = 1 Then
            'export data in tab#1
            Dim myPath As String = currPath + "\scala-basic_stats.txt"
            Dim sb As New StringBuilder()

            sb.AppendLine("BASIC STATS TAB")
            sb.AppendLine()
            sb.Append(TextBox3.Text)
            sb.AppendLine()
            sb.Append(TextBox4.Text)
            sb.AppendLine()
            sb.Append(TextBox7.Text)
            sb.AppendLine()
            sb.Append(TextBox10.Text)
            sb.AppendLine()
            sb.AppendLine()
            sb.AppendLine("UNIQUE ERRORS")
            sb.AppendLine()
            sb.AppendLine("Build Log:")

            If ListView1.Items.Count > 0 Then
                For Each item In ListView1.Items
                    sb.AppendLine(item.SubItems(1).Text & "  " & item.SubItems(0).Text)
                Next
            End If
            sb.AppendLine()
            sb.AppendLine("Scan Log:")
            If ListView2.Items.Count > 0 Then
                For Each item In ListView2.Items
                    sb.AppendLine(item.SubItems(1).Text & "  " & item.SubItems(0).Text)
                Next
            End If
            Using outfile As New StreamWriter(myPath)
                outfile.Write(sb.ToString())
            End Using
            MsgBox("Write Successful. File scala-basic_stats.txt was written in your current logs directory.", MsgBoxStyle.OkOnly, "Write Successful")


        ElseIf activeTab = 2 Then
            'export data in tab#2
            Dim myPath As String = currPath + "\scala-dependencies.txt"
            Dim sb As New StringBuilder()

            sb.AppendLine("MISSING DEPENDENCIES TAB")
            sb.AppendLine()
            sb.AppendLine("Missing Dependencies")
            sb.AppendLine()
            sb.AppendLine()
            If ListBox1.Items.Count > 0 Then
                For Each item In ListBox1.Items
                    sb.AppendLine(item.ToString())
                Next
            End If
            Using outfile As New StreamWriter(myPath)
                outfile.Write(sb.ToString())
            End Using
            MsgBox("Write Successful. File scala-dependencies.txt was written in your current logs directory.", MsgBoxStyle.OkOnly, "Write Successful")


        ElseIf activeTab = 3 Then
            'export data in tab#3
            Dim myPath As String = currPath + "\scala-excludes.txt"
            Dim sb As New StringBuilder()

            sb.AppendLine("EXCLUDES TAB")
            sb.AppendLine()
            sb.AppendLine("Exclude Files")
            sb.AppendLine()
            sb.AppendLine()
            If ListBox3.Items.Count > 0 Then
                For Each item In ListBox3.Items
                    sb.AppendLine(item.ToString())
                Next
            End If
            Using outfile As New StreamWriter(myPath)
                outfile.Write(sb.ToString())
            End Using
            MsgBox("Write Successful. File scala-excludes.txt was written in your current logs directory.", MsgBoxStyle.OkOnly, "Write Successful")

            'ElseIf activeTab = 4 Then
            '    'export data in tab#4
            '    Dim myPath As String = currPath + "\scala-search_results.txt"
            '    Dim sb As New StringBuilder()

            '    sb.AppendLine("SEARCH TAB")
            '    sb.AppendLine()
            '    sb.AppendLine("Search String:")
            '    sb.Append(TextBox5.Text)
            '    sb.AppendLine()
            '    sb.AppendLine()
            '    sb.AppendLine("Results:")
            '    sb.AppendLine()
            '    If ListBox2.Items.Count > 0 Then
            '        For Each item In ListBox2.Items
            '            sb.AppendLine(item.ToString())
            '        Next
            '    End If
            '    Using outfile As New StreamWriter(myPath)
            '        outfile.Write(sb.ToString())
            '    End Using
            '    MsgBox("Write Successful. File scala-search_results.txt was written in your current logs directory.", MsgBoxStyle.OkOnly, "Write Successful")

        ElseIf activeTab = 6 Then
            'export data in tab#6
            Dim myPath As String = currPath + "\scala-scan arguments.txt"
            Dim sb As New StringBuilder()

            sb.AppendLine("COMPARE ARGUMENTS TAB")
            sb.AppendLine()
            sb.AppendLine("BUILD LOG Arguments:")
            sb.AppendLine()
            sb.Append(TextBox12.Text)
            sb.AppendLine()
            If TextBox14.Text.Count > 1 Then
                sb.AppendLine("BUILD LOG 2 Arguments:")
                sb.AppendLine()
                sb.Append(TextBox14.Text)
                sb.AppendLine()
            End If
            sb.AppendLine()
            sb.AppendLine("SCAN LOG Arguments:")
            sb.AppendLine()
            sb.Append(TextBox11.Text)
            sb.AppendLine()
            If TextBox13.Text.Count > 1 Then
                sb.AppendLine("SCAN LOG 2 Arguments:")
                sb.AppendLine()
                sb.Append(TextBox13.Text)
                sb.AppendLine()
            End If

            Using outfile As New StreamWriter(myPath)
                outfile.Write(sb.ToString())
            End Using
            MsgBox("Write Successful. File scala-compare_arguments.txt was written in your current logs directory.", MsgBoxStyle.OkOnly, "Write Successful")

            End If
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'export data to a single file

        If (CheckFolderValidity(currPath) = 0) Then
            MsgBox("The Current Directory Does Not Contain Log Files. Check if you dragged the proper folder to the application", MsgBoxStyle.OkOnly, "ERROR")
            Exit Sub
        End If

        Dim myPath As String = currPath + "\scala-out.txt"
        Dim sb As New StringBuilder()

        sb.AppendLine("-----------------BASIC STATS TAB-----------------")
        sb.AppendLine()
        sb.Append(TextBox3.Text)
        sb.AppendLine()
        sb.Append(TextBox4.Text)
        sb.AppendLine()
        sb.Append(TextBox7.Text)
        sb.AppendLine()
        sb.Append(TextBox10.Text)
        sb.AppendLine()
        sb.AppendLine()
        sb.AppendLine("UNIQUE ERRORS")
        sb.AppendLine()
        sb.AppendLine("Build Log:")
        If ListView1.Items.Count > 0 Then
            For Each item In ListView1.Items
                sb.AppendLine(item.SubItems(1).Text & "  " & item.SubItems(0).Text)
            Next
        End If
        sb.AppendLine()
        sb.AppendLine("Scan Log:")
        If ListView2.Items.Count > 0 Then
            For Each item In ListView2.Items
                sb.AppendLine(item.SubItems(1).Text & "  " & item.SubItems(0).Text)
            Next
        End If
        sb.AppendLine()


        sb.AppendLine()
        sb.AppendLine()
        sb.AppendLine("-----------------MISSING DEPENDENCIES TAB-----------------")
        sb.AppendLine()
        sb.AppendLine("Missing Dependencies")
        sb.AppendLine()
        sb.AppendLine()
        If ListBox1.Items.Count > 0 Then
            For Each item In ListBox1.Items
                sb.AppendLine(item.ToString())
            Next
        End If
        sb.AppendLine()
        sb.AppendLine()
        sb.AppendLine("-----------------EXCLUDES TAB-----------------")
        sb.AppendLine()
        sb.AppendLine("Exclude Files")
        sb.AppendLine()
        sb.AppendLine()
        If ListBox3.Items.Count > 0 Then
            For Each item In ListBox3.Items
                sb.AppendLine(item.ToString())
            Next
        End If
        sb.AppendLine("COMPARE ARGUMENTS TAB")
        sb.AppendLine()
        sb.AppendLine("BUILD LOG Arguments:")
        sb.AppendLine()
        sb.Append(TextBox12.Text)
        sb.AppendLine()
        If TextBox14.Text.Count > 1 Then
            sb.AppendLine("BUILD LOG 2 Arguments:")
            sb.AppendLine()
            sb.Append(TextBox14.Text)
            sb.AppendLine()
        End If
        sb.AppendLine()
        sb.AppendLine("SCAN LOG Arguments:")
        sb.AppendLine()
        sb.Append(TextBox11.Text)
        sb.AppendLine()
        If TextBox13.Text.Count > 1 Then
            sb.AppendLine("SCAN LOG 2 Arguments:")
            sb.AppendLine()
            sb.Append(TextBox13.Text)
            sb.AppendLine()
        End If

        Using outfile As New StreamWriter(myPath)
            outfile.Write(sb.ToString())
        End Using
        MsgBox("Write Successful. File scala-out.txt was written in your current logs directory.", MsgBoxStyle.OkOnly, "Write Successful")

    End Sub

    '----------GENERATE RANDOM QUOTE ON START UP----------'
    Private Sub randomQuote()
        Dim quotes As New ArrayList()
        quotes.Add("need a light?")
        quotes.Add("laser")
        quotes.Add("feel, don't think, use your instincts")
        quotes.Add("what does the fox say?")
        quotes.Add("beam me up, scotty")
        quotes.Add("lend me some sugar, i am your neighbor")
        quotes.Add("yeahh.. that'd be great")
        quotes.Add("happy mondays")
        quotes.Add("holy guacamole batman!")
        quotes.Add("it's you and me, i know it's my destiny")
        quotes.Add("a martini. shaken, not stirred")
        quotes.Add("rock-paper-scissors-lizard-spock")
        quotes.Add("houston, we have a problem")
        quotes.Add("i don't wanna close my eyes, i don't wanna fall asleep")
        quotes.Add("oohh it's godzilla!")
        quotes.Add("i wanna be the very best, like no one ever was")
        quotes.Add("ain't nobody got time for that!")
        quotes.Add("leeerooyy jenkinsss")
        quotes.Add("thank you captain obvious")
        quotes.Add("i don't always drink beer, but when i do, i prefer dos equis")
        quotes.Add("shot to the heart, and you're to blame")
        quotes.Add("why did the chicken cross the road?")
        quotes.Add("aliens!")
        quotes.Add("rock a bye baby")
        quotes.Add("127.0.0.1 sweet 127.0.0.1")
        quotes.Add("42: the answer to the meaning of life")
        quotes.Add("1403")
        quotes.Add("hello world")
        quotes.Add("to catch them is my real test, to train them is my cause..")
        Dim x As Integer = quotes.Count
        Randomize()
        ToolStripStatusLabel3.Text = quotes(CInt(Math.Ceiling(Rnd() * x)) - 1)
    End Sub

    '----------LOAD FORM1----------'
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call randomQuote()
    End Sub



End Class
