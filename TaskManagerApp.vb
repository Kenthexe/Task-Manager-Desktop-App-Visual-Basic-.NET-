' A simple Task Manager Desktop App in Visual Basic (.NET Framework)
'
' Features:
'   - Add/Edit/Delete tasks
'   - Set deadlines
'   - Mark complete
'   - Save/load data locally (JSON file)
'   - OOP
'   - Task Priority (Low, Medium, High)
'   - Task Category/Tag
'   - Search/filter tasks by title, status, priority, or category
'   - Sort tasks by deadline, priority, or completion
'   - Overdue task highlighting
'   - Export tasks to CSV
'   - Bulk delete completed tasks
'   - Show task counts (total, completed, overdue)
'   - Confirmation dialogs for destructive actions

Imports System.IO
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports System.Linq

' Task class (OOP)
Public Enum TaskPriority
    Low
    Medium
    High
End Enum

Public Class TaskItem
    Public Property Title As String
    Public Property Description As String
    Public Property Deadline As DateTime
    Public Property IsCompleted As Boolean
    Public Property Priority As TaskPriority
    Public Property Category As String

    Public Sub New()
    End Sub

    Public Sub New(title As String, desc As String, deadline As DateTime, Optional completed As Boolean = False, Optional priority As TaskPriority = TaskPriority.Medium, Optional category As String = "")
        Me.Title = title
        Me.Description = desc
        Me.Deadline = deadline
        Me.IsCompleted = completed
        Me.Priority = priority
        Me.Category = category
    End Sub

    Public Overrides Function ToString() As String
        Dim status As String = If(IsCompleted, "[✓]", If(DateTime.Now > Deadline AndAlso Not IsCompleted, "[!]", "[ ]"))
        Dim prio As String = $"[{Priority.ToString().Substring(0, 1)}]"
        Dim cat As String = If(String.IsNullOrWhiteSpace(Category), "", $"#{Category} ")
        Return $"{status} {prio} {cat}{Title} - Due: {Deadline.ToShortDateString()}"
    End Function
End Class

' Main Form
Public Class TaskManagerForm
    Inherits Form

    Private tasks As New List(Of TaskItem)
    Private Const DATA_FILE As String = "tasks.json"

    ' UI Controls
    Private WithEvents lstTasks As New ListBox()
    Private WithEvents btnAdd As New Button()
    Private WithEvents btnEdit As New Button()
    Private WithEvents btnDelete As New Button()
    Private WithEvents btnMarkComplete As New Button()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnLoad As New Button()
    Private WithEvents btnExportCSV As New Button()
    Private WithEvents btnBulkDeleteCompleted As New Button()
    Private WithEvents cmbSort As New ComboBox()
    Private WithEvents txtSearch As New TextBox()
    Private WithEvents cmbFilterStatus As New ComboBox()
    Private WithEvents cmbFilterPriority As New ComboBox()
    Private WithEvents cmbFilterCategory As New ComboBox()
    Private lblCounts As New Label()

    Private filteredTasks As New List(Of TaskItem)

    Public Sub New()
        Me.Text = "Task Manager"
        Me.Size = New Drawing.Size(800, 500)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        lstTasks.Location = New Drawing.Point(20, 60)
        lstTasks.Size = New Drawing.Size(500, 350)
        lstTasks.Font = New Drawing.Font("Segoe UI", 10)
        Me.Controls.Add(lstTasks)

        btnAdd.Text = "Add"
        btnAdd.Location = New Drawing.Point(550, 60)
        btnAdd.Size = New Drawing.Size(100, 30)
        Me.Controls.Add(btnAdd)

        btnEdit.Text = "Edit"
        btnEdit.Location = New Drawing.Point(660, 60)
        btnEdit.Size = New Drawing.Size(100, 30)
        Me.Controls.Add(btnEdit)

        btnDelete.Text = "Delete"
        btnDelete.Location = New Drawing.Point(550, 100)
        btnDelete.Size = New Drawing.Size(100, 30)
        Me.Controls.Add(btnDelete)

        btnMarkComplete.Text = "Mark Complete"
        btnMarkComplete.Location = New Drawing.Point(660, 100)
        btnMarkComplete.Size = New Drawing.Size(100, 30)
        Me.Controls.Add(btnMarkComplete)

        btnSave.Text = "Save"
        btnSave.Location = New Drawing.Point(550, 140)
        btnSave.Size = New Drawing.Size(100, 30)
        Me.Controls.Add(btnSave)

        btnLoad.Text = "Load"
        btnLoad.Location = New Drawing.Point(660, 140)
        btnLoad.Size = New Drawing.Size(100, 30)
        Me.Controls.Add(btnLoad)

        btnExportCSV.Text = "Export CSV"
        btnExportCSV.Location = New Drawing.Point(550, 180)
        btnExportCSV.Size = New Drawing.Size(100, 30)
        Me.Controls.Add(btnExportCSV)

        btnBulkDeleteCompleted.Text = "Delete Completed"
        btnBulkDeleteCompleted.Location = New Drawing.Point(660, 180)
        btnBulkDeleteCompleted.Size = New Drawing.Size(100, 30)
        Me.Controls.Add(btnBulkDeleteCompleted)

        cmbSort.Location = New Drawing.Point(550, 230)
        cmbSort.Size = New Drawing.Size(210, 25)
        cmbSort.Items.AddRange({"Sort: Deadline", "Sort: Priority", "Sort: Completed"})
        cmbSort.SelectedIndex = 0
        Me.Controls.Add(cmbSort)

        txtSearch.Location = New Drawing.Point(20, 20)
        txtSearch.Size = New Drawing.Size(200, 25)
        txtSearch.PlaceholderText = "Search by title..."
        Me.Controls.Add(txtSearch)

        cmbFilterStatus.Location = New Drawing.Point(230, 20)
        cmbFilterStatus.Size = New Drawing.Size(100, 25)
        cmbFilterStatus.Items.AddRange({"All", "Active", "Completed", "Overdue"})
        cmbFilterStatus.SelectedIndex = 0
        Me.Controls.Add(cmbFilterStatus)

        cmbFilterPriority.Location = New Drawing.Point(340, 20)
        cmbFilterPriority.Size = New Drawing.Size(100, 25)
        cmbFilterPriority.Items.AddRange({"All", "Low", "Medium", "High"})
        cmbFilterPriority.SelectedIndex = 0
        Me.Controls.Add(cmbFilterPriority)

        cmbFilterCategory.Location = New Drawing.Point(450, 20)
        cmbFilterCategory.Size = New Drawing.Size(100, 25)
        cmbFilterCategory.Items.Add("All")
        cmbFilterCategory.SelectedIndex = 0
        Me.Controls.Add(cmbFilterCategory)

        lblCounts.Location = New Drawing.Point(20, 420)
        lblCounts.Size = New Drawing.Size(600, 30)
        lblCounts.Font = New Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        Me.Controls.Add(lblCounts)

        LoadTasks()
        UpdateCategoryFilter()
        ApplyFiltersAndSort()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim dlg As New TaskDialog()
        If dlg.ShowDialog() = DialogResult.OK Then
            tasks.Add(dlg.Task)
            UpdateCategoryFilter()
            ApplyFiltersAndSort()
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If lstTasks.SelectedIndex >= 0 Then
            Dim t As TaskItem = filteredTasks(lstTasks.SelectedIndex)
            Dim dlg As New TaskDialog(t)
            If dlg.ShowDialog() = DialogResult.OK Then
                Dim idx As Integer = tasks.IndexOf(t)
                If idx >= 0 Then
                    tasks(idx) = dlg.Task
                End If
                UpdateCategoryFilter()
                ApplyFiltersAndSort()
            End If
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If lstTasks.SelectedIndex >= 0 Then
            Dim t As TaskItem = filteredTasks(lstTasks.SelectedIndex)
            If MessageBox.Show("Delete selected task?", "Confirm", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                tasks.Remove(t)
                UpdateCategoryFilter()
                ApplyFiltersAndSort()
            End If
        End If
    End Sub

    Private Sub btnMarkComplete_Click(sender As Object, e As EventArgs) Handles btnMarkComplete.Click
        If lstTasks.SelectedIndex >= 0 Then
            Dim t As TaskItem = filteredTasks(lstTasks.SelectedIndex)
            t.IsCompleted = True
            ApplyFiltersAndSort()
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SaveTasks()
        MessageBox.Show("Tasks saved!", "Info")
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        LoadTasks()
        UpdateCategoryFilter()
        ApplyFiltersAndSort()
        MessageBox.Show("Tasks loaded!", "Info")
    End Sub

    Private Sub btnExportCSV_Click(sender As Object, e As EventArgs) Handles btnExportCSV.Click
        Dim sfd As New SaveFileDialog() With {.Filter = "CSV Files|*.csv", .FileName = "tasks.csv"}
        If sfd.ShowDialog() = DialogResult.OK Then
            Try
                Using sw As New StreamWriter(sfd.FileName)
                    sw.WriteLine("Title,Description,Deadline,Completed,Priority,Category")
                    For Each t In tasks
                        sw.WriteLine($"""{t.Title.Replace("""", """""")}"",""{t.Description.Replace("""", """""")}"",""{t.Deadline:yyyy-MM-dd}"",{t.IsCompleted},{t.Priority},{t.Category}")
                    Next
                End Using
                MessageBox.Show("Exported to CSV!", "Info")
            Catch ex As Exception
                MessageBox.Show("Error exporting: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub btnBulkDeleteCompleted_Click(sender As Object, e As EventArgs) Handles btnBulkDeleteCompleted.Click
        If MessageBox.Show("Delete ALL completed tasks?", "Confirm", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            tasks = tasks.Where(Function(t) Not t.IsCompleted).ToList()
            UpdateCategoryFilter()
            ApplyFiltersAndSort()
        End If
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        ApplyFiltersAndSort()
    End Sub

    Private Sub cmbFilterStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFilterStatus.SelectedIndexChanged
        ApplyFiltersAndSort()
    End Sub

    Private Sub cmbFilterPriority_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFilterPriority.SelectedIndexChanged
        ApplyFiltersAndSort()
    End Sub

    Private Sub cmbFilterCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFilterCategory.SelectedIndexChanged
        ApplyFiltersAndSort()
    End Sub

    Private Sub cmbSort_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSort.SelectedIndexChanged
        ApplyFiltersAndSort()
    End Sub

    Private Sub ApplyFiltersAndSort()
        Dim q = tasks.AsEnumerable()

        ' Search
        If Not String.IsNullOrWhiteSpace(txtSearch.Text) Then
            q = q.Where(Function(t) t.Title.IndexOf(txtSearch.Text, StringComparison.OrdinalIgnoreCase) >= 0)
        End If

        ' Status filter
        Select Case cmbFilterStatus.SelectedItem?.ToString()
            Case "Active"
                q = q.Where(Function(t) Not t.IsCompleted AndAlso t.Deadline >= DateTime.Today)
            Case "Completed"
                q = q.Where(Function(t) t.IsCompleted)
            Case "Overdue"
                q = q.Where(Function(t) Not t.IsCompleted AndAlso t.Deadline < DateTime.Today)
        End Select

        ' Priority filter
        Select Case cmbFilterPriority.SelectedItem?.ToString()
            Case "Low"
                q = q.Where(Function(t) t.Priority = TaskPriority.Low)
            Case "Medium"
                q = q.Where(Function(t) t.Priority = TaskPriority.Medium)
            Case "High"
                q = q.Where(Function(t) t.Priority = TaskPriority.High)
        End Select

        ' Category filter
        If cmbFilterCategory.SelectedIndex > 0 Then
            Dim cat = cmbFilterCategory.SelectedItem.ToString()
            q = q.Where(Function(t) t.Category = cat)
        End If

        ' Sorting
        Select Case cmbSort.SelectedIndex
            Case 0 ' Deadline
                q = q.OrderBy(Function(t) t.Deadline)
            Case 1 ' Priority
                q = q.OrderByDescending(Function(t) t.Priority)
            Case 2 ' Completed
                q = q.OrderBy(Function(t) t.IsCompleted)
        End Select

        filteredTasks = q.ToList()
        RefreshTaskList()
        UpdateCounts()
    End Sub

    Private Sub RefreshTaskList()
        lstTasks.Items.Clear()
        For Each t In filteredTasks
            Dim display = t.ToString()
            lstTasks.Items.Add(display)
        Next
        ' Overdue highlighting
        For i = 0 To filteredTasks.Count - 1
            If Not filteredTasks(i).IsCompleted AndAlso filteredTasks(i).Deadline < DateTime.Today Then
                lstTasks.Items(i) = "[!OVERDUE!] " & lstTasks.Items(i)
            End If
        Next
    End Sub

    Private Sub UpdateCounts()
        Dim total = tasks.Count
        Dim completed = tasks.Count(Function(t) t.IsCompleted)
        Dim overdue = tasks.Count(Function(t) Not t.IsCompleted AndAlso t.Deadline < DateTime.Today)
        lblCounts.Text = $"Total: {total}   Completed: {completed}   Overdue: {overdue}"
    End Sub

    Private Sub UpdateCategoryFilter()
        Dim cats = tasks.Select(Function(t) t.Category).Where(Function(c) Not String.IsNullOrWhiteSpace(c)).Distinct().OrderBy(Function(c) c).ToList()
        Dim sel = If(cmbFilterCategory.SelectedIndex >= 0, cmbFilterCategory.SelectedItem?.ToString(), "All")
        cmbFilterCategory.Items.Clear()
        cmbFilterCategory.Items.Add("All")
        For Each c In cats
            cmbFilterCategory.Items.Add(c)
        Next
        If sel IsNot Nothing AndAlso cmbFilterCategory.Items.Contains(sel) Then
            cmbFilterCategory.SelectedItem = sel
        Else
            cmbFilterCategory.SelectedIndex = 0
        End If
    End Sub

    Private Sub SaveTasks()
        Try
            Dim options As New JsonSerializerOptions With {.WriteIndented = True}
            Dim json As String = JsonSerializer.Serialize(tasks, options)
            File.WriteAllText(DATA_FILE, json)
        Catch ex As Exception
            MessageBox.Show("Error saving tasks: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadTasks()
        Try
            If File.Exists(DATA_FILE) Then
                Dim json As String = File.ReadAllText(DATA_FILE)
                tasks = JsonSerializer.Deserialize(Of List(Of TaskItem))(json)
            Else
                tasks = New List(Of TaskItem)()
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading tasks: " & ex.Message)
            tasks = New List(Of TaskItem)()
        End Try
    End Sub
End Class

' Dialog for Add/Edit Task
Public Class TaskDialog
    Inherits Form

    Public Property Task As TaskItem

    Private txtTitle As New TextBox()
    Private txtDesc As New TextBox()
    Private dtpDeadline As New DateTimePicker()
    Private chkCompleted As New CheckBox()
    Private cmbPriority As New ComboBox()
    Private txtCategory As New TextBox()
    Private btnOK As New Button()
    Private btnCancel As New Button()

    Public Sub New()
        Me.New(Nothing)
    End Sub

    Public Sub New(existingTask As TaskItem)
        Me.Text = If(existingTask Is Nothing, "Add Task", "Edit Task")
        Me.Size = New Drawing.Size(400, 370)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent

        Dim lblTitle As New Label() With {.Text = "Title:", .Location = New Drawing.Point(20, 20), .AutoSize = True}
        txtTitle.Location = New Drawing.Point(120, 20)
        txtTitle.Size = New Drawing.Size(240, 20)

        Dim lblDesc As New Label() With {.Text = "Description:", .Location = New Drawing.Point(20, 60), .AutoSize = True}
        txtDesc.Location = New Drawing.Point(120, 60)
        txtDesc.Size = New Drawing.Size(240, 60)
        txtDesc.Multiline = True

        Dim lblDeadline As New Label() With {.Text = "Deadline:", .Location = New Drawing.Point(20, 140), .AutoSize = True}
        dtpDeadline.Location = New Drawing.Point(120, 140)
        dtpDeadline.Size = New Drawing.Size(240, 20)
        dtpDeadline.Format = DateTimePickerFormat.Short

        Dim lblPriority As New Label() With {.Text = "Priority:", .Location = New Drawing.Point(20, 180), .AutoSize = True}
        cmbPriority.Location = New Drawing.Point(120, 180)
        cmbPriority.Size = New Drawing.Size(120, 25)
        cmbPriority.Items.AddRange({"Low", "Medium", "High"})
        cmbPriority.SelectedIndex = 1

        Dim lblCategory As New Label() With {.Text = "Category/Tag:", .Location = New Drawing.Point(20, 220), .AutoSize = True}
        txtCategory.Location = New Drawing.Point(120, 220)
        txtCategory.Size = New Drawing.Size(240, 20)

        chkCompleted.Text = "Completed"
        chkCompleted.Location = New Drawing.Point(120, 260)
        chkCompleted.AutoSize = True

        btnOK.Text = "OK"
        btnOK.Location = New Drawing.Point(80, 300)
        btnOK.Size = New Drawing.Size(100, 30)
        AddHandler btnOK.Click, AddressOf btnOK_Click

        btnCancel.Text = "Cancel"
        btnCancel.Location = New Drawing.Point(220, 300)
        btnCancel.Size = New Drawing.Size(100, 30)
        AddHandler btnCancel.Click, Sub() Me.DialogResult = DialogResult.Cancel

        Me.Controls.AddRange({lblTitle, txtTitle, lblDesc, txtDesc, lblDeadline, dtpDeadline, lblPriority, cmbPriority, lblCategory, txtCategory, chkCompleted, btnOK, btnCancel})

        If existingTask IsNot Nothing Then
            txtTitle.Text = existingTask.Title
            txtDesc.Text = existingTask.Description
            dtpDeadline.Value = existingTask.Deadline
            chkCompleted.Checked = existingTask.IsCompleted
            cmbPriority.SelectedIndex = CInt(existingTask.Priority)
            txtCategory.Text = existingTask.Category
            Task = existingTask
        End If
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtTitle.Text) Then
            MessageBox.Show("Title is required.")
            Return
        End If
        Dim prio As TaskPriority = CType(cmbPriority.SelectedIndex, TaskPriority)
        Task = New TaskItem(txtTitle.Text, txtDesc.Text, dtpDeadline.Value, chkCompleted.Checked, prio, txtCategory.Text.Trim())
        Me.DialogResult = DialogResult.OK
    End Sub
End Class

' Entry point
Module Program
    <STAThread>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New TaskManagerForm())
    End Sub
End Module
