

Public Class CheckedComboBox
    Inherits ComboBox
    Private _dropdown As Dropdown
    Private _CheckedMember As String
    Private _Int As String
    Private _dropdownWidth As Integer
    Private _RowCount As Integer
    Public Event CheckedChanging As CheckedChangingEventHandler
    Protected Sub onCheckedChanging(ByVal e As CheckedChangingArgs)
        If CheckedChangingEvent IsNot Nothing Then
            RaiseEvent CheckedChanging(Me, e)
        End If
    End Sub
    Friend Shadows Class Dropdown
        Inherits Form
        Private ccbParent As CheckedComboBox
        Private WithEvents cclb As DataGridView
        Private dropdownClosed As Boolean = True


        Public Sub New(ByVal Parent As CheckedComboBox)
            ccbParent = Parent
            InitializeComponent()

            Me.ShowInTaskbar = False
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.ControlBox = False
            Me.Text = ""
        End Sub
        Private Sub InitializeComponent()
            Me.cclb = New DataGridView
            Me.SuspendLayout()
            ' 
            ' cclb
            ' 
            Me.cclb.BorderStyle = System.Windows.Forms.BorderStyle.None
            Me.cclb.Dock = System.Windows.Forms.DockStyle.Fill
            Me.cclb.Location = New System.Drawing.Point(0, 0)
            Me.cclb.Name = "cclb"
            Me.cclb.Size = New System.Drawing.Size(47, 15)
            Me.cclb.RightToLeft = Me.RightToLeft
            Me.cclb.TabIndex = 0
            ' 
            ' Dropdown

            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.BackColor = System.Drawing.SystemColors.Menu
            Me.ClientSize = New System.Drawing.Size(47, 16)
            Me.ControlBox = False
            Me.KeyPreview = True
            Me.Controls.Add(Me.cclb)
            Me.ForeColor = System.Drawing.SystemColors.ControlText
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
            Me.MinimizeBox = False
            Me.Name = "ccbParent"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
            Me.ResumeLayout(False)
        End Sub
        Public Sub CloseDropdown(ByVal enactChanges As Boolean)
            If dropdownClosed Then
                Return
            End If
            Debug.WriteLine("CloseDropdown")
            ' Perform the actual selection and display of checked items.
            If enactChanges Then
                ccbParent.SelectedIndex = -1
                ' Set the text portion equal to the string comprising all checked items (if any, otherwise empty!).

            Else
                ' Caller cancelled selection - need to restore the checked items to their original state.
                Dim i As Integer = 0
                'Do While i < cclb.Items.Count
                '    cclb.SetItemChecked(i, checkedStateArr(i))
                '    i += 1
                'Loop
            End If
            ccbParent.Text = GetCheckedItemsStringValue()

            ' From now on the dropdown is considered closed. We set the flag here to prevent OnDeactivate() calling
            ' this method once again after hiding this window.
            dropdownClosed = True
            ' Set the focus to our parent CheckedComboBox and hide the dropdown check list.
            ccbParent.Focus()
            Me.Hide()
            ' Notify CheckedComboBox that its dropdown is closed. (NOTE: it does not matter which parameters we pass to
            ' OnDropDownClosed() as long as the argument is CCBoxEventArgs so that the method knows the notification has
            ' come from our code and not from the framework).
            ccbParent.OnDropDownClosed(Nothing)
        End Sub
        Protected Overrides Sub OnActivated(ByVal e As EventArgs)
            MyBase.OnActivated(e)
            ' ccbParent.Text = GetCheckedItemsStringValue()
            dropdownClosed = False
            'cclb.Focus()

        End Sub
        Protected Overrides Sub OnDeactivate(ByVal e As EventArgs)
            MyBase.OnDeactivate(e)
            CloseDropdown(True)
        End Sub
        Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
            MyBase.OnKeyDown(e)
            'If e.KeyCode = Keys.Space Then
            '    If cclb.CurrentRow IsNot Nothing Then
            '        If cclb.CurrentRow.RowType = Janus.Windows.GridEX.RowType.Record Then


            '            Dim dr As DataRow = CType(cclb.CurrentRow.DataRow, DataRowView).Row
            '            If cclb.CurrentRow.RowType = Janus.Windows.GridEX.RowType.Record Then
            '                Dim V As Boolean = nPnkrec.Data.Common.NVL(dr(ccbParent._CheckedMember), False)
            '                dr.BeginEdit()
            '                dr(ccbParent._CheckedMember) = Not V
            '                ccbParent.Text = GetCheckedItemsStringValue()
            '            End If
            '        End If
            '    End If
            'End If


        End Sub
        ReadOnly Property List As DataGridView
            Get
                Return cclb
            End Get
        End Property

        Public Function GetCheckedItemsStringValue() As String
            Dim strValue As String = ""
            Dim DT As DataTable = CType(cclb.DataSource, DataTable)

            If cclb.DataSource Is Nothing Then Return ""
            'For Each dr As DataRow In CType(cclb.DataSource, DataTable).Rows
            '    dr(ccbParent._CheckedMember) = False
            'Next
            'For Each jdr As Janus.Windows.GridEX.GridEXRow In cclb.GetCheckedRows
            '    Dim drinfo As DataRow = CType(jdr.DataRow, DataRowView).Row

            '    Dim drs As DataRow() = DT.Select(ccbParent.ValueMember & "=" & nPnkrec.Data.Common.NVL(drinfo(ccbParent.ValueMember), 0))
            '    If drs.Length > 0 Then
            '        drs(0)(ccbParent._CheckedMember) = True
            '    End If
            'Next
            For Each dr As DataRow In DT.Rows
                If IsDBNull(dr(ccbParent._CheckedMember)) Then
                    strValue &= IIf(strValue = "", "", ",") & dr(ccbParent.DisplayMember)
                End If
            Next
            Return strValue
        End Function
        Public Function GetCheckedValues() As String
            Dim strValue As String = ""
            'cclb.UpdateData()
            'For Each dvgr As Janus.Windows.GridEX.GridEXRow In cclb.GetRows
            '    If dvgr.Cells(ccbParent._CheckedMember).Value Then
            '        strValue &= IIf(strValue = "", "", ",") & dvgr.Cells(ccbParent.ValueMember).Value
            '    End If
            'Next
            If cclb.DataSource Is Nothing Then Return ""
            For Each dr As DataRow In CType(cclb.DataSource, DataTable).Rows
                '   If nPnkrec.Data.Common.NVL(dr(ccbParent._CheckedMember), False) Then
                strValue &= IIf(strValue = "", "", ",") & dr(ccbParent.ValueMember)
                'End If

            Next
            Return strValue
        End Function
        Public Function GetCheckedRows() As DataRow()
            Dim strValue As DataRow() = Nothing
            Dim i As Long = 0
            'cclb.UpdateData()
            'For Each dvgr As Janus.Windows.GridEX.GridEXRow In cclb.GetRows
            '    If nPnkrec.Data.Common.NVL(dvgr.Cells(ccbParent._CheckedMember).Value, False) Then
            '        ReDim Preserve strValue(i)
            '        strValue(i) = CType(dvgr.DataRow, DataRowView).Row
            '        i += 1
            '    End If
            'Next
            Return strValue
        End Function

        'Private Sub cclb_CellEdited(ByVal sender As Object, ByVal e As Janus.Windows.GridEX.ColumnActionEventArgs) Handles cclb.CellEdited
        '    Dim x = 1
        'End Sub

        'Private Sub cclb_CellValueChanged(ByVal sender As Object, ByVal e As Janus.Windows.GridEX.ColumnActionEventArgs) Handles cclb.CellValueChanged
        '    ccbParent.FindForm.Text &= cclb.CurrentRow.Cells(0).Value
        '    If cclb.CurrentRow IsNot Nothing Then
        '        If cclb.CurrentRow.RowType = Janus.Windows.GridEX.RowType.Record Then
        '            Dim dr As DataRow = CType(cclb.CurrentRow.DataRow, DataRowView).Row
        '            Dim V As Boolean = nPnkrec.Data.Common.NVL(cclb.CurrentRow.Cells(ccbParent._CheckedMember).Value, False)
        '            dr.BeginEdit()
        '            dr(ccbParent._CheckedMember) = V
        '            ccbParent.Text = GetCheckedItemsStringValue()
        '            ccbParent.onCheckedChanging(New CheckedChangingArgs(e.Column.Index, cclb.CurrentRow.RowIndex, V, dr))
        '        End If

        '    End If
        '    ccbParent.Text = GetCheckedItemsStringValue()
        'End Sub
        'Private Sub cclb_CellEdited(ByVal sender As Object, ByVal e As Janus.Windows.GridEX.ColumnActionEventArgs) Handles cclb.CellEdited
        '    ccbParent.Text = GetCheckedItemsStringValue()
        'End Sub

        'Private Sub cclb_CellUpdated(ByVal sender As Object, ByVal e As Janus.Windows.GridEX.ColumnActionEventArgs) Handles cclb.CellUpdated
        '    ccbParent.Text = GetCheckedItemsStringValue()
        '    'cclb.Focus()
        'End Sub

        'Private Sub cclb_CellValueChanged(ByVal sender As Object, ByVal e As Janus.Windows.GridEX.ColumnActionEventArgs) Handles cclb.CellValueChanged
        '    ccbParent.Text = GetCheckedItemsStringValue()
        '    'cclb.Focus()
        'End Sub

        Private Sub cclb_ChangeUICues(ByVal sender As Object, ByVal e As System.Windows.Forms.UICuesEventArgs) Handles cclb.ChangeUICues

        End Sub

        Private Sub cclb_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cclb.Click
            'If cclb.CurrentRow IsNot Nothing Then
            '    If cclb.CurrentRow.RowType = Janus.Windows.GridEX.RowType.Record Then
            '        Dim dr As DataRow = CType(cclb.CurrentRow.DataRow, DataRowView).Row
            '           Dim V As Boolean = nPnkrec.Data.Common.NVL(dr(ccbParent._CheckedMember), False)
            '            dr.BeginEdit()
            '            dr(ccbParent._CheckedMember) = Not V
            '            ccbParent.Text = GetCheckedItemsStringValue()
            '        End If

            'End If

        End Sub

        'Private Sub cclb_ColumnHeaderClick(ByVal sender As Object, ByVal e As Janus.Windows.GridEX.ColumnActionEventArgs) Handles cclb.ColumnHeaderClick

        'End Sub

        'Private Sub cclb_EditingCell(ByVal sender As Object, ByVal e As Janus.Windows.GridEX.EditingCellEventArgs) Handles cclb.EditingCell
        '    ccbParent.Text = GetCheckedItemsStringValue()

        '    ' cclb.Focus()
        'End Sub

        Private Sub cclb_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cclb.KeyDown

            If e.KeyCode = Keys.Escape Then
                CloseDropdown(False)
                Return
            End If
            If e.KeyCode = Keys.Space Then
                If cclb.CurrentRow IsNot Nothing Then

                    'If cclb.CurrentRow.RowType = Janus.Windows.GridEX.RowType.Record Then
                    '    Dim dr As DataRow = CType(cclb.CurrentRow.DataRow, DataRowView).Row
                    '    Dim V As Boolean = nPnkrec.Data.Common.NVL(dr(ccbParent._CheckedMember), False)
                    '    dr.BeginEdit()
                    '    dr(ccbParent._CheckedMember) = Not V
                    '    ccbParent.Text = GetCheckedItemsStringValue()
                    'End If
                End If

            End If
        End Sub


    End Class
    'Friend Class DataGridView
    '    Inherits System.Windows.Forms.DataGridView
    '    Private vInitColumns As String = ""
    '    Private ColumnArray As Array
    '    Public Sub New()

    '        '  tlbmrtPrint.Visible = False
    '        'If nPnKrec.Domain.Globals.CurrentOrganization.Alias <> Domain.DomainValues.X_ORGANIZATION_TYPE.Krec_Mashhad Or _
    '        '   nPnKrec.Domain.Globals.CurrentOrganization.Alias <> Domain.DomainValues.X_ORGANIZATION_TYPE.BistonPowerPlant Or _
    '        '   nPnKrec.Domain.Globals.CurrentOrganization.Alias <> Domain.DomainValues.X_ORGANIZATION_TYPE.QazvinPowerplant Then
    '        '  btnFullFilter.Visible = True
    '        'End If
    '        'btnChart.Visible = False
    '    End Sub

    '    Public Sub Initialize(ByVal GridColumns As String, Optional ByVal DataSource_ As Object = Nothing)

    '    End Sub

    '    Private Sub InitializeComponent()
    '        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
    '        Me.SuspendLayout()
    '        '
    '        'DataGridView
    '        '
    '        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
    '        Me.ResumeLayout(False)
    '    End Sub
    'End Class
    Public Sub New()
        MyBase.New()
        ' We want to do the drawing of the dropdown.
        Me.DrawMode = DrawMode.OwnerDrawVariable
        ' Default value separator.
        'Me.valueSeparator_Renamed = ", "
        ' This prevents the actual ComboBox dropdown to show, although it's not strickly-speaking necessary.
        ' But including this remove a slight flickering just before our dropdown appears (which is caused by
        ' the empty-dropdown list of the ComboBox which is displayed for fractions of a second).
        Me.DropDownHeight = 1
        ' This is the default setting - text portion is editable and user must click the arrow button
        ' to see the list portion. Although we don't want to allow the user to edit the text portion
        ' the DropDownList style is not being used because for some reason it wouldn't allow the text
        ' portion to be programmatically set. Hence we set it as editable but disable keyboard input (see below).
        Me.DropDownStyle = ComboBoxStyle.DropDown
        _dropdown = New Dropdown(Me)
        _dropdownWidth = Me.Width
        _dropdown.RightToLeft = Me.RightToLeft
        ' CheckOnClick style for the dropdown (NOTE: must be set after dropdown is created).
        'Me.CheckOnClick = True
        '_SearchForm = New SearchForm(Me)
    End Sub
    Protected Overrides Sub OnDropDown(ByVal e As EventArgs)
        MyBase.OnDropDown(e)
        DoDropDown()
        DoDropDown()
    End Sub

    Private Sub DoDropDown()
        If (Not _dropdown.Visible) Then
            Dim lt As Point = Parent.PointToScreen(New Point(Left, Top))
            Dim P As Point = GetDropDownPosition(_dropdown) + New Point(0, lt.Y + Me.Height)
            Dim rect As Rectangle = RectangleToScreen(Me.ClientRectangle)
            Dim scr As Rectangle = Screen.FromControl(Me).Bounds
            If P.Y > scr.Height Then
                P.Y = lt.Y - _dropdown.Height '- Me.Height / 2
            End If
            _dropdown.Location = P
            _dropdown.RightToLeft = Me.RightToLeft
            Dim ItemHeight As Integer = 25

            Dim count As Integer = _RowCount
            If count > Me.MaxDropDownItems Then
                count = Me.MaxDropDownItems
            ElseIf count = 0 Then
                count = 1
            End If
            GetWidth()
            _dropdown.Size = New Size(_dropdownWidth, (ItemHeight) * count)
            _dropdown.Show(Me)
        End If
    End Sub

    Protected Overrides Sub OnDropDownClosed(ByVal e As EventArgs)
        ' Call the handlers for this event only if the call comes from our code - NOT the framework's!
        ' NOTE: that is because the events were being fired in a wrong order, due to the actual dropdown list
        '       of the ComboBox which lies underneath our dropdown and gets involved every time.
        'If TypeOf e Is Dropdown.CCBoxEventArgs Then
        MyBase.OnDropDownClosed(e)
        'End If
    End Sub
    Public Sub Initialize(ByVal vDataTable As DataTable, ByVal vValueMember As String, ByVal vDisplayMember As String, ByVal vCheckedMember As String, ByVal InitGridString As String)
        If vDataTable Is Nothing Then Return
        _dropdown.List.Font = Me.Parent.Font
        Me.DisplayMember = vDisplayMember
        Me.ValueMember = vValueMember
        _CheckedMember = vCheckedMember
        Me.DataSource = vDataTable
        _RowCount = vDataTable.Rows.Count
        Dim strInitGridString As String
        If _CheckedMember <> "" Then
            strInitGridString = "" & ":" & _CheckedMember & ":30;" & InitGridString
        Else
            strInitGridString = InitGridString
        End If
        _Int = strInitGridString
        GetWidth()
        ' _dropdown.List.Initialize(strInitGridString)
        _dropdown.List.DataSource = vDataTable

       
    End Sub
    Private Sub GetWidth()
        'Dim sColumns As String() = _Int.Split(";")
        '_dropdownWidth = 0
        'For Each s As String In sColumns
        '    Dim sWidths As String() = s.Split(":")
        '    _dropdownWidth += sWidths(2)

        'Next
        'If _dropdownWidth < Me.Width Then
        '    _dropdownWidth = Me.Width
        'End If
        _dropdownWidth += 20

    End Sub
    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        MyBase.OnKeyPress(e)
        e.KeyChar = ""
        e.Handled = True
    End Sub
    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)
        If e.KeyCode = Keys.Down Then
            DoDropDown()
        End If
        e.Handled = (Not e.Alt) AndAlso Not (e.KeyCode = Keys.Tab) AndAlso Not ((e.KeyCode = Keys.Left) OrElse (e.KeyCode = Keys.Right) OrElse (e.KeyCode = Keys.Home) OrElse (e.KeyCode = Keys.End))
    End Sub
    Private Function GetDropDownPosition(ByVal dropDown As Dropdown) As Point
        Dim lt As Point = Parent.PointToScreen(New Point(Left, Top))
        Dim rb As Point = Parent.PointToScreen(New Point(Right, Bottom))
        Dim scr As Rectangle = Screen.FromControl(Me).Bounds
        Dim point As Point = New Point()
        If (((lt.X + dropDown.Width) > (scr.X + scr.Width)) AndAlso ((rb.X - dropDown.Width) >= scr.X)) Then
            point.X = rb.X - dropDown.Width
            If ((point.X + dropDown.Width) > (scr.X + scr.Width)) Then
                point.X = ((scr.X + scr.Width) - dropDown.Width)
            End If
        Else
            point.X = lt.X
            If (point.X < scr.X) Then
                point.X = scr.X
            End If
        End If


        If (((rb.Y + dropDown.Height) > (scr.Y + scr.Height)) AndAlso ((lt.Y - dropDown.Height) >= scr.Y)) Then
            point.Y = lt.Y - dropDown.Height
            If (point.Y < scr.Y) Then
                point.Y = scr.Y
            Else
                point.Y = rb.Y
                If ((point.Y + dropDown.Height) > (scr.Y + scr.Height)) Then
                    point.Y = ((scr.Y + scr.Height) - dropDown.Height)
                End If
            End If
        End If
        Return point
    End Function

End Class
Public Delegate Sub CheckedChangingEventHandler(ByVal sender As Object, ByVal e As CheckedChangingArgs)

Public Class CheckedChangingArgs
    Inherits EventArgs
    Private _ColumnIndex As Integer = 0
    Private _RowIndex As Integer = 0
    Private _value As Object
    Private _Row As DataRow
    Public ReadOnly Property Row As DataRow
        Get
            Return _Row
        End Get
    End Property
    Public ReadOnly Property Value As Object
        Get
            Return _value
        End Get
    End Property
    Public ReadOnly Property ColumnIndex As Integer
        Get
            Return _ColumnIndex
        End Get
    End Property
    Public ReadOnly Property RowIndex As Integer
        Get
            Return _RowIndex
        End Get
    End Property



    Public Sub New(ByVal ColumnIndex As Integer, ByVal RowIndex As Integer, ByVal value As Object, ByVal row As DataRow)
        _ColumnIndex = ColumnIndex
        _RowIndex = RowIndex
        _value = value
        _Row = row

    End Sub
End Class
