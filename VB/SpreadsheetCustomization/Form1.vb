Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace SpreadsheetCustomization
	Partial Public Class Form1
		Inherits DevExpress.XtraBars.Ribbon.RibbonForm

		Public Shared categories() As String = { "Meat/Poultry", "Condiments", "Seafood", "Dairy Products", "Grains/Cereals", "Beverages", "Confections" }
		Private workbook As IWorkbook
		Private worksheet As Worksheet

		Private dateColumn As CellRange
		Private discountColumn As CellRange
		Private categoryColumn As CellRange

		Private dateEdit As DateEdit
		Private lookUpEdit As LookUpEdit
		Private checkBox As CheckEdit

		Public Sub New()
			InitializeComponent()
			workbook = spreadsheetControl1.Document
			workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.Xlsx)

			worksheet = workbook.Worksheets("Sales report")
			dateColumn = worksheet("Table[Order Date]")
			categoryColumn = worksheet("Table[Category]")
			discountColumn = worksheet("Table[Discount]")

			' Create custom controls to be displayed instead of the cell editor and specify their settings.
			lookUpEdit = CreateLookUp()
			checkBox = CreateCheckBox()
			dateEdit = CreateDateEdit()

			' Specify the SpreadsheetControl's options.
			spreadsheetControl1.Options.Behavior.Selection.AllowExtendSelection = False
			spreadsheetControl1.Options.Behavior.Drag = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled
			spreadsheetControl1.Options.VerticalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden
			spreadsheetControl1.Options.HorizontalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden

			AddHandler spreadsheetControl1.SelectionChanged, AddressOf spreadsheetControl1_SelectionChanged
			AddHandler spreadsheetControl1.CellBeginEdit, AddressOf spreadsheetControl1_CellBeginEdit
			AddHandler spreadsheetControl1.MouseWheel, AddressOf spreadsheetControl1_MouseWheel
		End Sub

		Private Sub spreadsheetControl1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
			Dim activeCellBounds As Rectangle = GetActiveCellBounds()
			UpdateActiveEditorBounds(activeCellBounds)
		End Sub

		Private Sub UpdateActiveEditorBounds(ByVal newBounds As Rectangle)
			If newBounds.IsEmpty Then
				HideAllEditors()
				Return
			End If

			If lookUpEdit.Visible Then
				lookUpEdit.Bounds = newBounds
			ElseIf checkBox.Visible Then
				checkBox.Bounds = newBounds
			ElseIf dateEdit.Visible Then
				dateEdit.Bounds = newBounds
			End If
		End Sub

		Private Sub spreadsheetControl1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
			HideAllEditors()
		End Sub

		Private Sub HideAllEditors()
			lookUpEdit.Visible = False
			lookUpEdit.Parent = Nothing

			checkBox.Visible = False
			checkBox.Parent = Nothing

			dateEdit.Visible = False
			dateEdit.Parent = Nothing

			spreadsheetControl1.Focus()
		End Sub

		Private Sub spreadsheetControl1_CellBeginEdit(ByVal sender As Object, ByVal e As DevExpress.XtraSpreadsheet.SpreadsheetCellCancelEventArgs)
			' Access the active cell.
			Dim activeCell As Cell = spreadsheetControl1.ActiveCell
			' Obtain the bounds of the active cell.
			Dim activeCellRect As Rectangle = GetActiveCellBounds()
			' If the active cell is out of the visible range, return. 
			If activeCellRect.IsEmpty Then
				e.Cancel = True
				Return
			End If

			' If the currently selected cell is in the "Category" column of the worksheet table, 
			' display the LookUpEdit control instead of the cell editor. 
			If CanShowLookUp(activeCell) Then
				e.Cancel = True
				ShowLookUp(activeCell, activeCellRect)

			' If the currently selected cell is in the "Discount" column of the worksheet table, 
			' display the CheckEdit control instead of the cell editor.
			ElseIf CanShowCheckBox(activeCell) Then
				e.Cancel = True
				ShowCheckEdit(activeCell, activeCellRect)

			' If the currently selected cell is in the "Order Date" column of the worksheet table, 
			' display the DateEdit control instead of the cell editor.
			ElseIf CanShowDateEdit(activeCell) Then
				e.Cancel = True
				ShowDateEdit(activeCell.Value, activeCellRect)
			End If
		End Sub
		Private Function GetActiveCellBounds() As Rectangle
			Dim activeCell As Cell = spreadsheetControl1.ActiveCell
			Return spreadsheetControl1.GetCellBounds(activeCell.RowIndex, activeCell.ColumnIndex)
		End Function
		Private Function CanShowLookUp(ByVal activeCell As Cell) As Boolean
			Return If(worksheet Is workbook.Worksheets.ActiveWorksheet, categoryColumn.IsIntersecting(activeCell), False)
		End Function
		Private Function CanShowCheckBox(ByVal activeCell As Cell) As Boolean
			Return If(worksheet Is workbook.Worksheets.ActiveWorksheet, discountColumn.IsIntersecting(activeCell), False)
		End Function
		Private Function CanShowDateEdit(ByVal activeCell As Cell) As Boolean
			Return If(worksheet Is workbook.Worksheets.ActiveWorksheet, dateColumn.IsIntersecting(activeCell), False)
		End Function

		#Region "LookUp"
		Private Function CreateLookUp() As LookUpEdit
			Dim cmbBox As New LookUpEdit()
			cmbBox.Properties.DataSource = categories
			cmbBox.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
			cmbBox.Visible = False
			cmbBox.Parent = spreadsheetControl1
			AddHandler cmbBox.KeyDown, AddressOf OnLookUpKeyDown
			Return cmbBox
		End Function

		Private Sub OnLookUpKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
			If e.KeyCode = Keys.Return Then
				Dim editValue As Object = lookUpEdit.EditValue
				If editValue Is Nothing Then
					Return
				End If

				' Assign a value of the currently selected item in the LookUpEdit to the active cell.  
				spreadsheetControl1.ActiveCell.Value = editValue.ToString()
				lookUpEdit.Visible = False
				lookUpEdit.Parent = Nothing
				spreadsheetControl1.Focus()
			ElseIf e.KeyCode = Keys.Escape Then
				lookUpEdit.Visible = False
				lookUpEdit.Parent = Nothing
				spreadsheetControl1.Focus()
			End If
		End Sub
		Private Sub ShowLookUp(ByVal cell As Cell, ByVal bounds As Rectangle)
			lookUpEdit.EditValue = cell.Value.TextValue

			UpdateLookUpAppearance(cell)

			lookUpEdit.Parent = spreadsheetControl1
			lookUpEdit.Bounds = bounds
			lookUpEdit.Visible = True
			lookUpEdit.Focus()
		End Sub
		Private Sub UpdateLookUpAppearance(ByVal source As Cell)
			lookUpEdit.BackColor = source.Fill.BackgroundColor
			Dim font As SpreadsheetFont = source.Font
			lookUpEdit.ForeColor = font.Color
			lookUpEdit.Font = New Font(font.Name, CSng(font.Size), GetFontStyle(font))
		End Sub
		Private Function GetFontStyle(ByVal font As SpreadsheetFont) As FontStyle
			Dim result As FontStyle = FontStyle.Regular
			If font.Bold Then
				result = result Or FontStyle.Bold
			End If
			If font.Italic Then
				result = result Or FontStyle.Italic
			End If
			Return result
		End Function
		#End Region
		#Region "CheckBox"
		Private Function CreateCheckBox() As CheckEdit
			Dim box As New CheckEdit()
			box.Text = String.Empty
			box.Properties.GlyphAlignment = DevExpress.Utils.HorzAlignment.Center
			box.Visible = False
			box.Parent = spreadsheetControl1
			AddHandler box.KeyDown, AddressOf OnCheckBoxKeyDown
			Return box
		End Function
		Private Sub OnCheckBoxKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
			If e.KeyCode = Keys.Return Then
				' Assign a value indicating whether the editor is checked to the active cell.
				spreadsheetControl1.ActiveCell.Value = checkBox.Checked
				checkBox.Visible = False
				checkBox.Parent = Nothing
				spreadsheetControl1.Focus()
			ElseIf e.KeyCode = Keys.Escape Then
				checkBox.Visible = False
				checkBox.Parent = Nothing
				spreadsheetControl1.Focus()
			End If
		End Sub
		Private Sub ShowCheckEdit(ByVal cell As Cell, ByVal bounds As Rectangle)
			checkBox.Checked = cell.Value.BooleanValue

			checkBox.BackColor = cell.FillColor

			checkBox.Parent = spreadsheetControl1
			checkBox.Bounds = bounds
			checkBox.Visible = True
			checkBox.Focus()
		End Sub
		#End Region
		#Region "DateEdit"
		Private Function CreateDateEdit() As DateEdit
			Dim edit As New DateEdit()
			edit.Visible = False
			edit.Parent = spreadsheetControl1
			AddHandler edit.KeyDown, AddressOf OnDateEditKeyDown
			Return edit
		End Function
		Private Sub OnDateEditKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
			If e.KeyCode = Keys.Return Then
				' Assign a date value currently selected in the DateEdit control to the active cell.
				spreadsheetControl1.ActiveCell.Value = CDate(dateEdit.EditValue)
				dateEdit.Visible = False
				dateEdit.Parent = Nothing
				spreadsheetControl1.Focus()
			ElseIf e.KeyCode = Keys.Escape Then
				dateEdit.Visible = False
				dateEdit.Parent = Nothing
				spreadsheetControl1.Focus()
			End If
		End Sub
		Private Sub ShowDateEdit(ByVal value As CellValue, ByVal bounds As Rectangle)
			dateEdit.EditValue = value.DateTimeValue

			dateEdit.Parent = spreadsheetControl1
			dateEdit.Bounds = bounds
			dateEdit.Visible = True
			dateEdit.Focus()
		End Sub
		#End Region
	End Class
End Namespace
