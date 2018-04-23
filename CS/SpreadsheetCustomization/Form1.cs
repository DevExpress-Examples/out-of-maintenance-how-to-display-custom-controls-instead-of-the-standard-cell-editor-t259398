using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace SpreadsheetCustomization {
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm {
        public static string[] categories = { "Meat/Poultry", "Condiments", "Seafood", "Dairy Products", "Grains/Cereals", "Beverages", "Confections" };
        IWorkbook workbook;
        Worksheet worksheet;

        Range dateColumn;
        Range discountColumn;
        Range categoryColumn;

        DateEdit dateEdit;
        LookUpEdit lookUpEdit;
        CheckEdit checkBox;

        public Form1() {
            InitializeComponent();
            workbook = spreadsheetControl1.Document;
            workbook.LoadDocument("Documents\\Document.xlsx", DocumentFormat.Xlsx);

            worksheet = workbook.Worksheets["Sales report"];
            dateColumn = worksheet["Table[Order Date]"];
            categoryColumn = worksheet["Table[Category]"];
            discountColumn = worksheet["Table[Discount]"];

            // Create custom controls to be displayed instead of the cell editor and specify their settings.
            lookUpEdit = CreateLookUp();
            checkBox = CreateCheckBox();
            dateEdit = CreateDateEdit();

            // Specify the SpreadsheetControl's options.
            spreadsheetControl1.Options.Behavior.Selection.AllowExtendSelection = false;
            spreadsheetControl1.Options.Behavior.Drag = DevExpress.XtraSpreadsheet.DocumentCapability.Disabled;
            spreadsheetControl1.Options.VerticalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden;
            spreadsheetControl1.Options.HorizontalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden;

            spreadsheetControl1.SelectionChanged += spreadsheetControl1_SelectionChanged;
            spreadsheetControl1.CellBeginEdit += spreadsheetControl1_CellBeginEdit;
            spreadsheetControl1.MouseWheel += spreadsheetControl1_MouseWheel;
        }

        void spreadsheetControl1_MouseWheel(object sender, System.Windows.Forms.MouseEventArgs e) {
            Rectangle activeCellBounds = GetActiveCellBounds();
            UpdateActiveEditorBounds(activeCellBounds);
        }

        void UpdateActiveEditorBounds(Rectangle newBounds) {
            if (newBounds.IsEmpty) {
                HideAllEditors();
                return;
            }

            if (lookUpEdit.Visible)
                lookUpEdit.Bounds = newBounds;
            else if (checkBox.Visible)
                checkBox.Bounds = newBounds;
            else if (dateEdit.Visible)
                dateEdit.Bounds = newBounds;
        }

        void spreadsheetControl1_SelectionChanged(object sender, EventArgs e) {
            HideAllEditors();
        }

        void HideAllEditors() {
            lookUpEdit.Visible = false;
            lookUpEdit.Parent = null;

            checkBox.Visible = false;
            checkBox.Parent = null;

            dateEdit.Visible = false;
            dateEdit.Parent = null;

            spreadsheetControl1.Focus();
        }

        void spreadsheetControl1_CellBeginEdit(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellCancelEventArgs e) {
            // Access the active cell.
            Cell activeCell = spreadsheetControl1.ActiveCell;
            // Obtain the bounds of the active cell.
            Rectangle activeCellRect = GetActiveCellBounds();
            // If the active cell is out of the visible range, return. 
            if (activeCellRect.IsEmpty) {
                e.Cancel = true;
                return;
            }

            // If the currently selected cell is in the "Category" column of the worksheet table, 
            // display the LookUpEdit control instead of the cell editor. 
            if (CanShowLookUp(activeCell)) {
                e.Cancel = true;
                ShowLookUp(activeCell, activeCellRect);
            }

            // If the currently selected cell is in the "Discount" column of the worksheet table, 
            // display the CheckEdit control instead of the cell editor.
            else if (CanShowCheckBox(activeCell)) {
                e.Cancel = true;
                ShowCheckEdit(activeCell, activeCellRect);
            }

            // If the currently selected cell is in the "Order Date" column of the worksheet table, 
            // display the DateEdit control instead of the cell editor.
            else if (CanShowDateEdit(activeCell)) {
                e.Cancel = true;
                ShowDateEdit(activeCell.Value, activeCellRect);
            }
        }
        Rectangle GetActiveCellBounds() {
            Cell activeCell = spreadsheetControl1.ActiveCell;
            return spreadsheetControl1.GetCellBounds(activeCell.RowIndex, activeCell.ColumnIndex);
        }
        bool CanShowLookUp(Cell activeCell) {
            return worksheet == workbook.Worksheets.ActiveWorksheet ? categoryColumn.IsIntersecting(activeCell) : false;
        }
        bool CanShowCheckBox(Cell activeCell) {
            return worksheet == workbook.Worksheets.ActiveWorksheet ? discountColumn.IsIntersecting(activeCell) : false;
        }
        bool CanShowDateEdit(Cell activeCell) {
            return worksheet == workbook.Worksheets.ActiveWorksheet ? dateColumn.IsIntersecting(activeCell) : false;
        }

        #region LookUp
        LookUpEdit CreateLookUp() {
            LookUpEdit cmbBox = new LookUpEdit();
            cmbBox.Properties.DataSource = categories;
            cmbBox.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            cmbBox.Visible = false;
            cmbBox.Parent = spreadsheetControl1;
            cmbBox.KeyDown += OnLookUpKeyDown;
            return cmbBox;
        }

        void OnLookUpKeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Return) {
                object editValue = lookUpEdit.EditValue;
                if (editValue == null)
                    return;

                // Assign a value of the currently selected item in the LookUpEdit to the active cell.  
                spreadsheetControl1.ActiveCell.Value = editValue.ToString();
                lookUpEdit.Visible = false;
                lookUpEdit.Parent = null;
                spreadsheetControl1.Focus();
            }
            else if (e.KeyCode == Keys.Escape) {
                lookUpEdit.Visible = false;
                lookUpEdit.Parent = null;
                spreadsheetControl1.Focus();
            }
        }
        void ShowLookUp(Cell cell, Rectangle bounds) {
            lookUpEdit.EditValue = cell.Value.TextValue;

            UpdateLookUpAppearance(cell);

            lookUpEdit.Parent = spreadsheetControl1;
            lookUpEdit.Bounds = bounds;
            lookUpEdit.Visible = true;
            lookUpEdit.Focus();
        }
        void UpdateLookUpAppearance(Cell source) {
            lookUpEdit.BackColor = source.Fill.BackgroundColor;
            SpreadsheetFont font = source.Font;
            lookUpEdit.ForeColor = font.Color;
            lookUpEdit.Font = new Font(font.Name, (float)font.Size, GetFontStyle(font));
        }
        FontStyle GetFontStyle(SpreadsheetFont font) {
            FontStyle result = FontStyle.Regular;
            if (font.Bold)
                result |= FontStyle.Bold;
            if (font.Italic)
                result |= FontStyle.Italic;
            return result;
        }
        #endregion
        #region CheckBox
        CheckEdit CreateCheckBox() {
            CheckEdit box = new CheckEdit();
            box.Text = String.Empty;
            box.Properties.GlyphAlignment = DevExpress.Utils.HorzAlignment.Center;
            box.Visible = false;
            box.Parent = spreadsheetControl1;
            box.KeyDown += OnCheckBoxKeyDown;
            return box;
        }
        void OnCheckBoxKeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Return) {
                // Assign a value indicating whether the editor is checked to the active cell.
                spreadsheetControl1.ActiveCell.Value = checkBox.Checked;
                checkBox.Visible = false;
                checkBox.Parent = null;
                spreadsheetControl1.Focus();
            }
            else if (e.KeyCode == Keys.Escape) {
                checkBox.Visible = false;
                checkBox.Parent = null;
                spreadsheetControl1.Focus();
            }
        }
        void ShowCheckEdit(Cell cell, Rectangle bounds) {
            checkBox.Checked = cell.Value.BooleanValue;

            checkBox.BackColor = cell.FillColor;

            checkBox.Parent = spreadsheetControl1;
            checkBox.Bounds = bounds;
            checkBox.Visible = true;
            checkBox.Focus();
        }
        #endregion
        #region DateEdit
        DateEdit CreateDateEdit() {
            DateEdit edit = new DateEdit();
            edit.Visible = false;
            edit.Parent = spreadsheetControl1;
            edit.KeyDown += OnDateEditKeyDown;
            return edit;
        }
        void OnDateEditKeyDown(object sender, System.Windows.Forms.KeyEventArgs e) {
            if (e.KeyCode == Keys.Return) {
                // Assign a date value currently selected in the DateEdit control to the active cell.
                spreadsheetControl1.ActiveCell.Value = (DateTime)dateEdit.EditValue;
                dateEdit.Visible = false;
                dateEdit.Parent = null;
                spreadsheetControl1.Focus();
            }
            else if (e.KeyCode == Keys.Escape) {
                dateEdit.Visible = false;
                dateEdit.Parent = null;
                spreadsheetControl1.Focus();
            }
        }
        void ShowDateEdit(CellValue value, Rectangle bounds) {
            dateEdit.EditValue = value.DateTimeValue;

            dateEdit.Parent = spreadsheetControl1;
            dateEdit.Bounds = bounds;
            dateEdit.Visible = true;
            dateEdit.Focus();
        }
        #endregion
    }
}
