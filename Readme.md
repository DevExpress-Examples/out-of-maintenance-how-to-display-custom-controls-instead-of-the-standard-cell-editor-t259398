<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/SpreadsheetCustomization/Form1.cs) (VB: [Form1.vb](./VB/SpreadsheetCustomization/Form1.vb))
<!-- default file list end -->
# How to display custom controls instead of the standard cell editor


This example demonstrates how to display custom controls instead of a cell's in-place editor.<br>• If an end-user tries to edit a cell located in the "Order Date" column of a worksheet table, the <a href="https://documentation.devexpress.com/#windowsforms/clsDevExpressXtraEditorsDateEdittopic">DateEdit</a> control is displayed, so that the user can select the required date in the drop-down calendar. <br>• If the end-user tries to edit a cell in the "Category" column of a table, the <a href="https://documentation.devexpress.com/#windowsforms/clsDevExpressXtraEditorsLookUpEdittopic">LookUpEdit</a> appears allowing the user to select one of predefined values.<br>• And finally, if the end-user activates a cell located in the "Discount" column, the <a href="https://documentation.devexpress.com/#windowsforms/clsDevExpressXtraEditorsCheckEdittopic">CheckEdit</a> control is displayed. It gives the user the true/false option to apply a 10% discount to the total amount.<br><img src="https://raw.githubusercontent.com/DevExpress-Examples/how-to-display-custom-controls-instead-of-the-standard-cell-editor-t259398/15.1.4+/media/56508594-1a64-11e5-80bf-00155d62480c.png"><br>To implement this behavior, subscribe to the <a href="https://documentation.devexpress.com/#WindowsForms/DevExpressXtraSpreadsheetSpreadsheetControl_CellBeginEdittopic">SpreadsheetControl.CellBeginEdit</a> event that is raised before the cell editor is activated, and then use the <a href="https://documentation.devexpress.com/#WindowsForms/DevExpressXtraSpreadsheetSpreadsheetControl_GetCellBoundstopic">SpreadsheetControl.GetCellBounds</a> method to obtain boundaries of the currently edited cell. Cell boundaries are defined by an instance of the <strong>System.Drawing.Rectangle</strong> class. Display the custom control over the cell editor by assigning the returned rectangle to the <strong>Control.Bounds</strong> property to specify the custom control's size and location.<br><br><strong>Starting from v16.1</strong>, this method of supplying custom cell editors is outdated. For details and an updated example, refer to <a href="https://www.devexpress.com/Support/Center/Example/Details/T385401">How to: Assign Custom In-Place Editors to Worksheet Cells</a>.

<br/>


