# Create, View, Edit and Save Your Excel Files Using WPF Spreadsheet

This example demonstrates how to create, view, and edit the Excel files using our **WPF Spreadsheet** component.

[WPF Spreadsheet](https://www.syncfusion.com/spreadsheet-editor-sdk/wpf-spreadsheet-editor) (SfSpreadsheet) is an Excel-inspired control that allows you to create, view, edit, and format Microsoft Excel files without having Microsoft Excel installed. It provides an integrated ribbon to cover any possible business scenario.

### Create a new Excel workbook

You can create a new workbook by using the [Create](https://help.syncfusion.com/cr/wpf/Syncfusion.UI.Xaml.Spreadsheet.SfSpreadsheet.html#Syncfusion_UI_Xaml_Spreadsheet_SfSpreadsheet_Create_System_Int32_) method. By default, a workbook will be created with a single worksheet. Use the following code to load the Spreadsheet workbook with a specified number of worksheets.

``` csharp
spreadsheetControl.Create(2);
```

### View the existing Excel sheet

The following code snippet illustrates how to view Excel files using the **WPF Spreadsheet** control.

``` csharp
/// View the Excel file.
spreadsheetControl.Open (@"..\..\Data\GettingStarted.xlsx");
                   (or)
ExcelEngine excelEngine = new ExcelEngine();

IWorkbook workbook = excelEngine.Excel.Workbooks.Open(@"..\..\Data\GettingStarted.xlsx");

spreadsheetControl.Open(workbook);
                   (or)
using (FileStream fileStream = new FileStream(@"..\..\Data\ GettingStarted.xlsx”, FileMode.Open))
{
     spreadsheetControl.Open(fileStream);
}
```

### Edit the values in an Excel file

The Spreadsheet provides support for editing, so you can modify and commit the cell values in a workbook. The following code snippet illustrates how to edit data in an Excel file using the **WPF Spreadsheet** control.

``` csharp
/// Editing a specific cell value.
var range = spreadsheetControl.ActiveSheet.Range[2,2];

spreadsheetControl.ActiveGrid.SetCellValue(range, "Syncfusion");

spreadsheetControl.ActiveGrid.InvalidateCell(2,2);
```

### Saving an Excel sheet

You can save Excel workbooks using the Save method. If the workbook already exists in the system drive, then it will be saved in the same location. Otherwise, the Save dialog box will open to let you save the workbook in a specified location.

Refer to the following code snippet.

``` csharp
/// Save the changes made in the file. If the file is not created yet, then it prompts to enter the filename to save.
spreadsheetControl.Save();

/// Save the changes made in the file using the SaveFileDialog.
spreadsheetControl.SaveAs();
```

## Blog reference
[Create, View, Edit and Save Your Excel Files Using WPF Spreadsheet](https://www.syncfusion.com/blogs/post/create-view-edit-and-save-your-excel-files-using-wpf-spreadsheet.aspx)

