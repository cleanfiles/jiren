# jiren
export to excel
/* Export All Revit Schedules to One Excel file
 * Written by: K C Tang
 * Using: SharpDevelop.
 * 
 * Date last revised: 24/5/2015
 * Revision notes: Handle errors when there are no "Dim - " schedules to
 *    fill the "All Dim" worksheet to serve as the source of 
 *    the pivot table on the "QS Desc" worksheet.
 *    Handle errors when the folder is read only, 
 *    such as a folder of the sample projects provided by Revit.
 * Revision notes on 6/5/2015:
 *    Speed drastically increased (reduced to 1/4) by using Excel functions
 *    as much as possible instead of manipulating Excel file cell by cell.
 *    Pivot table used for "QS Desc" worksheet.
 *    Sheet header added.
 *    Frozen panes set.
 *    Page setup set.
 * Revision notes on 19/1/2015:
 *    Bug fixing and general improvements.
 * Date first released for use : 31/12/2014
 *    Notes: Export all Revit view schedules to one Excel file.
 *    Schedules with names beginning with "Dim - " will have:
 *    - columns "Type", "QS Tag" and "QS Unit" combined into one column;
 *    - column "QS Qty" moved to next to the combined column;
 *    - an "All Dim" worksheet created to contain all these schedules; and
 *    - a "QS Desc" worksheet created to contain unique list of "Type : QS Tag : QS Unit".
 * 
 */
 
// using libraries
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Excel = Microsoft.Office.Interop.Excel; 
/* Microsoft.Office.Interop.Excel must be added separately
 * by selecting SharpDevelop's menu: Project > Add References,
 * and searching for it, then selecting it.
 */
 
namespace KCTCL
{
  [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
  [Autodesk.Revit.DB.Macros.AddInId("E77FD3DE-05E8-4FD3-B85A-116F5B6F2EEF")]
  public partial class ThisDocument
  {
    private void Module_Startup(object sender, EventArgs e)
    {
 
    }
 
    private void Module_Shutdown(object sender, EventArgs e)
    {
 
    }
 
    #region Revit Macros generated code
    private void InternalStartup()
    {
      this.Startup += new System.EventHandler(Module_Startup);
      this.Shutdown += new System.EventHandler(Module_Shutdown);
    }
    #endregion
 
    public void ExportAllSchedulesToOneExcel()
    {
      // all data names must be intialized first and have their types declared with type names before them
      // define row number to insert column header
      const int col_header_row = 3; // const int = integer constant type
      // keep the starting time
      DateTime time_start = DateTime.Now; // Datetime = datetime type
      // select active Revit document
      Document doc = this.Document; // Document = document type
      // get filename from doc.Title
      string filename_no_ext = doc.Title; // string = string type
      // add ".rvt" temporarily to doc.Title not ending with ".rvt" 
      // because file explorer may have been set to hide the extension
      if (!filename_no_ext.EndsWith(".rvt")) // ! = not
      {
        filename_no_ext = filename_no_ext + ".rvt"; // + = join text together
      }
      // get active folder name by removing the full file name 
      // from the full pathname which contains the full file name
      string folder_name = doc.PathName.Replace(filename_no_ext, ""); // replace filename with nothing
      // change file extension to the current datetime string
      // to avoid overwriting existing files
      filename_no_ext = filename_no_ext.Replace(".rvt", 
        DateTime.Now.ToString("-yyyyMMdd-HHmmss")); // line considered complete only if ending with ";"
      // initilize Excel variables
      Excel.Application xlApp;
      Excel.Workbook xlWorkBook;
      Excel.Worksheet xlWorkSheet;
      Excel.Worksheet xlWorkSheetAllDim;
      Excel.Range xlRange;
      Excel.Range xlRange2;
      Excel.QueryTable xlQuery;
      xlApp = new Excel.Application();
      // check whether Excel is installed
      if (xlApp == null)
      {
        TaskDialog.Show("ExportAllSchedulesToOneExcel", "Excel is not installed!!");
        return;
      }
      // define an object to represent default value
      object default_value = System.Reflection.Missing.Value; // object = object type
      // create new workbook, which by default contains at least 1 worksheet
      xlWorkBook = xlApp.Workbooks.Add(default_value);
      // initialize 2 worksheet variables, all referring to Sheet1 for the time being
      xlWorkSheetAllDim = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
      xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
      // rename Sheet1 to contain contents of all future worksheets
      // with names starting with "Dim - "
      xlWorkSheetAllDim.Name = "All Dim";
      // maximize workbook window
      xlApp.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized;
      // show menu bars
      xlApp.Visible = true;
      // read viewschedules in Revit active document
      ViewScheduleExportOptions opt = new ViewScheduleExportOptions();
      FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));
      if (collector.ToElementIds().Count == 0) // == means compare for equality
      {
        TaskDialog.Show("ExportAllSchedulesToOneExcel", "No schedule available!!");
        // close workbook without saving
        xlWorkBook.Close(false, default_value, default_value);
        xlApp.Quit();
        // release objects
        releaseObject(xlWorkSheet);
        releaseObject(xlWorkSheetAllDim);
        releaseObject(xlWorkBook);
        releaseObject(xlApp);
        return;
      }
      // sort elements in collector in ascending order
      IOrderedEnumerable< ViewSchedule > sorted_collector =
        from ViewSchedule view_schedule in collector orderby view_schedule.Name ascending select view_schedule;
      // process schedule in ascending order
      int all_dim_new_row = 0;
      foreach (ViewSchedule view_schedule in sorted_collector)
      {
        // check if schedule name too long
        if (view_schedule.Name.Length > 31 )
        {
          TaskDialog.Show("ExportAllSchedulesToOneExcel", 
            view_schedule.Name + "\n" + "Schedule name should not be more than 31 characters!!");
          // release objects
          releaseObject(xlWorkSheet);
          releaseObject(xlWorkSheetAllDim);
          releaseObject(xlWorkBook);
          releaseObject(xlApp);
          return;
        }
      }
      foreach (ViewSchedule view_schedule in sorted_collector)
      {
        if (view_schedule.Name.StartsWith("<")) 
        {
           // skip schedule with name beginning with "<", such as "<Revision Schedule>"
        } else 
        {
          // reduce filename length longer than 31
          if (31 < view_schedule.Name.Length ) 
          {
            view_schedule.Name = view_schedule.Name.Substring(0, 14) + " name length > 31";
          }
          // replace special character with "_"
          view_schedule.Name = view_schedule.Name
            .Replace( ':', '_' )
            .Replace( '*', '_' )
            .Replace( '?', '_' )
            .Replace( '/', '_' )
            .Replace( '\\', '_' )
            .Replace( '[', '_' )
            .Replace( ']', '_' ); 
          // export schedule to txt file
          try {
            view_schedule.Export(folder_name, filename_no_ext + ".txt", opt);    
          } catch (Exception) {
            TaskDialog.Show("Exporting view schedules", 
                            "Errors occurred -\n" +
                            "possibly the folder is read only\n" +
                            "e.g. in the case of sample projects provided by Revit,\n" +
                            "save the project to another folder first.\n\n" +
                            "Close to exit program.");
            // release objects
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkSheetAllDim);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            return;
          }
          // add a worksheet
          xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add(default_value);
          // move it to become the last worksheet
          xlWorkSheet.Move(default_value, xlWorkBook.Worksheets[xlWorkBook.Worksheets.Count]);
          // name worksheet as schedule name
          xlWorkSheet.Name = view_schedule.Name;
          // import txt file into worksheet starting at cell at column A, one row above col_header_row
          xlQuery = xlWorkSheet.QueryTables.Add(
            "TEXT;" + folder_name + filename_no_ext + ".txt", 
            xlWorkSheet.get_Range("A" + (col_header_row - 1)));
          xlQuery.RefreshStyle = Excel.XlCellInsertionMode.xlInsertEntireRows;
          xlQuery.Refresh(false); // false means refresh but not return until refresh is finished 
          xlQuery.Delete(); // delete the query
          // input into All Dim worksheet for schedules with names starting with "Dim - "
          if (view_schedule.Name.StartsWith("Dim - ")) 
          {
            // insert blank new column A
            xlWorkSheet.get_Range("A1").EntireColumn.Insert();
            // find Type, QS Tag and QS Unit columns
            int col_Type = 0;
            int col_QS_Tag = 0;
            int col_QS_Unit = 0;
            int col_QS_Qty = 0;
            col_Type = xlColumnFindExact(xlWorkSheet, "Type", col_header_row);
            col_QS_Tag = xlColumnFindExact(xlWorkSheet, "QS Tag", col_header_row);
            col_QS_Unit = xlColumnFindExact(xlWorkSheet, "QS Unit", col_header_row);
            // define cell formula of column A starting from col_header_row
            xlRange = xlWorkSheet.Range["A" + col_header_row]; // source range to copy from
            xlRange2 = xlWorkSheet.Range["A" + col_header_row, "A" + xlRowLast(xlWorkSheet)]; // target range to copy to
            string range_formula = "=";
            string colon = "&\" : \"&"; // which stands for quote &" : "& unquote
            if (col_Type != 0)
            {
              range_formula += xlColumnAddress(col_Type) + col_header_row;
               }
               if (col_QS_Tag != 0)
               {
              range_formula += colon + xlColumnAddress(col_QS_Tag) + col_header_row;
            }
               if (col_QS_Unit != 0)
            {
              range_formula += colon + xlColumnAddress(col_QS_Unit) + col_header_row;
            }
            if (range_formula != "=")
            {
              xlRange.Formula = range_formula;
              xlRange.AutoFill(xlRange2, Excel.XlAutoFillType.xlFillCopy);
            }
            // remove cell formula and leave value for the combined column A
            xlRange2.Value2 = xlRange2.Value2;
            // remove Type, QS Tag and QS Unit columns
            col_Type = xlColumnFindExact(xlWorkSheet, "Type", col_header_row);
            if (col_Type != 0)
            {
              xlRange = (Excel.Range)xlWorkSheet.Cells[1, col_Type];
              xlRange.EntireColumn.Delete(default_value);
            }
            col_QS_Tag = xlColumnFindExact(xlWorkSheet, "QS Tag", col_header_row);
            if (col_QS_Tag != 0)
            {
              xlRange = (Excel.Range)xlWorkSheet.Cells[1, col_QS_Tag];
              xlRange.EntireColumn.Delete(default_value);
            }
            col_QS_Unit = xlColumnFindExact(xlWorkSheet, "QS Unit", col_header_row);
            if (col_QS_Unit != 0)
            {
              xlRange = (Excel.Range)xlWorkSheet.Cells[1, col_QS_Unit];
              xlRange.EntireColumn.Delete(default_value);
            }
            // move QS Qty column to column D, but if it is before column D, 
            // move to column E to compensate the shifting of columns after cut
            col_QS_Qty = xlColumnFindExact(xlWorkSheet, "QS Qty", col_header_row);
            if (col_QS_Qty < 4)
            {
              xlColumnMove(xlWorkSheet, col_QS_Qty, 5);
            } else
            {
              xlColumnMove(xlWorkSheet, col_QS_Qty, 4);
            }
            // move combined column A to column C
            xlColumnMove(xlWorkSheet, 1, 4);
            // bold down to col_header_row
            xlRange = xlWorkSheet.get_Range("A1", "A" + col_header_row);
            xlRange.EntireRow.Font.Bold = true;
            // copy whole worksheet to All Dim to the next new row
            all_dim_new_row += 1;
            xlWorkSheet.UsedRange.Copy(xlWorkSheetAllDim.get_Range("A"+all_dim_new_row));
            all_dim_new_row = xlRowLast(xlWorkSheetAllDim);
          } else
          {
            // bold down to col_header_row
            xlWorkSheet.get_Range("A1", "A" + col_header_row).EntireRow.Font.Bold = true;
          }
          // delete txt file
          System.IO.File.Delete(folder_name + filename_no_ext + ".txt");
        }
      }
      // move it to become the first worksheet
      xlWorkSheetAllDim.Move(xlWorkBook.Worksheets[1]);
      // add and name a worksheet to contain unique QS Desc
      xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add(default_value);
      xlWorkSheet.Name = "QS Desc";
      // generate pivot table if "All Dim" has data rows
      if (xlRowLast(xlWorkSheetAllDim) > col_header_row)
      {
        // define pivot table data source
        xlRange = xlWorkSheetAllDim.get_Range("C" + col_header_row,"D" + xlRowLast(xlWorkSheetAllDim));
        Excel.PivotCache xlPivotCache = xlWorkBook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, xlRange);
        Excel.PivotTables xlPivotTables = (Excel.PivotTables)xlWorkSheet.PivotTables();
        // define pivot table in the QS Desc worksheet
        Excel.PivotTable xlPivotTable = xlPivotTables.Add(xlPivotCache, xlWorkSheet.Range["A2"], "QS Desc", default_value, default_value);
        xlPivotTable.SmallGrid = false;
        xlPivotTable.ShowTableStyleRowStripes = true;
        xlPivotTable.TableStyle2 = "PivotStyleLight1";
        Excel.PivotField xlPivotField = (Excel.PivotField)xlPivotTable.PivotFields("Type : QS Tag : QS Unit");
        xlPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
        xlPivotTable.AddDataField(xlPivotTable.PivotFields("QS Qty"), "Sum of QS Qty", Excel.XlConsolidationFunction.xlSum);
      }
      // format QS Qty column
      xlRange = xlWorkSheet.get_Range("B1"); 
      xlRange.EntireColumn.NumberFormat = "#,##0.00";
      // move it to become the first worksheet
      xlWorkSheet.Move(xlWorkBook.Worksheets[1]);
      // loop through all worksheets
      int loop_A_Mx = xlWorkBook.Worksheets.Count;
      for (int loop_A = loop_A_Mx; loop_A >= 1; loop_A--)
      {
        // get worksheet
        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(loop_A);
        // freeze top rows
        ((Excel._Worksheet)xlWorkSheet).Activate(); // cast to Excel._Worksheet to avoid the ambiguity that "Activate" is also used as an event
        xlWorkSheet.Application.ActiveWindow.SplitRow = col_header_row;
        xlWorkSheet.Application.ActiveWindow.FreezePanes = true;
        // insert new column A
        xlWorkSheet.get_Range("A1").EntireColumn.Insert();
        // assign column A with row number
        int last_row = xlRowLast(xlWorkSheet);
        xlWorkSheet.Cells[1,1] = 1;
        xlRange = xlWorkSheet.get_Range("A1","A" + xlRowLast(xlWorkSheet));
        xlRange.Font.Bold = false;
        xlRange.DataSeries(default_value,
          Excel.XlDataSeriesType.xlDataSeriesLinear,
          Excel.XlDataSeriesDate.xlDay,
          "1", default_value, default_value);
        // autofit column widths
        xlWorkSheet.Columns.EntireColumn.AutoFit();
        // assign cell B1 with filename
        xlWorkSheet.Cells[1,2] = filename_no_ext.ToUpper();
        xlWorkSheet.get_Range("B1").Font.Bold = true;
        // define page setup
        xlWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
        xlWorkSheet.PageSetup.PrintTitleRows = "$1:$" + col_header_row;
        xlWorkSheet.PageSetup.LeftFooter = filename_no_ext;
        xlWorkSheet.PageSetup.RightFooter = "&P/&N";
        xlWorkSheet.PageSetup.Zoom = false; // needs to be false for FitToPagesWide to work
        xlWorkSheet.PageSetup.FitToPagesTall = false; // need to be false for FitToPagesWide to work
        xlWorkSheet.PageSetup.FitToPagesWide = 1;
      }
      // save workbook
      xlWorkBook.SaveAs(folder_name + filename_no_ext,
        default_value, default_value, default_value, 
        default_value, default_value,
        Excel.XlSaveAsAccessMode.xlNoChange, 
        default_value, true, default_value, 
        default_value, true);
      // release objects
      releaseObject(xlWorkSheet);
      releaseObject(xlWorkSheetAllDim);
      releaseObject(xlWorkBook);
      releaseObject(xlApp);
      TaskDialog.Show("Export All Schedules To One Excel", 
        "Finished!" + "\nTime Spent " + DateTime.Now.Subtract(time_start).Seconds + " seconds");
    }
 
    private string xlCellAddress(int row, int col)
    {
      // change cell address from (100, 1) to (A100) style
      string prompt = (row + "\n\t" + col + "\n\t");   
      if (row < 1 || row > 1048576) 
      {
        TaskDialog.Show("Excel Row Number", "Error - must be within 1 - 1048576!!");
        return null;
      }
      // append row number to alphabetical column reference
      return xlColumnAddress(col) + row.ToString();
    }
 
    private string xlCellValue2(Excel.Worksheet w_s, int row, int col)
    {
      // return value of worksheet cell
      Excel.Range xlRange = (Excel.Range)w_s.Cells[row,col];
      if (xlRange.Value2 != null)
      {
        return xlRange.Value2.ToString();
      } else
      {
        return "";
      }
    }
 
    private string xlColumnAddress(int col)
    {
      // convert column number to alphabetical reference
      if (col < 1 || col > 16384) 
      {
        TaskDialog.Show("Excel Column Number", "Error - must be within 1 - 16384!!");
        return null;
      }
      int remainder = 0;
      string result = "";
      for (int loop_A = 0; loop_A < 3; loop_A++) 
      {
        // get remainder after division by 26
        remainder = ((col - 1) % 26) + 1;
        if (remainder != 0) 
        {
          // match the remainder to alphabets A to Z where A is char 65
          // precede the alphabet to the previous result
          result = Convert.ToChar(remainder + 64).ToString() + result;
        }
        col = ((col - 1) / 26);
        // do it three times
      }
      return result;
    }
 
    private int xlColumnFindExact(Excel.Worksheet w_s, string find_what, int which_row)
    {
      object default_value = System.Reflection.Missing.Value;
      Excel.Range xlFound = w_s.get_Range(which_row + ":" + which_row);
      xlFound = xlFound.Find(find_what, default_value, 
        Excel.XlFindLookIn.xlValues, 
        Excel.XlLookAt.xlWhole,
        Excel.XlSearchOrder.xlByRows, 
        Excel.XlSearchDirection.xlNext, 
        false, default_value, default_value);
      if (xlFound != null) 
      {
        return xlFound.Column;
      } else 
      {
        return 0;
      }
    }
 
    private void xlColumnMove(Excel.Worksheet w_s, int col_from, int col_to)
    {
      if ((col_from != 0) & (col_to != 0) & (col_from != col_to))
      {
        Excel.Range from_range = (Excel.Range)w_s.Cells[1,col_from];
        Excel.Range to_range = (Excel.Range)w_s.Cells[1,col_to];
        to_range.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, from_range.EntireColumn.Cut());
      }
    }
 
    private int xlRowLast(Excel.Worksheet w_s)
    {
      // return last used row number of worksheet
      return w_s.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell,Type.Missing).Row;
    }
 
    private void releaseObject(object obj)
    {
      try
      {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      }
      catch (Exception ex)
      {
        obj = null;
        TaskDialog.Show("Excel file created","Exception Occurred while releasing object " + ex.ToString());
      }
      finally
      {
        GC.Collect();
      }
    }
  }
} 
