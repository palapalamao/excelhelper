using System;
using System.Collections.Generic;
using System.Text;
using App = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;
namespace OverTimeStatistics
{
    /// <summary>
    /// 
    /// </summary>
    public class Excel : IDisposable
    {
        
        #region ... Variables  ...
        /// <summary>
        /// 
        /// </summary>
        App.Application mApp;
        /// <summary>
        /// 
        /// </summary>
        App.Workbook mWorkbook;
        /// <summary>
        /// 
        /// </summary>
        App.Sheets mSheets;
        /// <summary>
        /// 
        /// </summary>
        App.Worksheet mWorksheet;
        /// <summary>
        /// 
        /// </summary>
        App.Range mRange;


        private List<string> mSheetList;
        #endregion ...Variables...

        #region ... Events     ...

        #endregion ...Events...

        #region ... Constructor...
        /// <summary>
        /// Constructor of Excel
        /// </summary>
        public Excel()
        {
            Initialize();
        }

        public Excel(string path, bool visible)
        {
            mApp = new App.Application();
            Open(path, visible);
        }
        #endregion ...Constructor...

        #region ... Properties ...


        /// <summary>
        /// Gets or sets the sheet list.
        /// </summary>
        /// <value>
        /// The sheet list.
        /// </value> 



        /// <summary>
        /// Gets or sets the sheet list.
        /// </summary>
        /// <value>
        /// The sheet list.
        /// </value>
        public List<string> SheetList
        {
            get
            {
                return mSheetList;
            }
            set
            {
                if (mSheetList != value)
                {
                    mSheetList = value;
                }
            }
        }

        /// <summary>
        /// Get the current number of worksheets
        /// </summary>
        public int WorksheetCount
        {
            get
            {
                return mWorkbook.Worksheets.Count;
            }
        }
        /// <summary>
        /// Gets row count of current work sheet.
        /// </summary>
        public int RowCount
        {
            get
            {
                return mWorksheet.UsedRange.Rows.Count;
            }
        }
        /// <summary>
        /// Gets used column count
        /// </summary>
        public int ColumnCount
        {
            get
            {
                return mWorksheet.UsedRange.Columns.Count;
            }
        }
        /// <summary>
        /// Return a list of worksheet names
        /// </summary>
        public List<string> WorksheetNames
        {
            get
            {
                App.Worksheet worksheetName;

                List<string> names = new List<string>();

                for (int i = 0; i < mWorkbook.Sheets.Count; i++)
                {
                    worksheetName = (App.Worksheet)mSheets.get_Item(i + 1);
                    names.Add(worksheetName.Name);
                }

                return names;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool Visible
        {
            get
            {
                return mApp.Visible;
            }
            set
            {
                if (mApp.Visible != value)
                {
                    mApp.Visible = value;
                }
            }
        }


        #endregion ...Properties...

        #region ... Methods    ...
        /// <summary>
        /// 
        /// </summary>
        void Initialize()
        {
            mApp = new App.Application();
            mWorksheet = new App.Worksheet();
            mWorkbook = mApp.Workbooks.Add(true);
            mWorksheet = (App.Worksheet)mWorkbook.ActiveSheet;
            mSheets = mWorkbook.Worksheets;
            mApp.Visible = false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        public bool AddWorksheet(string name)
        {
            try
            {
                mWorkbook.Worksheets.Add(After: mSheets.get_Item(mWorkbook.Worksheets.Count));
                App.Worksheet sheet = (App.Worksheet)mSheets.get_Item(mWorkbook.Worksheets.Count);
                if (sheet.Name == name)
                    return false;
                sheet.Name = name;
                SetCurrentWorksheet(name);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }


        /// <summary>
        /// Set the current worksheet by name
        /// </summary>
        /// <param name="name"></param>
        public void SetCurrentWorksheet(string name)
        {
            for (int i = 0; i < mWorkbook.Sheets.Count; i++)
            {
                App.Worksheet worksheetCurrent = (App.Worksheet)mSheets.get_Item(i + 1);
                if (name.Equals(worksheetCurrent.Name))
                {
                    SetCurrentWorksheet(i + 1);
                    break;
                }
            }
        }


        /// <summary>
        /// Set the current worksheet by name
        /// </summary>
        /// <param name="name"></param>
        public void DelWorksheet(string name)
        {
            SetCurrentWorksheet(name);
            mWorksheet.Delete();
            // SetCurrentWorksheet(1);
        }
        /// <summary>
        /// Set the current worksheet by index
        /// </summary>
        /// <param name="index"></param>
        public void SetCurrentWorksheet(int index)
        {
            if (index < 1)
                throw new Exception("Index out of bounds");
            if (index > mWorkbook.Sheets.Count)
                throw new Exception("Index out of bounds");

            mWorksheet = (App.Worksheet)mSheets.get_Item(index);
            mWorksheet.Activate();
        }
        /// <summary>
        /// 
        /// </summary>
        public void Clean()
        {
            try
            {
                if (mWorkbook != null)
                {
                    mWorkbook.Close(false, null, null);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mWorkbook);
                    mApp.Workbooks.Close();
                }
                if (mApp != null)
                {
                    mApp.Quit();
                    mApp.Workbooks.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mApp);
                }

                mApp = null;
                mSheets = null;
                mWorksheet = null;
                mWorkbook = null;

            }
            catch (Exception) { }
            finally
            {
                GC.Collect();
            }
        }
        /// <summary>
        /// Saves the current excel document
        /// </summary>
        public void Save()
        {
            mWorkbook.Save();
        }
        /// <summary>
        /// Saves as excel 2007 format
        /// </summary>
        /// <param name="fullPath">Full path with .xlsx extension</param>
        public void SaveAs2007(string fullPath)
        {
            mWorkbook.SaveAs(fullPath, 51);
        }

        public void SaveCopyAs(string strExcelFileName)
        {
            mWorkbook.SaveCopyAs(strExcelFileName);
        }

        /// <summary>
        /// Saves as excel 2003 format
        /// </summary>
        /// <param name="fullPath">Full path with .xls extension</param>
        public void SaveAs2003(string fullPath)
        {
            mWorkbook.SaveAs(fullPath, App.XlFileFormat.xlWorkbookNormal);
        }
        /// <summary>
        /// Close 
        /// </summary>
        public void Close()
        {
            Clean();
        }
        /// <summary>
        /// Open an exisiting excel Document
        /// </summary>
        public bool Open(string path, bool visible)
        {
            if (!path.EndsWith("xls") && !path.EndsWith("xlsx"))
            {
                Console.WriteLine("Invalid file format");
                return false;
            }
            if (!File.Exists(path))
            {
                Console.WriteLine("File does not exist");
                return false;
            }

            try
            {
                mWorkbook = mApp.Workbooks.Open(path, Type.Missing, false);
                mWorksheet = (App.Worksheet)mWorkbook.Worksheets[1];
                mSheets = mWorkbook.Worksheets;
                mApp.Visible = visible;
            }
            catch (Exception E)
            {
                Console.WriteLine(E.Message);
                Clean();
                return false;
            }

            return true;
        }
        /// <summary>
        /// Open an exisiting excel Document
        /// </summary>
        /// <param name="path">File path.</param>
        /// <returns></returns>
        public bool Open(string path)
        {
            return Open(path, true);
        }
        /// <summary>
        /// Set cell's value.
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public Excel SetCell(int rowIndex, int columnIndex, string value)
        {
            mApp.Cells[rowIndex, columnIndex] = value;
            return this;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public Excel ColumnAutoFit(int rowIndex, int columnIndex)
        {
            App.Range range = (App.Range)mApp.Columns[columnIndex];
            range.AutoFit();
            return this;
        }

        /// <summary>
        /// Gets cell's value.
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public string GetCell(int rowIndex, int columnIndex)
        {
            //mWorksheet.get_Range(
            App.Range range = mWorksheet.UsedRange;
            App.Range row = (App.Range)range.Rows[rowIndex];
            App.Range cell = (App.Range)row.Cells[columnIndex];
            if (cell.get_Value() == null)
                return "";
            return cell.get_Value().ToString();
        }





        /// <summary>
        /// Gets the cell comment.
        /// </summary>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <returns></returns>
        public string GetCellComment(int rowIndex, int columnIndex)
        {
            //mWorksheet.get_Range(
            App.Range range = mWorksheet.UsedRange;
            App.Range row = (App.Range)range.Rows[rowIndex];
            App.Range cell = (App.Range)row.Cells[columnIndex];
            if (cell.Comment == null)
                return "";
            return cell.Comment.Text();
        }






        /// <summary>
        /// Sets the cell comment.
        /// </summary>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <returns></returns>
        public void SetCellComment(int rowIndex, int columnIndex, string comment)
        {
            //mWorksheet.get_Range(
            App.Range range = mWorksheet.UsedRange;
            App.Range row = (App.Range)range.Rows[rowIndex];
            App.Range cell = (App.Range)row.Cells[columnIndex];

            if (cell.Comment == null)
            {
                cell.AddComment();
                cell.Comment.Text(comment);
            }
            else
            {
                cell.Comment.Text(comment);
            }
        }



        /// <summary>
        /// Sets the width of the column.
        /// </summary>
        /// <param name="sRow">The s row.</param>
        /// <param name="sCol">The s col.</param>
        /// <param name="width">The width.</param>
        /// <returns></returns>
        public Excel SetColumnWidth(int sRow, int sCol, int width)
        {
            object sell1 = mWorksheet.Cells[sRow, sCol];
            object sell2 = mWorksheet.Cells[sRow, sCol];
            mRange = mWorksheet.get_Range(sell1, sell2);
            mRange.ColumnWidth = width;
            return this;
        }


        /// <summary>
        /// Sets the sel format text.
        /// </summary>
        /// <param name="sRow">The s row.</param>
        /// <param name="sCol">The s col.</param>
        /// <returns></returns>
        public Excel SetSelFormatText(int sRow, int sCol)
        {
            object sell1 = mWorksheet.Cells[sRow, sCol];
            object sell2 = mWorksheet.Cells[sRow, sCol];
            mRange = mWorksheet.get_Range(sell1, sell2);
            mRange.NumberFormatLocal = "@";
            return this;
        }
        /// <summary>
        /// Sets the alignment left.
        /// </summary>
        /// <param name="sRow">The s row.</param>
        /// <param name="sCol">The s col.</param>
        /// <returns></returns>
        public Excel SetAlignmentLeft(int sRow, int sCol)
        {
            object sell1 = mWorksheet.Cells[sRow, sCol];
            object sell2 = mWorksheet.Cells[sRow, sCol];
            mRange = mWorksheet.get_Range(sell1, sell2);
            mRange.HorizontalAlignment = 2;
            mRange.VerticalAlignment = 2;
            return this;
        }


        /// <summary>
        /// Sets the alignment left.
        /// </summary>
        /// <param name="sRow">The s row.</param>
        /// <param name="sCol">The s col.</param>
        /// <returns></returns>
        public Excel SetTextFormat(int sRow, int sCol)
        {
            object sell1 = mWorksheet.Cells[sRow, sCol];
            object sell2 = mWorksheet.Cells[sRow, sCol];
            mRange = mWorksheet.get_Range(sell1, sell2);
            mRange.NumberFormatLocal = "@";
            return this;
        }


        /// <summary>
        /// Formats the font in a cell, bold italic and underline take a bool as a value.
        /// Fontsize font color and font type are all nullable so you can write null if you dont want to specify
        /// </summary>
        public void FormatCellFont(string location, bool bold, bool italic, bool underline, double? fontsize, Color? fontcolor, string fontname)
        {
            mRange = mWorksheet.get_Range(location);

            mRange.Font.Bold = bold;
            mRange.Font.Italic = italic;
            mRange.Font.Underline = underline;

            if (fontsize != null)
                mRange.Font.Size = fontsize;
            if (fontcolor != null)
                mRange.Font.Color = ColorTranslator.ToOle(fontcolor.Value);
            if (!string.IsNullOrEmpty(fontname))
                mRange.Font.Name = fontname;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="location"></param>
        /// <param name="formatAction"></param>
        public Excel FormatCell(string location, Action<App.Range> formatAction)
        {
            mRange = mWorksheet.get_Range(location);
            formatAction(mRange);
            return this;
        }

        public Excel FreezePanes(int sRow, int sCol, int eRow, int eCol)
        {
            SetCellSelect(sRow, sCol, eRow, eCol);
            mApp.ActiveWindow.FreezePanes = true;
            return this;
        }
        /// <summary>
        /// Sets the range background.
        /// </summary>
        /// <param name="sRow">The s row.</param>
        /// <param name="sCol">The s col.</param>
        /// <param name="eRow">The e row.</param>
        /// <param name="eCol">The e col.</param>
        /// <param name="colorIndex">Index of the color.</param>
        public void SetRangeBackground(int sRow, int sCol, int eRow, int eCol, int colorIndex)
        {
            object sell1 = mWorksheet.Cells[sRow, sCol];
            object sell2 = mWorksheet.Cells[eRow, eCol];
            mRange = mWorksheet.get_Range(sell1, sell2);
            mRange.Interior.ColorIndex = colorIndex;

        }

        /// <summary>
        /// Sets the color of the range font.
        /// </summary>
        /// <param name="sRow">The s row.</param>
        /// <param name="sCol">The s col.</param>
        /// <param name="eRow">The e row.</param>
        /// <param name="eCol">The e col.</param>
        /// <param name="colorIndex">Index of the color.</param>
        public void SetRangeFontColor(int sRow, int sCol, int eRow, int eCol, int colorIndex)
        {
            object sell1 = mWorksheet.Cells[sRow, sCol];
            object sell2 = mWorksheet.Cells[eRow, eCol];
            mRange = mWorksheet.get_Range(sell1, sell2);
            mRange.Font.ColorIndex = colorIndex;
        }


        /// <summary>
        /// Sets the cell select.
        /// </summary>
        /// <param name="sRow">The s row.</param>
        /// <param name="sCol">The s col.</param>
        /// <param name="eRow">The e row.</param>
        /// <param name="eCol">The e col.</param>
        public void SetCellSelect(int sRow, int sCol, int eRow, int eCol)
        {
            object sell1 = mWorksheet.Cells[sRow, sCol];
            object sell2 = mWorksheet.Cells[eRow, eCol];
            mRange = mWorksheet.get_Range(sell1, sell2);
            mRange.Select();
        }


        /// <summary>
        /// Copy the cell select.
        /// </summary>
        /// <param name="sRow">The s row.</param>
        /// <param name="sCol">The s col.</param>
        /// <param name="eRow">The e row.</param>
        /// <param name="eCol">The e col.</param>
        public void CopyCellSelect(int sRow, int sCol, int eRow, int eCol)
        {
            object sell1 = mWorksheet.Cells[sRow, sCol];
            object sell2 = mWorksheet.Cells[eRow, eCol];
            mRange = mWorksheet.get_Range(sell1, sell2);
            //mRange.Select();
            //mRange.Activate();
            mRange.Copy();
        }



        /// <summary>
        /// Copy the cell select.
        /// </summary>
        /// <param name="sRow">The s row.</param>
        /// <param name="sCol">The s col.</param>
        /// <param name="eRow">The e row.</param>
        /// <param name="eCol">The e col.</param>
        public void PasteCellSelect(int sRow, int sCol, int eRow, int eCol)
        {
            object sell1 = mWorksheet.Cells[sRow, sCol];
            object sell2 = mWorksheet.Cells[eRow, eCol];
            mRange = mWorksheet.get_Range(sell1, sell2);
            mRange.Select();
            mRange.PasteSpecial();
        }

        #endregion ...Methods...

        #region ... Interfaces ...
        #region IDisposable Members
        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            Clean();
        }

        #endregion
        #endregion ...Interfaces...
    }
}
