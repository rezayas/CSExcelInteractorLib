using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelInteractorLib
{
    public class ExcelInteractor
    {
        string _fileName = "";
        Excel.Application _xlApp;
        Excel.Workbook _xlWorkBook;
        Excel._Worksheet _xlActiveWorkSheet;

        public enum enumRangeDirection : int
        {
            RightEnd = 1,
            LeftEnd = 2,
            DownEnd = 3,
            UpEnd = 4,
        }
        public enum enumAxis : int
        {
            x = 0,
            y = 1,
        }
        public enum enumBorder : int
        {
            Top = 1,
            Bottom = 2,
            Left = 3,
            Right = 4,
        }
        public enum enumAlignment : int
        {
            Left = 1,
            Center = 2,            
            Right = 3,            
        }

        public bool Visible
        {
            get
            {
                return _xlApp.Visible;
            }
            set
            {
                _xlApp.Visible = value;
            }
        }

        public void OpenAnExistingExcelFile(bool visible)
        {
            _xlApp = new Excel.Application();

            // show an open file dialog window
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Open the Excel file you would like to open.";
            dlg.Filter = "Excel Micro-Enabled Files (*.xlsm)|*.xlsm|Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            //dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Templates);
            dlg.ShowReadOnly = true;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _fileName = dlg.FileName;
            }

            //open the excel file                       
            object misValue = System.Reflection.Missing.Value;            
            _xlWorkBook = _xlApp.Workbooks.Open(_fileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            _xlApp.Visible = visible;
            _xlActiveWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.get_Item(1);
        }
        public void CreateANewExcelFile()
        {
            // crete a new Excel file
            _xlApp = new Excel.Application();
        }

        public void Recalculate()
        {
            _xlWorkBook.Application.Calculate();
        }
        public void Save()
        {
            _xlWorkBook.Save();
        }
        public void Quit()
        {
            _xlApp.Quit();
        }
        public string GetFileName()
        {
            return _xlWorkBook.Name;
        }

        public void ActivateSheet(int sheetIndex)
        {
            _xlActiveWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.get_Item(sheetIndex);            
        }
        public void ActivateSheet(string sheetName)
        {
            _xlActiveWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.get_Item(sheetName);
        }

        public Excel.Worksheet GetASheet(string sheetName)
        {
            return (Excel.Worksheet)_xlWorkBook.Worksheets.get_Item(sheetName);
        }
        private Excel.Range GetCell(int rowIndex, int colIndex)
        {
            return (Excel.Range)_xlActiveWorkSheet.Cells[rowIndex, colIndex];
        }
        private Excel.Range GetCell(string cellName)
        {
            return (Excel.Range)_xlActiveWorkSheet.get_Range(cellName, cellName);
        }
        public Excel.Range GetCell(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            Excel.Range myRange1 = GetCell(rowIndex, colIndex);
            Excel.Range myRange2;
            switch (direction)
            {
                case enumRangeDirection.RightEnd:
                    myRange2 = myRange1.get_End(Excel.XlDirection.xlToRight);
                    break;
                case enumRangeDirection.LeftEnd:
                    myRange2 = myRange1.get_End(Excel.XlDirection.xlToLeft);
                    break;
                case enumRangeDirection.DownEnd:
                    myRange2 = myRange1.get_End(Excel.XlDirection.xlDown);
                    break;
                case enumRangeDirection.UpEnd:
                    myRange2 = myRange1.get_End(Excel.XlDirection.xlUp);
                    break;
                default:
                    myRange2 = myRange1;
                    break;
            }
            return myRange2;
        }        

        public Excel.Range GetRange(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex)
        {
            Excel.Range myRange1 = GetCell(rowStartIndex, colStartIndex);
            Excel.Range myRange2 = GetCell(rowEndIndex, colEndIndex);
            return _xlActiveWorkSheet.get_Range(myRange1, myRange2);
        }
        public Excel.Range GetRange(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            Excel.Range myRange1 = GetCell(rowIndex, colIndex);
            Excel.Range myRange2;
            switch (direction)
            {
                case enumRangeDirection.RightEnd:
                    myRange2 = myRange1.get_End(Excel.XlDirection.xlToRight);
                    break;
                case enumRangeDirection.DownEnd:
                    myRange2 = myRange1.get_End(Excel.XlDirection.xlDown);
                    break;
                default:
                    myRange2 = myRange1;
                    break;
            }

            return _xlActiveWorkSheet.get_Range(myRange1, myRange2);
        }        
        public Excel.Range GetRange(string baseCellName, int rowOffset, int colOffset, enumRangeDirection direction)
        {
            Excel.Range myRange = (Excel.Range)_xlActiveWorkSheet.get_Range(baseCellName, baseCellName);
            Excel.Range myRange1 = (Excel.Range)_xlActiveWorkSheet.Cells[myRange.Row + rowOffset, myRange.Column + colOffset];
            Excel.Range myRange2;
            switch (direction)
            {
                case enumRangeDirection.RightEnd:
                    myRange2 = myRange1.get_End(Excel.XlDirection.xlToRight);
                    break;
                case enumRangeDirection.DownEnd:
                    myRange2 = myRange1.get_End(Excel.XlDirection.xlDown);
                    break;
                default:
                    myRange2 = myRange1;
                    break;
            }

            return _xlActiveWorkSheet.get_Range(myRange1, myRange2);
        }
        public Excel.Range GetMatrix(int firstCellRowIndex, int firstCellColIndex)
        {
            Excel.Range myRange1 = GetCell(firstCellRowIndex, firstCellColIndex);
            Excel.Range myRange2 = myRange1.get_End(Excel.XlDirection.xlToRight).get_End(Excel.XlDirection.xlDown);
            return _xlActiveWorkSheet.get_Range(myRange1, myRange2);
        }
        public Excel.Range GetMatrix(int firstCellRowIndex, int firstCellColIndex, int lastCellColIndex)
        {
            Excel.Range myRange1 = GetCell(firstCellRowIndex, firstCellColIndex);
            Excel.Range myRange2 = GetCell(RowIndex(firstCellRowIndex, firstCellColIndex, enumRangeDirection.DownEnd), lastCellColIndex);
            return _xlActiveWorkSheet.get_Range(myRange1, myRange2);
        }
        
        public int RowIndex(string cellName)
        {
            return GetCell(cellName).Row;
        }
        public int RowIndex(string baseCellName, int rowOffset, int colOffset)
        {
            return GetCell(RowIndex(baseCellName) + rowOffset, ColIndex(baseCellName) + colOffset).Row;  
        }
        public int RowIndex(int rowIndex, int colIndex)
        {
            return GetCell(rowIndex, colIndex).Row;
        }
        public int RowIndex(string cellName, enumRangeDirection direction)
        {
            return GetCell(RowIndex(cellName), ColIndex(cellName), direction).Row;
        }
        public int RowIndex(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            return GetCell(rowIndex, colIndex, direction).Row;
        }
        public int LastRowWithData()
        {
            return _xlActiveWorkSheet.UsedRange.Rows.Count;
        }
        public int LastColumnWithData()
        {
            return _xlActiveWorkSheet.UsedRange.Columns.Count;
        }
        public int LastRowWithDataInThisColumn(int colIndex)
        {
            return GetCell(_xlActiveWorkSheet.Rows.Count, colIndex, enumRangeDirection.UpEnd).Row;
        }
        public int LastColWithDataInThisRow(int rowIndex)
        {
            return GetCell(rowIndex, _xlActiveWorkSheet.Columns.Count, enumRangeDirection.LeftEnd).Column;
        }

        public int ColIndex(string cellName)
        {
            return GetCell(cellName).Column;
        }
        public int ColIndex(string baseCellName, int rowOffset, int colOffset)
        {
            return GetCell(RowIndex(baseCellName) + rowOffset, ColIndex(baseCellName) + colOffset).Column;
        }
        public int ColIndex(int rowIndex, int colIndex)
        {
            return GetCell(rowIndex, colIndex).Column;
        }
        public int ColIndex(string cellName, enumRangeDirection direction)
        {
            return GetCell(RowIndex(cellName), ColIndex(cellName), direction).Column;
        }
        public int ColIndex(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            return GetCell(rowIndex, colIndex, direction).Column;
        }

        public object ReadCellFromActiveSheet(int rowIndex, int colIndex)
        {
            // get the cell
            Excel.Range myRange = GetCell(rowIndex, colIndex);
            if (myRange.Value2 == null)
                return "";
            else
                return myRange.Value2;
        }
        public object ReadCellFromActiveSheet(string cellName)
        {
            // get the sheet
            Excel.Range myRange = GetCell(cellName);
            if (myRange.Value2 == null)
                return "";
            else
                return myRange.Value2;
            
        }
        public object ReadCellFromActiveSheet(string baseCellName, int rowOffset, int colOffset)
        {
            // get the cells
            Excel.Range myRange = GetCell(baseCellName);
            Excel.Range myTargetRange = GetCell(myRange.Row + rowOffset, myRange.Column + colOffset);            
            if (myTargetRange.Value2 == null)
                return "";
            else
                return myTargetRange.Value2;           
        }

        public double[,] ReadMatrixFromActiveSheet(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex)
        {
            double[,] values;

            Excel.Range myRange = GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex);
            if (myRange.Cells.Count == 1)
            {
                values = new double[1, 1];
                values[0, 0] = (double)myRange.Cells.Value;
            }
            else
            {
                Array arrValues = (Array)myRange.Cells.Value;
                values = new double[rowEndIndex - rowStartIndex + 1, colEndIndex - colStartIndex + 1];
                Array.Copy(arrValues, values, arrValues.Length);
            }
 
            return values;
        }
        public double[,] ReadMatrixFromActiveSheet(int firstCellRowIndex, int firstCellColIndex)
        {
            Excel.Range myRange = GetMatrix(firstCellRowIndex, firstCellColIndex);
            Array arrValues = (Array)myRange.Cells.Value;

            double[,] values = new double[arrValues.GetLength(0), arrValues.GetLength(1)];
            Array.Copy(arrValues, values, arrValues.Length);
            
            return values;
        }
        public double[,] ReadMatrixFromActiveSheet(int firstCellRowIndex, int firstCellColIndex, int lastCellColIndex)
        {
            Excel.Range myRange = GetMatrix(firstCellRowIndex, firstCellColIndex, lastCellColIndex);
            Array arrValues = (Array)myRange.Cells.Value;

            double[,] values = new double[arrValues.GetLength(0), arrValues.GetLength(1)];
            Array.Copy(arrValues, values, arrValues.Length);

            return values;
        } 

        public Array ReadObjectMatrixFromActiveSheet(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex)
        {
            Excel.Range myRange = GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex);
            Array arrValues = (Array)myRange.Cells.Value;
            return arrValues;
        } 

        public double[] ReadRangeFromActiveSheet(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            // get the range
            Excel.Range myRange = GetRange(rowIndex, colIndex, direction);

            int i = 0;
            double[] values = new double[myRange.Count];
            foreach (Excel.Range thisRange in myRange.Cells)
            {
                if (thisRange.Value2 == null)
                    values[i] = 0;
                else
                    values[i] = (double)thisRange.Value2;

                i += 1;
            }
            return values;            
        }        
        public double[] ReadRangeFromActiveSheet(string baseCellName, int rowOffset, int colOffset, enumRangeDirection direction)
        {
            // get the range
            Excel.Range myRange = GetRange(baseCellName,rowOffset,colOffset,direction);

            int i = 0;
            double[] values = new double[myRange.Count];
            foreach (Excel.Range thisRange in myRange.Cells)
            {
                if (thisRange.Value2 == null)
                    values[i] = 0;
                else
                    values[i] = (double)thisRange.Value2;

                i += 1;
            }
            return values;
        }
        public double[] ReadRangeFromActiveSheet(int firstRowIndex, int colIndex, int lastRowIndex)
        {
            // get the range
            Excel.Range myRange = GetRange(firstRowIndex, colIndex, lastRowIndex, colIndex);

            int i = 0;
            double[] values = new double[myRange.Count];
            foreach (Excel.Range thisRange in myRange.Cells)
            {
                if (thisRange.Value2 == null)
                    values[i] = 0;
                else
                    values[i] = (double)thisRange.Value2;

                i += 1;
            }
            return values;
        }   

        public string[] ReadStringRangeFromActiveSheet(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            // get the range
            Excel.Range myRange = GetRange(rowIndex, colIndex, direction);

            int i = 0;
            string[] values = new string[myRange.Count];
            foreach (Excel.Range thisRange in myRange.Cells)
            {
                values[i] = (string)thisRange.Value2;
                i += 1;
            }
            return values;
        }
        public string[] ReadStringRangeFromActiveSheet(string baseCellName, int rowOffset, int colOffset, enumRangeDirection direction)
        {
            // get the range
            Excel.Range myRange = GetRange(baseCellName,rowOffset,colOffset, direction);

            int i = 0;
            string[] values = new string[myRange.Count];
            foreach (Excel.Range thisRange in myRange.Cells)
            {
                values[i] = (string)thisRange.Value2;
                i += 1;
            }
            return values;
        }

        public void WriteToCell(object value, int rowIndex, int colIndex)
        {
            _xlActiveWorkSheet.Cells[rowIndex, colIndex] = value;
        }
        public void WriteToCell(object value, string cellName)
        {
            Excel.Range myRange = (Excel.Range)_xlActiveWorkSheet.get_Range(cellName, cellName);
            myRange.Value2 = value;
        }
        public void WriteToCell(object value, string cellName, int rowOffset, int colOffset)
        {            
            //Excel.Range myRange = GetCell(cellName);            
            //Excel.Range myRange1 = (Excel.Range)_xlActiveWorkSheet.Cells[rowOffset(cellName) + rowOffset, ColIndex(cellName) + colOffset];
            //myRange1.Value2 = value;
            GetCell(RowIndex(cellName) + rowOffset, ColIndex(cellName) + colOffset).Value2 = value;    
        }

        public void WriteToRow(double[] values, string cellName, int rowOffset, int colOffset )
        {
            // convert row into matrix
            double[,] rowValues = new double[1, values.Length];
            for (int i = 0; i < values.Length; i++)
                rowValues[0, i] = values[i];

            int rowStartIndex = RowIndex(cellName) + rowOffset;
            int colStartIndex = ColIndex(cellName) + colOffset;

            GetRange(rowStartIndex, colStartIndex, rowStartIndex, colStartIndex + values.Length - 1).Value2 = rowValues;

        }
        public void WriteToRow(double[] values, int rowStartIndex, int colStartIndex)
        {
            // convert row into matrix
            double[,] rowValues = new double[1, values.Length];
            for (int i = 0; i < values.Length; i++)
                rowValues[0, i] = values[i];

            GetRange(rowStartIndex, colStartIndex, rowStartIndex, colStartIndex + values.Length - 1).Value2 = rowValues;
        }
        public void WriteToRow(string[] values, string cellName, int rowOffset, int colOffset)
        {
            // convert row into matrix
            string[,] rowValues = new string[1, values.Length];
            for (int i = 0; i < values.Length; i++)
                rowValues[0, i] = values[i];

            int rowStartIndex = RowIndex(cellName) + rowOffset;
            int colStartIndex = ColIndex(cellName) + colOffset;

            GetRange(rowStartIndex, colStartIndex, rowStartIndex, colStartIndex + values.Length - 1).Value2 = rowValues;
        }
        public void WriteToRow(string[] values, int rowStartIndex, int colStartIndex)
        {
            // convert row into matrix
            string[,] rowValues = new string[1, values.Length];
            for (int i = 0; i < values.Length; i++)
                rowValues[0, i] = values[i];

            GetRange(rowStartIndex, colStartIndex, rowStartIndex, colStartIndex + values.Length - 1).Value2 = rowValues;
        }

        public void WriteToColumn(double[] values, string cellName, int rowOffset, int colOffset)
        {
            // convert column into matrix
            double[,] colValues = new double[values.Length, 1];
            for (int i = 0; i < values.Length; i++)
                colValues[i, 0] = values[i];

            int rowStartIndex = RowIndex(cellName) + rowOffset;
            int colStartIndex = ColIndex(cellName) + colOffset;

            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.Length - 1, colStartIndex).Value2 = colValues;
        }
        public void WriteToColumn(int[] values, string cellName, int rowOffset, int colOffset)
        {
            // convert column into matrix
            double[,] colValues = new double[values.Length, 1];
            for (int i = 0; i < values.Length; i++)
                colValues[i, 0] = values[i];

            int rowStartIndex = RowIndex(cellName) + rowOffset;
            int colStartIndex = ColIndex(cellName) + colOffset;

            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.Length - 1, colStartIndex).Value2 = colValues;
        }
        public void WriteToColumn(double[] values, int rowStartIndex, int colStartIndex)
        {
            // convert column into matrix
            double[,] colValues = new double[values.Length, 1];
            for (int j = 0; j < values.Length; j++)
                colValues[j, 0] = values[j];

            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.Length - 1, colStartIndex).Value2 = colValues;
        }
        public void WriteToColumn(int[] values, int rowStartIndex, int colStartIndex)
        {
            // convert column into matrix
            int[,] colValues = new int[values.Length, 1];
            for (int j = 0; j < values.Length; j++)
                colValues[j, 0] = values[j];

            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.Length - 1, colStartIndex).Value2 = colValues;
        }
        public void WriteToColumn(string[] values, int rowStartIndex, int colStartIndex)
        {
            // convert column into matrix
            string[,] colValues = new string[values.Length, 1];
            for (int j = 0; j < values.Length; j++)
                colValues[j, 0] = values[j];

            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.Length - 1, colStartIndex).Value2 = colValues;
        }
        public void WriteToMatrix(double[,] values, int rowStartIndex, int colStartIndex)
        {
            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.GetLength(0) - 1, colStartIndex + values.GetLength(1)).Value2 = values;
        }
        public void WriteToMatrix(int[,] values, int rowStartIndex, int colStartIndex)
        {
            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.GetLength(0) - 1, colStartIndex + values.GetLength(1)).Value2 = values;
        }
        public void WriteToMatrix(string[,] values, int rowStartIndex, int colStartIndex)
        {
            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.GetLength(0) - 1, colStartIndex + values.GetLength(1)).Value2 = values;
        }
        public void WriteToMatrix(double[,] values, string cellName, int rowOffset, int colOffset)
        {
            int rowStartIndex = RowIndex(cellName) + rowOffset;
            int colStartIndex = ColIndex(cellName) + colOffset;

            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.GetLength(0) - 1, colStartIndex + values.GetLength(1)).Value2 = values;
        }
        public void WriteToMatrix(int[,] values, string cellName, int rowOffset, int colOffset)
        {
            int rowStartIndex = RowIndex(cellName) + rowOffset;
            int colStartIndex = ColIndex(cellName) + colOffset;

            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.GetLength(0) - 1, colStartIndex + values.GetLength(1)).Value2 = values;
        }
        public void WriteToMatrix(string[,] values, string cellName, int rowOffset, int colOffset)
        {
            int rowStartIndex = RowIndex(cellName) + rowOffset;
            int colStartIndex = ColIndex(cellName) + colOffset;

            GetRange(rowStartIndex, colStartIndex, rowStartIndex + values.GetLength(0) - 1, colStartIndex + values.GetLength(1)).Value2 = values;
        }
        public void ClearAll(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex)
        {
            GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex).Clear();
        }
        public void ClearAll(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            GetRange(rowIndex, colIndex, direction).Clear();
        }
        public void ClearContent(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            GetRange(rowIndex, colIndex, direction).ClearContents();
        }
        public void ClearContent(int rowStartIndex, int colStartIndex, int rowEndIndex,int colEndIndex)
        {
            GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex).ClearContents();            
        }
        public void ClearBorders(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex)
        {
            Excel.Borders thisBorders = GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex).Borders;

            thisBorders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            thisBorders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            thisBorders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            thisBorders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            thisBorders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            thisBorders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone;

            //GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex)
            //    .Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex)
            //    .Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex)
            //    .Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex)
            //    .Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex)
            //    .Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex)
            //    .Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }

        // formatting
        public void FormatNumber(int rowIndex, int colIndex, string formatString)
        {
            Excel.Range thisRange = GetCell(rowIndex, colIndex);
            FormatNumber(thisRange, formatString);
        }
        public void FormatNumber(string cellName, int rowOffset, int colOffset, string formatString)
        {
            Excel.Range thisRange = GetCell(RowIndex(cellName) + rowOffset, ColIndex(cellName) + colOffset);
            FormatNumber(thisRange, formatString);
        }
        public void FormatNumber(string cellName, string formatString)
        {
            Excel.Range thisRange = GetCell(cellName);
            FormatNumber(thisRange, formatString);
        }
        public void FormatNumber(string baseCellName, int rowOffset, int colOffset, enumRangeDirection direction, string formatString)
        {
            Excel.Range thisRange = GetRange(RowIndex(baseCellName) + rowOffset, ColIndex(baseCellName) + colOffset, direction);
            FormatNumber(thisRange, formatString);
        }
        public void FormatNumber(int rowIndex, int colIndex, enumRangeDirection direction, string formatString)
        {
            Excel.Range thisRange = GetRange(rowIndex, colIndex, direction);
            FormatNumber(thisRange, formatString);
        }
        public void FormatNumber(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex, string formatString)
        {
            Excel.Range thisRange = GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex);
            FormatNumber(thisRange, formatString);
        }
        private void FormatNumber(Excel.Range thisRange, string formatString)
        {
            thisRange.NumberFormat = formatString;
        }

        public void ClearFormat(int rowIndex, int colIndex)
        {
            GetCell(rowIndex, colIndex).ClearFormats();
        }
        public void ClearFormat(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            GetRange(rowIndex, colIndex, direction).ClearFormats();
        }
        public void ClearFormat(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex)
        {
            GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex).ClearFormats();
        }
        public void MakeBold(int rowIndex, int colIndex)
        {            
            GetCell(rowIndex, colIndex).Font.Bold = true;            
        }
        public void MakeBold(string cellName)
        {
            GetCell(cellName).Font.Bold = true;
        }
        public void MakeBold(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            GetRange(rowIndex, colIndex, direction).Font.Bold = true;
        }
        public void MakeBold(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex)
        {
            GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex).Font.Bold = true;
        }

        public void Align(int rowIndex, int colIndex, enumAlignment alighment)
        {
            GetCell(rowIndex, colIndex).HorizontalAlignment = AlignmentIndex(alighment);
        }
        public void Align(string cellName, enumAlignment alighment)
        {
            GetCell(cellName).HorizontalAlignment = AlignmentIndex(alighment);
        }
        public void Align(string cellName, int rowOffset, int colOffset, enumAlignment alighment)
        {
            GetCell(RowIndex(cellName, rowOffset, colOffset), ColIndex(cellName, rowOffset, colOffset)).HorizontalAlignment = AlignmentIndex(alighment);
        }
        public void Align(string baseCellName, int rowOffset, int colOffset, enumRangeDirection direction, enumAlignment alighment)
        {
            Excel.Range thisRange = GetRange(RowIndex(baseCellName) + rowOffset, ColIndex(baseCellName) + colOffset, direction);
            thisRange.HorizontalAlignment = AlignmentIndex(alighment);
        }
        public void Align(int rowIndex, int colIndex, enumRangeDirection direction, enumAlignment alighment)
        {
            GetRange(rowIndex, colIndex, direction).HorizontalAlignment = AlignmentIndex(alighment);
        }
        public void Align(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex, enumAlignment alighment)
        {
            GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex).HorizontalAlignment = AlignmentIndex(alighment);
        }
        public void AlignAMatrix(int rowStartIndex, int colStartIndex, enumAlignment alighment)
        {
            GetMatrix(rowStartIndex, colStartIndex).HorizontalAlignment = AlignmentIndex(alighment);
        }

        public void AddABorder(int rowIndex, int colIndex, enumBorder border)
        {
            Excel.Range thisRange = GetCell(rowIndex, colIndex);
            AddABorder(thisRange, border);
        }
        public void AddABorder(string cellName, enumBorder border)
        {
            Excel.Range thisRange = GetCell(cellName);
            AddABorder(thisRange, border);
        }
        public void AddABorder(int rowIndex, int colIndex, enumRangeDirection direction, enumBorder border)
        {
            Excel.Range thisRange = GetRange(rowIndex, colIndex, direction);
            AddABorder(thisRange, border);
        }
        public void AddABorder(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex, enumBorder border)
        {
            Excel.Range thisRange = GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex);
            AddABorder(thisRange, border);
        }
        public void AddABorder(string baseCellName, int rowOffset, int colOffset, enumRangeDirection direction, enumBorder border)
        {
            Excel.Range thisRange = GetRange(RowIndex(baseCellName) + rowOffset, ColIndex(baseCellName) + colOffset, direction);
            AddABorder(thisRange, border);
        }
        private void AddABorder(Excel.Range thisRange, enumBorder border)
        {
            Excel.XlBordersIndex borderIndex = BorderIndex(border);
            thisRange.Borders.get_Item(borderIndex).LineStyle = Excel.XlLineStyle.xlContinuous;
            thisRange.Borders.get_Item(borderIndex).Weight = Excel.XlBorderWeight.xlThin;            
        }

        public void WrapText(int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex)
        {
            Excel.Range thisRange = GetRange(rowStartIndex, colStartIndex, rowEndIndex, colEndIndex);
            WrapText(thisRange);
        }
        public void WrapText(string baseCellName, int rowOffset, int colOffset, enumRangeDirection direction)
        {
            Excel.Range thisRange = GetRange(RowIndex(baseCellName) + rowOffset, ColIndex(baseCellName) + colOffset, direction);
            WrapText(thisRange);
        }
        public void WrapText(int rowIndex, int colIndex, enumRangeDirection direction)
        {
            Excel.Range thisRange = GetRange(rowIndex, colIndex, direction);
            WrapText(thisRange);
        }
        private void WrapText(Excel.Range thisRange)
        {
            thisRange.WrapText = true;            
        }

        // charts
        public void UpdateScatterChartValues(string chartName, enumAxis axis, int rowStartIndex, int colStartIndex, int rowEndIndex, int colEndIndex)
        {  
            //Excel.Chart myChart = (Excel.Chart)_xlActiveWorkSheet.ChartObjects(chartName);
            Excel.ChartObject myChart = (Excel.ChartObject)_xlActiveWorkSheet.ChartObjects(chartName);
            Excel.Chart oChart = (Excel.Chart)myChart.Chart;
            
            Excel.SeriesCollection mySeriesCollection = (Excel.SeriesCollection)oChart.SeriesCollection(Type.Missing);
            Excel.Series mySeries = (Excel.Series)mySeriesCollection.Item(1);

            Excel.Range myRange1 = (Excel.Range)_xlActiveWorkSheet.Cells[rowStartIndex, colStartIndex];
            Excel.Range myRange2 = (Excel.Range)_xlActiveWorkSheet.Cells[rowEndIndex, colEndIndex];

            switch (axis)
            {
                case enumAxis.x:
                    mySeries.XValues = (Excel.Range)_xlActiveWorkSheet.get_Range(myRange1, myRange2);                    
                    break;
                case enumAxis.y:
                    mySeries.Values = (Excel.Range)_xlActiveWorkSheet.get_Range(myRange1, myRange2);
                    break;
            }

            //Excel.Axis axis = (Excel.Axis)oChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            //axis.MinimumScale = 0;
            //axis.MaximumScale = 1;

        }

        // PRIVATE
        private Excel.XlBordersIndex BorderIndex(enumBorder border)
        {
            Excel.XlBordersIndex borderIndex = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom;

            switch (border)
            {
                case enumBorder.Top:
                    borderIndex = Excel.XlBordersIndex.xlEdgeTop;
                    break;
                case enumBorder.Bottom:
                    borderIndex = Excel.XlBordersIndex.xlEdgeBottom;
                    break;
                case enumBorder.Left:
                    borderIndex = Excel.XlBordersIndex.xlEdgeLeft;
                    break;
                case enumBorder.Right:
                    borderIndex = Excel.XlBordersIndex.xlEdgeRight;
                    break;                    
            }
            return borderIndex;
        }

        private Excel.XlHAlign AlignmentIndex(enumAlignment alignment)
        {
            Excel.XlHAlign alignmentIndex = Excel.XlHAlign.xlHAlignCenter;

            switch (alignment)
            {
                case enumAlignment.Center:
                    alignmentIndex = Excel.XlHAlign.xlHAlignCenter;
                    break;
                case enumAlignment.Right:
                    alignmentIndex = Excel.XlHAlign.xlHAlignRight;
                    break;
                case enumAlignment.Left:
                    alignmentIndex = Excel.XlHAlign.xlHAlignLeft;
                    break;
            }
            return alignmentIndex;
        }

    }
}
