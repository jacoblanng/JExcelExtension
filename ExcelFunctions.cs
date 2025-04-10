using System.Data.Common;
using System.Numerics;
using Excel = Microsoft.Office.Interop.Excel;

namespace JExcelExtension;

//Functions for inputting and extracting data in an excel sheet
public class ExcelFunctions
{
    //Resources for use in class
    private string? localStringOne;
    private string? localStringTwo;

    private int localIntOne;
    private int localIntTwo;
    
    //Resources for use in and out of class
    public SheetRange sheetRange;

    public ExcelFunctions() {}
    public ExcelFunctions(ref Excel.Worksheet sheetRef)
    {
        sheetRange = new SheetRange(ref sheetRef);
    }

    //
    //THIS FUNCTION IS REQUIRED TO BE RUN AFTER YOU ARE FINISHED WITH "ExcelFunctions" OTHERWISE EXCEL PROCESS WON'T CLOSE
    //
    public void Free()
    {
        sheetRange.Free();
    }

    //
    //REPEATED FUNCTIONS
    //

    void setRange(Excel.Range range)
    {
        sheetRange.setRange(range);
    }
    void setRange(SheetRange sheetRange)
    {
        sheetRange.setRange(sheetRange.range);
    }
    void setRange(int x1, int y1, int x2, int y2)
    {
        localStringOne = x1.intToColumnLettering() + y1.ToString();
        localStringTwo = x2.intToColumnLettering() + y2.ToString();

        sheetRange.setRange(localStringOne, localStringTwo);
    }

    void setRangeSingular(Excel.Range range)
    {
        localIntOne = range.Column - 1;
        localIntTwo = range.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);

        sheetRange.setRange(localStringOne);
    }
    void setRangeSingular(SheetRange sheetRange)
    {
        localIntOne = sheetRange.Column - 1;
        localIntTwo = sheetRange.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);

        sheetRange.setRange(localStringOne);
    }
    void setRangeSingular(int x, int y)
    {
        localStringOne = ExcelExtension.getCoord(x, y);

        sheetRange.setRange(localStringOne);
    }

    //
    //END OF REPEATED FUNCTIONS
    //

    //
    //THESE FUNCTIONS ARE PURELY EXPERIMENTAL
    //

    //Gets a cell value and converts it to the specified type
    //Range version
    public T cellToType<T>(Excel.Range range)
    {
        setRangeSingular(range);

        if (sheetRange.Value2 == null)
            return default(T);

        T result = Convert.ChangeType(sheetRange.Value2, typeof(T));

        return result;
    }
    //SheetRange version
    public T cellToType<T>(SheetRange _sheetRange)
    {
        setRangeSingular(_sheetRange);

        if (sheetRange.Value2 == null)
            return default;

        T result = Convert.ChangeType(sheetRange.Value2, typeof(T));

        return result;
    }
    //Number version
    public T cellToType<T>(int x, int y)
    {
        setRangeSingular(x, y);

        if (sheetRange.Value2 == null)
            return default;

        T result = Convert.ChangeType(sheetRange.Value2, typeof(T));

        return result;
    }

    public void insertCellType<T> (T value, Excel.Range range)
    {
        setRangeSingular(range);

        if (value == null)
            return;

        sheetRange.Value2 = value;
    }

    public void insertCellType<T>(T value, SheetRange _sheetRange)
    {
        setRangeSingular(_sheetRange);

        if (value == null)
            return;

        sheetRange.Value2 = value;
    }
    public void insertCellType<T>(T value, int x, int y)
    {
        setRangeSingular(x, y);

        if (value == null)
            return;

        sheetRange.Value2 = value;
    }

    //
    //END OF EXPERIMENTAL FUNCTIONS
    //

    //Gets string from cell. If range will always take the first column and row
    //Range version
    public string cellToString(Excel.Range range)
    {
        setRangeSingular(range);

        if (sheetRange.Value2 == null)
            return "";

        return sheetRange.Value2.ToString();
    }

    //SheetRange version
    public string cellToString(SheetRange _sheetRange)
    {
        setRangeSingular(_sheetRange);

        if (sheetRange.Value2 == null)
            return "";

        return sheetRange.Value2.ToString();
    }

    //Number version
    public string cellToString(int x, int y)
    {
        setRangeSingular(x, y);

        if (sheetRange.Value2 == null)
            return "";

        return sheetRange.Value2.ToString();
    }


    //Gets a string array of a column from "startRow" to "endRow"
    public string[] columnToStrings(int column, int startRow, int endRow)
    {
        //Makes the appropriate strings
        localStringOne = column.intToColumnLettering() + startRow.ToString();
        localStringTwo = column.intToColumnLettering() + endRow.ToString();

        //Makes range based on the previous two strings
        sheetRange.setRange(localStringOne, localStringTwo);

        //Makes string array with the size of all cells in "range"
        string[] result = new string[sheetRange.Count];

        if (sheetRange.UsedRowsCount <= startRow - endRow)
            return result;

        //Loops through all cells in "range" and adds it to appropriate index in "result"
        localIntOne = -1;
        foreach (Excel.Range c in sheetRange.range)
        {
            localIntOne++;

            if (c.Value2 == null)
            {
                result[localIntOne] = "";
                continue;
            }

            result[localIntOne] = c.Value2.ToString();
        }

        //Returns result
        return result;
    }

    //Gets a string array of a whole selected column
    public string[] columnToStrings(int column, int startRow)
    {
        if (startRow > sheetRange.UsedRowsCount)
            return new string[0];

        //Makes the appropriate strings
        localStringOne = column.intToColumnLettering() + startRow.ToString();
        localStringTwo = column.intToColumnLettering() + sheetRange.UsedRowsCount.ToString();

        //Makes range based on the previous two strings
        sheetRange.setRange(localStringOne, localStringTwo);

        //Makes string array with the size of all cells in "range"
        string[] result = new string[sheetRange.Count];

        if(sheetRange.UsedRowsCount <= startRow - sheetRange.UsedRowsCount)
            return result;

        //Loops through all cells in "range" and adds it to appropriate index in "result"
        localIntOne = -1;
        foreach (Excel.Range c in sheetRange.range)
        {
            localIntOne++;

            if (c.Value2 == null)
            {
                result[localIntOne] = "";
                continue;
            }

            result[localIntOne] = c.Value2.ToString();
        }

        //Returns result
        return result;
    }

    //Gets a string array of a whole selected column with all rows
    public string[] columnToStrings(int column)
    {
        //Makes the appropriate strings
        localStringOne = column.intToColumnLettering() + "1";
        localStringTwo = column.intToColumnLettering() + sheetRange.UsedRowsCount.ToString();

        //Makes range based on the previous two strings
        sheetRange.setRange(localStringOne, localStringTwo);

        //Makes string array with the size of all cells in "range"
        string[] result = new string[sheetRange.Count];

        //Loops through all cells in "range" and adds it to appropriate index in "result"
        localIntOne = -1;
        foreach (Excel.Range c in sheetRange.range)
        {
            localIntOne++;

            if (c.Value2 == null)
            {
                result[localIntOne] = "";
                continue;
            }

            result[localIntOne] = c.Value2.ToString();
        }

        //Returns result
        return result;
    }

    //Gets a string array of a row from "startColumn" to "endColumn"
    public string[] rowToStrings(int row, int startColumn, int endColumn)
    {
        //Makes the appropriate strings
        localStringOne = startColumn.intToColumnLettering() + row;
        localStringTwo = endColumn.intToColumnLettering() + row;

        //Makes range based on the previous two strings
        sheetRange.setRange(localStringOne, localStringTwo);

        //Makes string array with the size of all cells in "range"
        string[] result = new string[sheetRange.Count];

        //Loops through all cells in "range" and adds it to appropriate index in "result"
        localIntOne = -1;
        foreach (Excel.Range c in sheetRange.range)
        {
            localIntOne++;

            if (c.Value2 == null)
            {
                result[localIntOne] = "";
                continue;
            }

            result[localIntOne] = c.Value2.ToString();
        }

        //Returns result
        return result;
    }

    //Gets a string array of a whole selected row
    public string[] rowToStrings(int row, int startColumn)
    {
        if (startColumn > sheetRange.UsedColumnCount)
            return new string[0];

        //Makes the appropriate strings
        localStringOne = startColumn.intToColumnLettering() + row;
        localStringTwo = (sheetRange.UsedColumnCount - 1).intToColumnLettering() + row;

        //Makes range based on the previous two strings
        sheetRange.setRange(localStringOne, localStringTwo);

        //Makes string array with the size of all cells in "range"
        string[] result = new string[sheetRange.Count];

        //Loops through all cells in "range" and adds it to appropriate index in "result"
        localIntOne = -1;
        foreach (Excel.Range c in sheetRange.range)
        {
            localIntOne++;

            if (c.Value2 == null)
            {
                result[localIntOne] = "";
                continue;
            }

            result[localIntOne] = c.Value2.ToString();
        }

        //Returns result
        return result;
    }

    //Gets a string array of a whole selected row with all columns
    public string[] rowToStrings(int row)
    {
        //Makes the appropriate strings
        localStringOne = 0.intToColumnLettering() + row;
        localStringTwo = (sheetRange.UsedColumnCount - 1).intToColumnLettering() + row;

        //Makes range based on the previous two strings
        sheetRange.setRange(localStringOne, localStringTwo);

        //Makes string array with the size of all cells in "range"
        string[] result = new string[sheetRange.Count];

        //Loops through all cells in "range" and adds it to appropriate index in "result"
        localIntOne = -1;
        foreach (Excel.Range c in sheetRange.range)
        {
            localIntOne++;

            if (c.Value2 == null)
            {
                result[localIntOne] = "";
                continue;
            }

            result[localIntOne] = c.Value2.ToString();
        }

        //Returns result
        return result;
    }
    //Gets a string array of all values in between the selected points
    public string[,] rangeToStrings(int startColumn, int startRow, int endColumn, int endRow)
    {
        localStringOne = startColumn.intToColumnLettering() + startRow.ToString();
        localStringTwo = endColumn.intToColumnLettering() + endRow.ToString();

        sheetRange.setRange(localStringOne, localStringTwo);

        string[,] result = new string[sheetRange.UsedColumnCount, sheetRange.UsedRowsCount];

        foreach (Excel.Range c in sheetRange.range)
        {
            if (c.Value2 != null)
            {
                result[c.Column - 1, c.Row - 1] = c.Value2.ToString();
            }
            else if (c.Column <= sheetRange.UsedColumnCount && c.Row <= sheetRange.UsedRowsCount)
            {
                result[c.Column - 1, c.Row - 1] = "";
            }
        }

        return result;
    }

    //Gets string array from of the selected point to the end of the sheet
    public string[,] rangeToStrings(int startColumn, int startRow)
    {
        if (startRow > sheetRange.UsedRowsCount)
            return new string[0,0];

        if (startColumn > sheetRange.UsedColumnCount)
            return new string[0,0];

        localStringOne = startColumn.intToColumnLettering() + startRow.ToString();
        localStringTwo = sheetRange.UsedColumnCount.intToColumnLettering() + sheetRange.UsedRowsCount.ToString();

        sheetRange.setRange(localStringOne, localStringTwo);

        string[,] result = new string[sheetRange.UsedColumnCount, sheetRange.UsedRowsCount];

        foreach (Excel.Range c in sheetRange.range)
        {
            if (c.Value2 != null)
            {
                result[c.Column - 1, c.Row - 1] = c.Value2.ToString();
            }
            else if (c.Column <= sheetRange.UsedColumnCount && c.Row <= sheetRange.UsedRowsCount)
            {
                result[c.Column - 1, c.Row - 1] = "";
            }
        }

        return result;
    }


    //Inserts string into specified cell 
    //Range version
    public void insertCellString(string str, Excel.Range range)
    {
        if (str == "" || str == null)
            return;

        localIntOne = range.Column - 1;
        localIntTwo = range.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);

        sheetRange.setRange(localStringOne);

        sheetRange.Value2 = str;
    }


    //SheetRange version
    public void insertCellString(string str, SheetRange _sheetRange)
    {
        if (str == "" || str == null)
            return;

        localIntOne = _sheetRange.Column - 1;
        localIntTwo = _sheetRange.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);

        sheetRange.setRange(localStringOne);

        sheetRange.Value2 = str;
    }

    //Number version
    public void insertCellString(string str, int x, int y)
    {
        if (str == "" || str == null)
            return;

        localStringOne = ExcelExtension.getCoord(x, y);

        sheetRange.setRange(localStringOne);

        sheetRange.Value2 = str;
    }

    //Inserts "insertString" at all rows specified by "insertAt" in "column"
    public void insertStringAt(int column, int[] insertAt, string insertString)
    {
        localStringOne = column.intToColumnLettering();

        foreach (int i in insertAt)
        {
            sheetRange.setRange(localStringOne + i.ToString());

            sheetRange.Value2 = insertString;
        }
    }


    //Inserts "strings" into an excel column of choosing. Will override used cells. Size depends on string length.
    //Range version
    public void insertColumnStrings(string[] strings, Excel.Range range)
    {
        localIntOne = range.Column - 1;
        localIntTwo = range.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne, localIntTwo + strings.Length - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toColumn();
    }

    //SheetRange version
    public void insertColumnStrings(string[] strings, SheetRange _sheetRange)
    {
        localIntOne = _sheetRange.Column - 1;
        localIntTwo = _sheetRange.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne, localIntTwo + strings.Length - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toColumn();
    }

    //Number version
    public void insertColumnStrings(string[] strings, int x, int y)
    {
        localStringOne = ExcelExtension.getCoord(x, y);
        localStringTwo = ExcelExtension.getCoord(x, y + strings.Length - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toColumn();
    }

    //Inserts a "strings" into an excel row of choosing. Will override used cells. Size depends on string length.
    //Range version
    public void insertRowStrings(string[] strings, Excel.Range range)
    {
        localIntOne = range.Column - 1;
        localIntTwo = range.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne + strings.Length - 1, localIntTwo);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toRow();
    }

    //SheetRange version
    public void insertRowStrings(string[] strings, SheetRange _sheetRange)
    {
        localIntOne = _sheetRange.Column - 1;
        localIntTwo = _sheetRange.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne + strings.Length - 1, localIntTwo);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toRow();
    }

    //Number version
    public void insertRowStrings(string[] strings, int x, int y)
    {
        localStringOne = ExcelExtension.getCoord(x, y);
        localStringTwo = ExcelExtension.getCoord(x + strings.Length - 1, y);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toRow();
    }

    //Inserts a string array into multiple excel columns and rows of choosing. Will override used cells. Size depends on string length.
    //Number version
    public void insertStrings(string[,] strings, int x, int y)
    {
        localStringOne = ExcelExtension.getCoord(x, y);
        localStringTwo = ExcelExtension.getCoord(x + strings.GetLength(1) - 1, y + strings.GetLength(0) - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings;
    }

    //Range version
    public void insertStrings(string[,] strings, Excel.Range range)
    {
        localIntOne = range.Column - 1;
        localIntTwo = range.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne + strings.GetLength(1) - 1, localIntTwo + strings.GetLength(0) - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings;
    }

    //SheetRange version
    public void insertStrings(string[,] strings, SheetRange _sheetRange)
    {
        localIntOne = _sheetRange.Column - 1;
        localIntTwo = _sheetRange.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne + strings.GetLength(1) - 1, localIntTwo + strings.GetLength(0) - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings;
    }

    //Checks if the specified cell is empty. If range will always take first column and row
    //Range version
    public bool isCellEmpty(Excel.Range range)
    {
        setRangeSingular(range);

        if (sheetRange.Value2 == null)
            return true;
        return false;
    }
    //SheetRange version
    public bool isCellEmpty(SheetRange _sheetRange)
    {
        setRangeSingular(_sheetRange);

        if (sheetRange.Value2 == null)
            return true;
        return false;
    }
    //Number version
    public bool isCellEmpty(int x, int y)
    {
        setRangeSingular(x, y);

        if (sheetRange.Value2 == null)
            return true;
        return false;
    }

    //Checks if the specified range is empty
    //Range version
    public bool isRangeEmpty(Excel.Range range)
    {
        setRange(range);

        foreach(Excel.Range c in sheetRange.range)
        {
            if (c.Value2 != null)
            {
                return false;
            }
        }

        return true;
    }
    //SheetRange version
    public bool isRangeEmpty(SheetRange _sheetRange)
    {
        setRange(_sheetRange);

        foreach (Excel.Range c in sheetRange.range)
        {
            if (c.Value2 != null)
            {
                return false;
            }
        }

        return true;
    }
    //Number version
    public bool isRangeEmpty(int x1, int y1, int x2, int y2)
    {
        setRange(x1, y1, x2, y2);

        foreach (Excel.Range c in sheetRange.range)
        {
            if (c.Value2 != null)
            {
                return false;
            }
        }

        return true;
    }

    //Colors the specified cell. If range will always take first column and row. Uses Excel color index
    //Range version
    public void colorCell(int colorIndex, Excel.Range range)
    {
        localIntOne = range.Column - 1;
        localIntTwo = range.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);

        sheetRange.setRange(localStringOne);

        sheetRange.range.Interior.ColorIndex = colorIndex;
    }

    //SheetRange version
    public void colorCell(int colorIndex, SheetRange _sheetRange)
    {
        localIntOne = _sheetRange.Column - 1;
        localIntTwo = _sheetRange.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);

        sheetRange.setRange(localStringOne);

        sheetRange.range.Interior.ColorIndex = colorIndex;
    }

    //Number version
    public void colorCell(int colorIndex, int x, int y)
    {
        localStringOne = ExcelExtension.getCoord(x, y);

        sheetRange.setRange(localStringOne);

        sheetRange.range.Interior.ColorIndex = colorIndex;
    }


    //Colors the specified range. Uses Excel color index
    //Range version
    public void colorRange(int colorIndex, Excel.Range range)
    {
        localStringOne = ExcelExtension.getCoord(range.Column - 1, range.Row);
        localStringTwo = ExcelExtension.getCoord(range.Columns.Count - 1, range.Rows.Count);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.range.Interior.ColorIndex = colorIndex;
    }

    //SheetRange version
    public void colorRange(int colorIndex, SheetRange _sheetRange)
    {
        localStringOne = ExcelExtension.getCoord(_sheetRange.Column - 1, _sheetRange.Row);
        localStringTwo = ExcelExtension.getCoord(_sheetRange.UsedColumnCount - 1, _sheetRange.UsedRowsCount);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.range.Interior.ColorIndex = colorIndex;
    }

    //Number version
    public void colorRange(int colorIndex, int x1, int y1, int x2, int y2)
    {
        localStringOne = ExcelExtension.getCoord(x1, y1);
        localStringTwo = ExcelExtension.getCoord(x2, y2);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.range.Interior.ColorIndex = colorIndex;
    }


    public void emptySheet()
    {
        localStringOne = ExcelExtension.getCoord(0, 1);
        localStringTwo = ExcelExtension.getCoord(sheetRange.UsedColumnCount, sheetRange.UsedRowsCount);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = "";
        colorRange(0, sheetRange);
    }

    public void emptySheet(int startColumn, int startRow)
    {
        localStringOne = ExcelExtension.getCoord(startColumn, startRow);
        localStringTwo = ExcelExtension.getCoord(sheetRange.UsedColumnCount, sheetRange.UsedRowsCount);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = "";
    }

    public void emptyColumn(int column)
    {
        localStringOne = ExcelExtension.getCoord(column, 1);
        localStringTwo = ExcelExtension.getCoord(column, sheetRange.UsedRowsCount);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = "";
    }

    public void emptyColumn(int column, int startRow)
    {
        localStringOne = ExcelExtension.getCoord(column, startRow);
        localStringTwo = ExcelExtension.getCoord(column, sheetRange.UsedRowsCount);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = "";
    }

    public void emptyRow(int row)
    {
        localStringOne = ExcelExtension.getCoord(0, row);
        localStringTwo = ExcelExtension.getCoord(sheetRange.UsedColumnCount, row);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = "";
    }

    public void emptyRow(int startColumn, int row)
    {
        localStringOne = ExcelExtension.getCoord(startColumn, row);
        localStringTwo = ExcelExtension.getCoord(sheetRange.UsedColumnCount, row);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = "";
    }
}

