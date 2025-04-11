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

    //WARNING
    //THIS FUNCTION IS REQUIRED TO BE RUN AFTER YOU ARE FINISHED WITH "ExcelFunctions" OTHERWISE EXCEL PROCESS WON'T CLOSE
    //
    public void Free()
    {
        sheetRange.Free();
    }

    //
    //REPEATED FUNCTIONS
    //

    //Loops through all cells in "range" and adds it to appropriate index in "result"
    private void getRangeTypes<T> (ref T[] array) {
        int i = 0;
        foreach (Excel.Range c in sheetRange.range)
        {
            if (c.Value2 != null)
                array[i] = Convert.ChangeType(c.Value2, typeof(T));

            i++;
        }
    }

    //Loops through all cells in "range" and adds it to appropriate index in "result" but for 2D arrays
    private void getRangeTypes<T>(ref T[,] array)
    {
        foreach (Excel.Range c in sheetRange.range)
        {
            if (c.Value2 == null)
                continue;

            array[c.Column - 1, c.Row - 1] = Convert.ChangeType(c.Value2, typeof(T));
        }
    }

    //END OF REPEATED FUNCTIONS
    //THESE FUNCTIONS HERE AFTER WILL LATER BECOME RECOMMENDED AND STABLE FUNCTIONS
    //EXPERIMENTAL FUNCTIONS HERE AFTER

    //Gets a cell value and converts it to the specified type
    //Range version
    public T cellToType<T>(Excel.Range _range)
    {
        sheetRange.setRange(_range);

        if (sheetRange.Value2 == null)
            return default(T);

        T result = Convert.ChangeType(sheetRange.Value2, typeof(T));

        return result;
    }
    //SheetRange version
    public T cellToType<T>(SheetRange _sheetRange)
    {
        sheetRange.setRange(_sheetRange);

        if (sheetRange.Value2 == null)
            return default;

        T result = Convert.ChangeType(sheetRange.Value2, typeof(T));

        return result;
    }
    //Number version
    public T cellToType<T>(int x, int y)
    {
        sheetRange.setRange(x, y);

        if (sheetRange.Value2 == null)
            return default;

        T result = Convert.ChangeType(sheetRange.Value2, typeof(T));

        return result;
    }

    //Gets a type array of a column from "startRow" to "endRow"
    public T[] columnToTypes<T>(int column, int startRow, int endRow)
    {
        //Checks if range would be valid
        if (endRow <= startRow)
            return new T[0];

        if (startRow > sheetRange.UsedRowCount)
            return new T[0];

        //Sets range
        sheetRange.setRange(column, startRow, column, endRow);

        //Makes type array with the size of all cells in "range"
        T[] result = new T[sheetRange.Count];

        if (sheetRange.UsedRowCount <= startRow - endRow)
            return result;

        //Gets all values in range and transfers them to result
        getRangeTypes(ref result);

        return result;
    }

    //Gets a type array of a whole selected column from "startRow"
    public T[] columnToTypes<T>(int column, int startRow)
    {
        //Checks if range would be valid
        if (startRow > sheetRange.UsedRowCount)
            return new T[0];

        //Sets range
        sheetRange.setRange(column, startRow, column, sheetRange.UsedRowCount);

        //Makes type array with the size of all cells in "range"
        T[] result = new T[sheetRange.Count];

        //Gets all values in range and transfers them to result
        getRangeTypes(ref result);

        return result;
    }

    //Gets a type array of a whole selected column with all rows
    public T[] columnToTypes<T>(int column)
    {
        //Sets range
        sheetRange.setRange(column, 1, column, sheetRange.UsedRowCount);

        //Makes type array with the size of all cells in "range"
        T[] result = new T[sheetRange.Count];

        //Gets all values in range and transfers them to result
        getRangeTypes(ref result);

        return result;
    }


    //Gets a type array of a row from "startColumn" to "endColumn"
    public T[] rowToTypes<T>(int row, int startColumn, int endColumn)
    {
       //Checks if range would be valid
        if (endColumn <= startColumn)
            return new T[0];
        if (startColumn > sheetRange.UsedColumnCount)
            return new T[0];

        //Sets range
        sheetRange.setRange(startColumn, row, endColumn, row);

        //Makes type array with the size of all cells in "range"
        T[] result = new T[sheetRange.Count];

        //Gets all values in range and transfers them to result
        getRangeTypes(ref result);

        return result;
    }

    //Gets a string array of a whole selected row from "startColumn"
    public T[] rowToTypes<T>(int row, int startColumn)
    {
        //Checks if range would be valid
        if (startColumn > sheetRange.UsedColumnCount)
            return new T[0];

        //Sets range
        sheetRange.setRange(startColumn, row, sheetRange.UsedColumnCount, row);

        //Makes type array with the size of all cells in "range"
        T[] result = new T[sheetRange.Count];

        //Gets all values in range and transfers them to result
        getRangeTypes(ref result);

        return result;
    }
    //Gets a type array of a whole selected row with all columns
    public T[] rowToTypes<T>(int row)
    {
        //Sets range
        sheetRange.setRange(0, row, sheetRange.UsedColumnCount, row);

        //Makes type array with the size of all cells in "range"
        T[] result = new T[sheetRange.Count];

        //Gets all values in range and transfers them to result
        getRangeTypes(ref result);

        return result;
    }

    //Gets a type array of all values in between the selected points
    public T[,] rangeToTypes<T>(int startColumn, int startRow, int endColumn, int endRow)
    {
        //Checks if range would be valid
        if (endRow <= startRow)
            return new T[0, 0];
        if(endColumn <= startColumn)
            return new T[0, 0];

        //Sets range
        sheetRange.setRange(startColumn, startRow, endColumn, endRow);

        //Makes type array with the size of all cells in "range"
        T[,] result = new T[endColumn - startColumn, endRow - startRow];

        //Gets all values in range and transfers them to result
        getRangeTypes(ref result);

        return result;
    }

    //Gets type array from of the selected point to the end of the sheet
    public T[,] rangeToTypes<T>(int startColumn, int startRow)
    {
        //Checks if range would be valid
        if (startRow > sheetRange.UsedRowCount)
            return new T[0, 0];
        if (startColumn > sheetRange.UsedColumnCount)
            return new T[0, 0];

        //Sets range
        sheetRange.setRange(startColumn, startRow, sheetRange.UsedColumnCount, sheetRange.UsedRowCount);

        T[,] result = new T[sheetRange.UsedColumnCount - startColumn, sheetRange.UsedRowCount - startRow];

        //Gets all values in range and transfers them to result
        getRangeTypes(ref result);

        return result;
    }

    //Gets string array from the whole sheet
    public T[,] rangeToTypes<T>()
    {
        //Sets range
        sheetRange.setRange(0, 1, sheetRange.UsedColumnCount, sheetRange.UsedRowCount);

        T[,] result = new T[sheetRange.UsedColumnCount, sheetRange.UsedRowCount];

        //Gets all values in range and transfers them to result
        getRangeTypes(ref result);

        return result;
    }

    //Inserts value into specified cell 
    //Range version
    public void insertCellType<T> (T value, Excel.Range _range)
    {
        sheetRange.setRange(_range);

        if (value == null)
            return;

        sheetRange.Value2 = value;
    }
    //SheetRange version
    public void insertCellType<T>(T value, SheetRange _sheetRange)
    {
        sheetRange.setRange(_sheetRange);

        if (value == null)
            return;

        sheetRange.Value2 = value;
    }
    //Number version
    public void insertCellType<T>(T value, int x, int y)
    {
        sheetRange.setRange(x, y);

        if (value == null)
            return;

        sheetRange.Value2 = value;
    }

    //Inserts "insertType" at all rows specified by "insertAt" in specified column
    public void insertTypeAt<T>(int column, int[] insertAt, T insertType)
    {
        if(insertType == null)
            return;

        foreach (int i in insertAt)
        {
            sheetRange.setRange(column, i);

            sheetRange.Value2 = insertType;
        }
    }

    //Inserts "types" into an excel column of choosing. Size depends on array length.
    //Range version
    public void insertColumnTypes<T>(T[] types, Excel.Range _range)
    {
        sheetRange.setRange(_range);

        sheetRange.Value2 = types.toTypeColumn();
    }

    //SheetRange version
    public void insertColumnTypes<T>(T[] types, SheetRange _sheetRange)
    {
        sheetRange.setRange(_sheetRange);

        sheetRange.Value2 = types.toTypeColumn();
    }

    //Number version
    public void insertColumnTypes<T>(T[] types, int x, int y)
    {
        sheetRange.setRange(x, y, x, y + types.Length - 1);

        sheetRange.Value2 = types.toTypeColumn();
    }

    //Inserts a "types" into an excel row of choosing. Size depends on array length.
    //Range version
    public void insertRowTypes<T>(T[] types, Excel.Range _range)
    {
        sheetRange.setRange(_range);

        sheetRange.Value2 = types.toTypeRow();
    }

    //SheetRange version
    public void insertRowTypes<T>(T[] types, SheetRange _sheetRange)
    {
        sheetRange.setRange(_sheetRange);

        sheetRange.Value2 = types.toTypeRow();
    }

    //Number version
    public void insertRowTypes<T>(T[] types, int x, int y)
    {
        sheetRange.setRange(x, y, x + types.Length - 1, y);

        sheetRange.Value2 = types.toTypeRow();
    }

    //Inserts a 2D array of types straight into specified position in excel
    //Number version
    public void insertTypes<T>(T[,] types, int x, int y)
    {
        sheetRange.setRange(x, y, x + types.GetLength(1) - 1, y + types.GetLength(0) - 1);

        sheetRange.Value2 = types;
    }

    //Range version
    public void insertTypes<T>(T[,] types, Excel.Range _range)
    {
        sheetRange.setRange(_range.Column - 1, _range.Row, _range.Column - 1 + types.GetLength(1) - 1, _range.Row + types.GetLength(0) - 1);

        sheetRange.Value2 = types;
    }

    //SheetRange version
    public void insertTypes<T>(T[,] types, SheetRange _sheetRange)
    {
        sheetRange.setRange(_sheetRange.Column - 1, _sheetRange.Row, _sheetRange.Column - 1 + types.GetLength(1) - 1, _sheetRange.Row + types.GetLength(0) - 1);

        sheetRange.Value2 = types;
    }

    //END OF EXPERIMENTAL FUNCTIONS
    //EVERYTHING HERE AFTER WILL LATER BE MOVED TO A LEGACY CLASS
    //STABLE AND RECOMMENDED FUNCTIONS HERE AFTER

    //Gets string from cell. If range will always take the first column and row
    //Range version
    public string cellToString(Excel.Range _range)
    {
        sheetRange.setRange(_range);

        if (sheetRange.Value2 == null)
            return "";

        return sheetRange.Value2.ToString();
    }

    //SheetRange version
    public string cellToString(SheetRange _sheetRange)
    {
        sheetRange.setRange(_sheetRange);

        if (sheetRange.Value2 == null)
            return "";

        return sheetRange.Value2.ToString();
    }

    //Number version
    public string cellToString(int x, int y)
    {
        sheetRange.setRange(x, y);

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

        if (sheetRange.UsedRowCount <= startRow - endRow)
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
        if (startRow > sheetRange.UsedRowCount)
            return new string[0];

        //Makes the appropriate strings
        localStringOne = column.intToColumnLettering() + startRow.ToString();
        localStringTwo = column.intToColumnLettering() + sheetRange.UsedRowCount.ToString();

        //Makes range based on the previous two strings
        sheetRange.setRange(localStringOne, localStringTwo);

        //Makes string array with the size of all cells in "range"
        string[] result = new string[sheetRange.Count];

        if(sheetRange.UsedRowCount <= startRow - sheetRange.UsedRowCount)
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
        localStringTwo = column.intToColumnLettering() + sheetRange.UsedRowCount.ToString();

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

        string[,] result = new string[sheetRange.UsedColumnCount, sheetRange.UsedRowCount];

        foreach (Excel.Range c in sheetRange.range)
        {
            if (c.Value2 != null)
            {
                result[c.Column - 1, c.Row - 1] = c.Value2.ToString();
            }
            else if (c.Column <= sheetRange.UsedColumnCount && c.Row <= sheetRange.UsedRowCount)
            {
                result[c.Column - 1, c.Row - 1] = "";
            }
        }

        return result;
    }

    //Gets string array from of the selected point to the end of the sheet
    public string[,] rangeToStrings(int startColumn, int startRow)
    {
        if (startRow > sheetRange.UsedRowCount)
            return new string[0,0];

        if (startColumn > sheetRange.UsedColumnCount)
            return new string[0,0];

        localStringOne = startColumn.intToColumnLettering() + startRow.ToString();
        localStringTwo = sheetRange.UsedColumnCount.intToColumnLettering() + sheetRange.UsedRowCount.ToString();

        sheetRange.setRange(localStringOne, localStringTwo);

        string[,] result = new string[sheetRange.UsedColumnCount, sheetRange.UsedRowCount];

        foreach (Excel.Range c in sheetRange.range)
        {
            if (c.Value2 != null)
            {
                result[c.Column - 1, c.Row - 1] = c.Value2.ToString();
            }
            else if (c.Column <= sheetRange.UsedColumnCount && c.Row <= sheetRange.UsedRowCount)
            {
                result[c.Column - 1, c.Row - 1] = "";
            }
        }

        return result;
    }

    //Gets string array from the whole sheet
    public string[,] rangeToStrings()
    {
        localStringOne = 0.intToColumnLettering() + "1";
        localStringTwo = sheetRange.UsedColumnCount.intToColumnLettering() + sheetRange.UsedRowCount.ToString();

        sheetRange.setRange(localStringOne, localStringTwo);

        string[,] result = new string[sheetRange.UsedColumnCount, sheetRange.UsedRowCount];

        foreach (Excel.Range c in sheetRange.range)
        {
            if (c.Value2 != null)
            {
                result[c.Column - 1, c.Row - 1] = c.Value2.ToString();
            }
            else if (c.Column <= sheetRange.UsedColumnCount && c.Row <= sheetRange.UsedRowCount)
            {
                result[c.Column - 1, c.Row - 1] = "";
            }
        }

        return result;
    }

    //Inserts string into specified cell 
    //Range version
    public void insertCellString(string str, Excel.Range _range)
    {
        if (str == "" || str == null)
            return;

        localIntOne = _range.Column - 1;
        localIntTwo = _range.Row;

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
    public void insertColumnStrings(string[] strings, Excel.Range _range)
    {
        localIntOne = _range.Column - 1;
        localIntTwo = _range.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne, localIntTwo + strings.Length - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toStringColumn();
    }

    //SheetRange version
    public void insertColumnStrings(string[] strings, SheetRange _sheetRange)
    {
        localIntOne = _sheetRange.Column - 1;
        localIntTwo = _sheetRange.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne, localIntTwo + strings.Length - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toStringColumn();
    }

    //Number version
    public void insertColumnStrings(string[] strings, int x, int y)
    {
        localStringOne = ExcelExtension.getCoord(x, y);
        localStringTwo = ExcelExtension.getCoord(x, y + strings.Length - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toStringColumn();
    }

    //Inserts a "strings" into an excel row of choosing. Will override used cells. Size depends on string length.
    //Range version
    public void insertRowStrings(string[] strings, Excel.Range _range)
    {
        localIntOne = _range.Column - 1;
        localIntTwo = _range.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne + strings.Length - 1, localIntTwo);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toStringRow();
    }

    //SheetRange version
    public void insertRowStrings(string[] strings, SheetRange _sheetRange)
    {
        localIntOne = _sheetRange.Column - 1;
        localIntTwo = _sheetRange.Row;

        localStringOne = ExcelExtension.getCoord(localIntOne, localIntTwo);
        localStringTwo = ExcelExtension.getCoord(localIntOne + strings.Length - 1, localIntTwo);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toStringRow();
    }

    //Number version
    public void insertRowStrings(string[] strings, int x, int y)
    {
        localStringOne = ExcelExtension.getCoord(x, y);
        localStringTwo = ExcelExtension.getCoord(x + strings.Length - 1, y);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings.toStringRow();
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
    public void insertStrings(string[,] strings, Excel.Range _range)
    {
        localIntOne = _range.Column - 1;
        localIntTwo = _range.Row;

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

    //END OF STABLE FUNCTIONS
    //ANYTHING HERE AFTER WILL NOT BE MOVED TO LEGACY BUT WILL BE UPDATED OR ADDED ON TO
    //FUNCTIONS HERE AFTER ARE MISCELLANEOUS FUNCTIONS

    //Checks if the specified cell is empty. If range will always take first column and row
    //Range version
    public bool isCellEmpty(Excel.Range _range)
    {
        sheetRange.setRange(_range);

        if (sheetRange.Value2 == null)
            return true;
        return false;
    }
    //SheetRange version
    public bool isCellEmpty(SheetRange _sheetRange)
    {
        sheetRange.setRange(_sheetRange);

        if (sheetRange.Value2 == null)
            return true;
        return false;
    }
    //Number version
    public bool isCellEmpty(int x, int y)
    {
        sheetRange.setRange(x, y);

        if (sheetRange.Value2 == null)
            return true;
        return false;
    }

    //Checks if the specified range is empty
    //Range version
    public bool isRangeEmpty(Excel.Range range)
    {
        sheetRange.setRange(range);

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
        sheetRange.setRange(_sheetRange);

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
        sheetRange.setRange(x1, y1, x2, y2);

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
    public void colorCell(int colorIndex, Excel.Range _range)
    {
        sheetRange.setRange(_range);

        sheetRange.ColorIndex = colorIndex;
    }

    //SheetRange version
    public void colorCell(int colorIndex, SheetRange _sheetRange)
    {
        sheetRange.setRange(_sheetRange);

        sheetRange.ColorIndex = colorIndex;
    }

    //Number version
    public void colorCell(int colorIndex, int x, int y)
    {
        sheetRange.setRange(x, y);

        sheetRange.ColorIndex = colorIndex;
    }
    //Range previously set version
    public void colorCell(int colorIndex)
    {
        sheetRange.ColorIndex = colorIndex;
    }

    //Colors the specified range. Uses Excel color index
    //Range version
    public void colorRange(int colorIndex, Excel.Range _range)
    {
        sheetRange.setRange(_range);

        sheetRange.ColorIndex = colorIndex;
    }

    //SheetRange version
    public void colorRange(int colorIndex, SheetRange _sheetRange)
    {
        sheetRange.setRange(_sheetRange);

        sheetRange.ColorIndex = colorIndex;
    }

    //Number version
    public void colorRange(int colorIndex, int x1, int y1, int x2, int y2)
    {
        sheetRange.setRange(x1, y1, x2, y2);

        sheetRange.ColorIndex = colorIndex;
    }
    //Range previously set version
    public void colorRange(int colorIndex)
    {
        sheetRange.ColorIndex = colorIndex;
    }

    //Clears sheet
    //Whole sheet
    public void clearSheet()
    {
        sheetRange.setRange(0, 1, sheetRange.UsedColumnCount, sheetRange.UsedRowCount);

        sheetRange.Value2 = "";
        colorRange(0);
    }
    //Clears sheet from start "startColumn" and "startRow", anything before these will not be cleared.
    public void clearSheet(int startColumn, int startRow)
    {
        sheetRange.setRange(startColumn, startRow, sheetRange.UsedColumnCount, sheetRange.UsedRowCount);

        sheetRange.Value2 = "";
        colorRange(0);
    }

    //Empties column of values
    //Whole column
    public void emptyColumn(int column)
    {
        sheetRange.setRange(column, 1, column, sheetRange.UsedRowCount);

        sheetRange.Value2 = "";
    }
    //Empties column from "startRow", anything before this will not be cleared.
    public void emptyColumn(int column, int startRow)
    {
        sheetRange.setRange(column, startRow, column, sheetRange.UsedRowCount);

        sheetRange.Value2 = "";
    }
    //Empties row of values
    //Whole row
    public void emptyRow(int row)
    {
        sheetRange.setRange(0, row, sheetRange.UsedColumnCount, row);

        sheetRange.Value2 = "";
    }
    //Empties row from start "startColumn", anything before this will not be cleared.
    public void emptyRow(int startColumn, int row)
    {
        sheetRange.setRange(startColumn, row, sheetRange.UsedColumnCount, row);

        sheetRange.Value2 = "";
    }
}

