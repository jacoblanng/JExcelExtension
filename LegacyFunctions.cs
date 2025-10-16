using Excel = Microsoft.Office.Interop.Excel;

namespace JExcelExtension;

//LEGACY CLASS FOR CODE MADE PRIOR TO 1.018
//NOT RECOMMENDED
//This class is purely for code that used ExcelFunctions before version 1.018 so that only the initialized class needs to be renamed and
//everything thereafter should work as normal. All code here will no longer be updated and is therefore deprecated. It is recommended to
//use the normal ExcelFunctions as its should be more efficient, flexible and isn't locked to strings for data.

public class LegacyFunctions
{
    //Resources for use in class
    private string? localStringOne;
    private string? localStringTwo;

    private int localIntOne;
    private int localIntTwo;

    //Resources for use in and out of class
    public SheetRange sheetRange;

    public LegacyFunctions() { }
    public LegacyFunctions(ref Excel.Worksheet sheetRef)
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
    //MAIN FUNCTIONS
    //

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

        if (sheetRange.UsedRowCount <= startRow - sheetRange.UsedRowCount)
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
            return new string[0, 0];

        if (startColumn > sheetRange.UsedColumnCount)
            return new string[0, 0];

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

    //END OF MAIN FUNCTIONS
    //MISCELLANEOUS FUNCTIONS
    //

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
    public void emptySheet()
    {
        sheetRange.setRange(0, 1, sheetRange.UsedColumnCount, sheetRange.UsedRowCount);

        sheetRange.Value2 = "";
        colorRange(0);
    }
    //Clears sheet from start "startColumn" and "startRow", anything before these will not be cleared.
    public void emptySheet(int startColumn, int startRow)
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
