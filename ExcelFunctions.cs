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

    //THIS FUNCTION IS REQUIRED TO BE RUN AFTER YOU ARE FINISHED WITH "ExcelFunctions" OTHERWISE EXCEL PROCESS WON'T CLOSE
    public void Free()
    {
        sheetRange.Free();
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
    //Range version
    public void insertStrings(string[,] strings, int x, int y)
    {
        localStringOne = ExcelExtension.getCoord(x, y);
        localStringTwo = ExcelExtension.getCoord(x + strings.GetLength(1) - 1, y + strings.GetLength(0) - 1);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = strings;
    }

    //Inserts a string array into multiple excel columns and rows of choosing. Will override used cells. Size depends on string length.
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

    public void emptySheet()
    {
        localStringOne = ExcelExtension.getCoord(0, 1);
        localStringTwo = ExcelExtension.getCoord(sheetRange.UsedColumnCount, sheetRange.UsedRowsCount);

        sheetRange.setRange(localStringOne, localStringTwo);

        sheetRange.Value2 = "";
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

