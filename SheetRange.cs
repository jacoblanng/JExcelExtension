﻿using Excel = Microsoft.Office.Interop.Excel;

namespace JExcelExtension;

//A custom struct to organize excel range and sheets to shorten code and simple
public struct SheetRange
{
    public SheetRange() {}
    public SheetRange(ref Excel.Worksheet sheetRef)
    {
        sheet = sheetRef;
        range = sheet.get_Range("a1");
    }

    //THIS FUNCTION IS REQUIRED TO BE RUN AFTER YOU ARE FINISHED WITH "SheetRange" OTHERWISE EXCEL PROCESS WON'T CLOSE
    public void Free()
    {
        if (sheet != null)
        {
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            sheet = null;
        }
        if(range != null)
        {
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
        }
    }

    public void import(ref Excel.Worksheet sheetRef)
    {
        sheet = sheetRef;
        range = sheet.get_Range("a1");
    }

    public Excel.Range range;
    public Excel.Worksheet sheet;

    public void setRange(Excel.Range _range)
    {
        range = sheet.get_Range(_range);
    }
    public void setRange(SheetRange _sheetRange)
    {
        range = sheet.get_Range(_sheetRange.range);
    }
    public void setRange(int aX, int aY)
    {
        range = sheet.get_Range(ExcelExtension.getCoord(aX, aY));
    }
    public void setRange(int aX, int aY, int bX, int bY)
    {
        range = sheet.get_Range(ExcelExtension.getCoord(aX, aY), ExcelExtension.getCoord(bX, bY));
    }
    public void setRange(string x)
    {
        range = sheet.get_Range(x);
    }
    public void setRange(string x, string y)
    {
        range = sheet.get_Range(x, y);
    }

    //WARNING DO NOT SET THIS VARIABLE TO A DECIMAL TYPE AS THAT WILL ALWAYS TRHOW AN ERROR
    public dynamic Value2
    {
        get => range.Value2;
        set => range.Value2 = value;
    }

    public dynamic ColorIndex
    {
        get => range.Interior.ColorIndex;
        set => range.Interior.ColorIndex = value;
    }
    public int Count => range.Count;
    public int Column => range.Column;
    public int Row => range.Row;
    public int UsedRowCount => sheet.UsedRange.Rows.Count;
    public int UsedColumnCount => sheet.UsedRange.Columns.Count;
}
