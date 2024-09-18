using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JExcelExtension;

//This class contains static versions of ExcelTools
public static class ExcelExtension
{
    private static ExcelTools tool = new ExcelTools();

    public static string getCoord(int x, int y) => tool.getExcelCoord(x, y);
    public static string intToColumnLettering(this int value) => tool.getExcelColumnLetters(value);
    public static string arrayToString(this string[] array, string seperator) => tool.arrayToString(array, seperator);
    public static string[] trimArray(this string[] array) => tool.trimArray(array);
    public static string[] summarizeArray(this string[] array) => tool.summarizeArray(array);
    public static string[] splitArrayByString(this string[] array, string splitBy) => tool.splitArrayByString(array, splitBy);
    public static string[] mergeWith(this string[] mergeInto, string[] mergeWith, string spacing) => tool.mergeArrays(mergeInto, mergeWith, spacing);
    public static string[] replaceString(this string[] array, string target, string newString) => tool.replaceString(array, target, newString);
    public static string[,] toColumn(this string[] array) => tool.arrayToColumnFormat(array);
    public static string[,] toRow(this string[] array) => tool.arrayToRowFormat(array);
    public static string[,] splitArrayByString2D(this string[] array, string splitBy) => tool.splitArrayByStringTwo(array, splitBy);
}