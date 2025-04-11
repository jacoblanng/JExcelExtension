using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JExcelExtension;

//This class contains static versions of ExcelTools
public static class ExcelExtension
{
    private static ExcelTools tool = new();

    public static string getCoord(int x, int y) => tool.getExcelCoord(x, y);
    public static string intToColumnLettering(this int value) => tool.getExcelColumnLetters(value);
    public static string arrayToString(this string[] array, string seperator) => tool.arrayToString(array, seperator);
    public static string[] trimArray(this string[] array) => tool.trimArray(array);
    public static string[] summarizeArray(this string[] array) => tool.summarizeArray(array);
    public static string[] splitArrayByString(this string[] array, string splitBy) => tool.splitArrayByString(array, splitBy);
    public static string[] splitArrayByChars(this string[] array, char[] splitBy) => tool.splitArrayByChars(array, splitBy);
    public static string[] mergeWith(this string[] mergeInto, string[] mergeWith, string spacing) => tool.mergeArrays(mergeInto, mergeWith, spacing);
    public static string[] replaceString(this string[] array, string target, string newString) => tool.replaceString(array, target, newString);
    public static string[] replaceChar(this string[] array, char targetChar, char newChar) => tool.replaceChar(array, targetChar, newChar);
    public static T[,] toTypeColumn<T>(this T[] array) => tool.typesToColumnFormat(array);
    public static string[,] toStringColumn(this string[] array) => tool.stringsToColumnFormat(array);
    public static T[,] toTypeRow<T>(this T[] array) => tool.typesToRowFormat(array);
    public static string[,] toStringRow(this string[] array) => tool.stringsToRowFormat(array);
    public static string[,] splitArrayByString2D(this string[] array, string splitBy) => tool.splitArrayByStringTwo(array, splitBy);
    public static string[,] splitArrayByChars2D(this string[] array, char[] splitBy) => tool.splitArrayByCharsTwo(array, splitBy);
}