namespace JExcelExtension;

//This class contains a functions that does not require excel but is good for converting and translating data for excel use
public class ExcelTools
{
    //Easy to use array of the english alphabet that is used for columns in excel. Starts at 0. Ends at 24.
    private readonly string[] letterMap =
    {
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
    };

    //Function for getting the first or two letters of a column based on int up to a max of 701. Cannot get coloumns past "ZZ"
    public string getExcelColumnLetters(int column)
    {
        //Prevents numbers too large to be handled. 
        if (column >= 702)
            return "";

        //If more than one letter is needed
        if (column >= 26)
        {
            string letterColumn = "";

            int loopAmount = -1;
            int lastLetter = column;

            //Letter to check how many times input can go through "letterMap"
            for (int i = column; i >= 26; i -= 26)
            {
                loopAmount++;
                lastLetter -= 26;
            }

            //First letter is based on how many times the previous loop went through "letterMap"
            if (loopAmount <= 25)
                letterColumn = letterColumn + letterMap[loopAmount];
            //Adds last letter normally if it isn't less than zero
            if (lastLetter >= 0)
                letterColumn = letterColumn + letterMap[lastLetter];

            //Returns two letters representing the correct coloumn based on input
            return letterColumn;
        }

        //Returns one letter based on "letterMap"
        return letterMap[column];
    }

    //Function for turning a set of numbers into excel readable coordinates
    public string getExcelCoord(int x, int y)
    {
        return getExcelColumnLetters(x) + y.ToString();
    }

    //Function for turning array into one string with seperator
    public string arrayToString(string[] array, string seperator)
    {
        if (array.Length == 0)
            return "";

        if (array.Length <= 1)
        if (array[0] == null || array[0] == "")
            return "";

        string result = "";

        result += array[0];
        
        for(int i = 1; i <= array.Length - 1; i++)
        {
            result += seperator + array[i];
        }

        return result;
    }

    //Function that returns a string array made from two string arrays with the ability to add a string in between to seperate
    public string[] mergeArrays(string[] firstArray, string[] secondArray, string betweenString)
    {
        int arrayLength = 0;

        if (firstArray.Length >= secondArray.Length)
        {
            arrayLength = firstArray.Length;
        }
        else
        {
            arrayLength = secondArray.Length;
        }

        string[] result = new string[arrayLength];

        for (int i = 0; i <= arrayLength - 1; i++)
        {
            if (firstArray[i] != null && secondArray != null)
            {
                result[i] = firstArray[i] + betweenString + secondArray[i];
                continue;
            }
            if (firstArray[i] != null)
            {
                result[i] = firstArray[i];
                continue;
            }
            if (secondArray[i] != null)
            {
                result[i] = secondArray[i];
                continue;
            }

            result[i] = "";
        }

        return result;
    }

    //Function that merges two arrays in parallel into a two dimensional array
    public T[,] mergeArraysParallel<T>(T[] firstArray, T[] secondArray)
    {
        if(firstArray == null || secondArray == null)
            return new T[0,0];

        int length = firstArray.Length;

        if(secondArray.Length >= length)
            length = secondArray.Length;

        T[,] result = new T[length, 2];

        for (int i = 0; i < length; i++)
        {
            if(i < firstArray.Length)
                result[i, 0] = firstArray[i];

            if(i < secondArray.Length)
                result[i, 1] = secondArray[i];
        }

        return result;
    }

    //Function that merges an array in parallel of a two dimensional array at the end of the 2nd dimension
    public T[,] mergeArraysParallel<T>(T[,] firstArray, T[] secondArray)
    {
        if (firstArray == null || secondArray == null)
            return new T[0, 0];

        int firstLength = firstArray.GetLength(0);
        int secondLength = firstArray.GetLength(1);

        if (secondArray.Length >= firstLength)
            firstLength = secondArray.Length;

        T[,] result = new T[firstLength, 2];
        for (int l = 0; l < secondLength; l++)
        {
            for (int i = 0; i < firstLength; i++)
            {
                if (i < firstArray.GetLength(0))
                    result[i, l] = firstArray[i, l];
            }
        }

        for(int i = 0; i < firstLength; i++)
        {
            if (i < secondArray.Length)
                result[i, secondLength] = secondArray[i];
        }

        return result;
    }

    //Splits a string array by splitting each string into multiple strings making the array longer via singular string
    public string[] splitArrayByString(string[] array, string splitBy)
    {
        int x = array.Length - 1;

        foreach (string s in array)
        {
            if (s == null || s == "")
                continue;

            x += s.Split(splitBy).Length;
        }

        string[] result = new string[x + 1];

        int i = 0;
        foreach (string s in array)
        {
            if (s == null || s == "")
                continue;

            string[] splitArrayByString = s.Split(splitBy);

            foreach (string l in splitArrayByString)
            {
                result[i] = l;
                i++;
            }
        }

        return result;
    }

    //Splits a string array by splitting each string into multiple strings making the array longer via char array
    public string[] splitArrayByChars(string[] array, char[] splitBy)
    {
        int x = array.Length - 1;

        foreach (string s in array)
        {
            if (s == null || s == "")
                continue;

            x += s.Split(splitBy).Length;
        }

        string[] result = new string[x + 1];

        int i = 0;
        foreach (string s in array)
        {
            if (s == null || s == "")
                continue;

            string[] splitArrayByString = s.Split(splitBy);

            foreach (string l in splitArrayByString)
            {
                result[i] = l;
                i++;
            }
        }

        return result;
    }

    //Splits a single dimension string array into a two dimension string array by splitting each string into multiple strings via singular string
    public string[,] splitArrayByStringTwo(string[] array, string splitBy)
    {
        int x = array.Length - 1;
        int y = 0;

        foreach (string s in array)
        {
            if (s == null || s == "")
                continue;

            int amountOfSplits = s.Split(splitBy).Length;

            if (amountOfSplits >= y)
                y = amountOfSplits;
        }

        string[,] result = new string[x + 1, y];

        for (int i = 0; i <= x; i++)
        {
            string[] splitString = array[i].Split(splitBy);


            for (int l = 0; l < y; l++)
            {
                if (splitString.Length - 1 < l || splitString[l] == null)
                    continue;

                result[i, l] = splitString[l];
            }
        }

        return result;
    }

    //Splits a single dimension string array into a two dimension string array by splitting each string into multiple strings via char array
    public string[,] splitArrayByCharsTwo(string[] array, char[] splitBy)
    {
        int x = array.Length - 1;
        int y = 0;

        foreach (string s in array)
        {
            if(s == null || s == "")
                continue;

            int amountOfSplits = s.Split(splitBy).Length;

            if (amountOfSplits >= y)
                y = amountOfSplits;
        }

        string[,] result = new string[x + 1, y];

        for (int i = 0; i <= x; i++)
        {
            if (array[i] == "" || array[i] == null)
                continue;

            string[] splitString = array[i].Split(splitBy);

            for (int l = 0; l < y; l++)
            {
                if (splitString.Length - 1 < l || splitString[l] == null)
                    continue;

                result[i, l] = splitString[l];
            }
        }

        return result;
    }

    //Trims an array by removing all white space strings in array ensuring a shorter array but no data lost
    public string[] trimArray(string[] strings)
    {
        int amountToTrim = 0;

        for (int i = strings.Length - 1; i >= 0; i--)
        {
            if(strings[i] != null)
            if(strings[i].Trim() != "")
                continue;

            amountToTrim++;
        }

        string[] result = new string[strings.Length - amountToTrim];

        int resultPos = 0;

        for (int i = 0; i < strings.Length; i++)
        {
            if(strings[i] == null)
                continue;
            if (strings[i].Trim() == "")
                continue;

            result[resultPos] = strings[i];
            resultPos++;
        }

        return result;
    }

    //Shortens a string array down to only the unique strings getting rid of duplicate strings
    public T[] summarizeArray<T>(T[] types)
    {
        T[] result = new T[types.Distinct().Count()];

        result = types.Distinct().ToArray();

        return result;
    }

    //Replaces all strings that matches the "targetString" with "newString" throughout the array
    public string[] replaceString(string[] array, string targetString, string newString)
    {
        string[] result = new string[array.Length];

        int i = 0;
        foreach (string s in array)
        {
            if (s == null || s == "")
            {
                i++;
                continue;
            }
            result[i] = s.Replace(targetString, newString);

            i++;
        }

        return result;
    }

    //Replaces all chars that matches the "targetChar" with "newChar" throughout the array
    public string[] replaceChar(string[] array, char targetChar, char newChar)
    {
        string[] result = new string[array.Length];

        int i = 0;
        foreach (string s in array)
        {
            if (s == null || s == "") {
                i++;
                continue;
            }
            result[i] = s.Replace(targetChar, newChar);

            i++;
        }

        return result;
    }

    //Turns a normal one dimensional array of types into a type array that is fit for inserting into excel columns
    public T[,] typesToColumnFormat<T>(T[] types)
    {
        T[,] result = new T[types.Length, 1];

        for(int i = 0; i < types.Length; i++)
        {
            if(types[i] != null)
                result[i, 0] = types[i];
        }

        return result;
    }

    //Only used for LegacyFunctions
    //Turns a normal one dimensional string array into a string array that is fit for inserting into excel columns
    public string[,] stringsToColumnFormat(string[] strings)
    {
        string[,] result = new string[strings.Length, 1];

        int i = 0;
        foreach (string s in strings)
        {
            if (s != null && s != "")
                result[i, 0] = s;

            i++;
        }

        return result;
    }

    //Turns a normal one dimensional type array into a type array that is fit for inserting into excel rows
    public T[,] typesToRowFormat<T>(T[] types)
    {
        T[,] result = new T[1, types.Length];

        for(int i = 0; i < types.Length; i++)
        {
            if (types[i] != null)
                result[0, i] = types[i];
        }

        return result;
    }

    //Only used for LegacyFunctions
    //Turns a normal one dimensional string array into a string array that is fit for inserting into excel rows
    public string[,] stringsToRowFormat(string[] strings)
    {
        string[,] result = new string[1, strings.Length];

        int i = 0;
        foreach (string s in strings)
        {
            if (s != null && s != "")
                result[0, i] = s;

            i++;
        }

        return result;
    }
}

