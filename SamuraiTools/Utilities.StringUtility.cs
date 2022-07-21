using System;
using System.Collections.Generic;
using System.Text;

namespace SamuraiTools.Utilities
{
    public static class StringUtility
    {
        //credit to user Antoine on stackoverflow
        public static string[] SplitCsv(string line)
        {
            List<string> result = new List<string>();
            StringBuilder currentString = new StringBuilder(string.Empty);
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++) // For each character
            {
                if (line[i] == '\"' && (i == 0 || line[i - 1] != '\\')) // Quotes are closing or opening
                    inQuotes = !inQuotes;
                else if (line[i] == ',') // Comma
                {
                    if (!inQuotes) // If not in quotes, end of current string, add it to result
                    {
                        result.Add(currentString.ToString());
                        currentString.Clear();
                    }
                    else
                        currentString.Append(line[i]); // If in quotes, just add it
                }
                else if (!char.IsControl(line[i])) // Add any other character to current string as long as it isn't a control character
                    currentString.Append(line[i]);
            }
            result.Add(currentString.ToString());
            return result.ToArray(); // Return array of all strings
        }
    }
}
