using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSVParserTest
{
    public class CSVParser
    {
        public enum ModeType
        {
            All,
            SkipHeader,
            HeaderAsKey
        }

        ModeType _mode = ModeType.HeaderAsKey;
        public ModeType Mode
        {
            get
            {
                return _mode;
            }
            set
            {
                _mode = value;
            }
        }

        char _delimiter = ';';
        public char Delimiter
        {
            get
            {
                return _delimiter;
            }
            set
            {
                _delimiter = value;
            }
        }

        public List<Dictionary<string, string>> Parse(string content)
        {
            // split content in parts and determine number of columns
            List<string> parts = new List<string>();

            StringBuilder current = new StringBuilder();

            int numberOfColumns = -1;
            int numberOfRows = 0;
            bool rowsHaveEqualNumberOfColumns = true;

            bool betweenDoubleQuotes = false;
            for (int i = 0; i < content.Length; i++)
            {
                char c = content[i];
                if (!betweenDoubleQuotes && (c == Delimiter || c == '\n'))
                {
                    parts.Add(current.ToString());
                    current.Clear();

                    if (c == '\n')
                    {
                        numberOfRows++;
                        if (numberOfColumns == -1)
                        {
                            numberOfColumns = parts.Count;
                        }
                        else if (numberOfRows * numberOfColumns != parts.Count)
                        {
                            rowsHaveEqualNumberOfColumns = false;
                        }
                    }

                    continue;
                }

                if (!betweenDoubleQuotes && c == '"')
                {
                    betweenDoubleQuotes = true;
                    continue;
                }

                if (betweenDoubleQuotes && c == '"')
                {
                    if (i < content.Length - 1 && content[i + 1] == '"')
                    {
                        i++;
                        current.Append(c);
                        continue;
                    }
                    else
                    {
                        betweenDoubleQuotes = false;
                        continue;
                    }
                }

                current.Append(c);
            }


            // some checks...
            if (content.LastOrDefault() != '\n')
            {
                throw new Exception("doesn't end with a newline");
            }

            if (!rowsHaveEqualNumberOfColumns)
            {
                throw new Exception("rows don't have the same amount of columns");
            }


            // create output format
            List<string> keys;
            if (Mode == ModeType.HeaderAsKey)
            {
                keys = parts.Take(numberOfColumns).ToList();
            }
            else
            {
                keys = Enumerable.Range(1, numberOfColumns).Select(c => GetExcelColumnName(c)).ToList();
            }

            List<string> dataParts;
            if (Mode != ModeType.All)
            {
                dataParts = parts.Skip(numberOfColumns).ToList();
                numberOfRows--;
            }
            else
            {
                dataParts = parts;
            }


            List<Dictionary<string, string>> result = new List<Dictionary<string, string>>();
            for (int row = 0; row < numberOfRows; row++)
            {
                var rowData = dataParts.GetRange(row * numberOfColumns, numberOfColumns);

                Dictionary<string, string> rowDict = new Dictionary<string, string>();
                for (int column = 0; column < numberOfColumns; column++)
                {
                    rowDict.Add(keys[column], rowData[column]);
                }

                result.Add(rowDict);
            }

            return result;
        }

        // http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string csv = "header 1;header 2;header 3\ndata 1;data 2;data 3\n\"data;4\";\"data\"\"5\";\"data\n6\"\n";

            var parser = new CSVParser();

            parser.Mode = CSVParser.ModeType.All;
            var parsedAll = parser.Parse(csv);
            Console.WriteLine(parsedAll[1]["B"]); // prints "data 2"

            parser.Mode = CSVParser.ModeType.SkipHeader;
            var parsedSkipHeader = parser.Parse(csv);
            Console.WriteLine(parsedSkipHeader[0]["B"]); // prints "data 2"

            parser.Mode = CSVParser.ModeType.HeaderAsKey;
            var parsedHeaderAsKey = parser.Parse(csv);
            Console.WriteLine(parsedHeaderAsKey[0]["header 2"]); // prints "data 2"

            Console.ReadKey();
        }
    }
}
