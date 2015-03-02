using System.Data;

namespace Converter
{
    static class CSVReader
    {
        public static DataTable GetDataTableFromCsvFile(string filePath)
        {
            const string extension = ".csv";
            var headerDelimiter = new[] { @";" };
            var lineDelimiter = new[] { @""";""" };

            if (System.IO.Path.GetExtension(filePath) != extension) return null;

            var table = new DataTable();

//            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read,FileShare.ReadWrite))
//            using (var reader = new StreamReader(stream, Encoding.Default))
//            {
//                var doFirstLine = false;
//                while (!reader.EndOfStream)
//                {
//                    //Create header
//                    if (!doFirstLine)
//                    {
//                        var firstLine = reader.ReadLine();
//                        if (String.IsNullOrEmpty(firstLine)) return null;
//                        var columnNames = firstLine.Split(headerDelimiter, StringSplitOptions.None);
//
//
//                        foreach (var name in columnNames)
//                        {
//                            var column = String.IsNullOrEmpty(name) ? new DataColumn() : new DataColumn(name);
//                            column.AllowDBNull = true;
//                            table.Columns.Add(column);
//                        }
//
//                        doFirstLine = true;
//                    }
//                    //fillRows
//                    else
//                    {
//                        var line = reader.ReadLine();
//                        if (!String.IsNullOrEmpty(line))
//                        {
//                            var cells = line.Split(lineDelimiter, StringSplitOptions.None);
//                            if (cells.Count() != table.Columns.Count)
//                            {
//                                //well Our plan is fucked
//                            }
//                            var row = table.NewRow();
//                            cells = cells.Select(x => x.Trim('"')).ToArray();
//                            row.ItemArray = cells;
////                            for (var i = 0; i < cells.GetUpperBound(0); i++)
////                            {
////                                row[i] = cells[i].Replace("\"", "");
////                            }
//                            table.Rows.Add(row);
//                        }
//                    }
//                }
            return table;
        }
    }
}
