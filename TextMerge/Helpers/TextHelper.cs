using System;
using System.Data;
using System.IO;
using System.Windows.Media;

namespace TextMerge.Helpers
{
    public class TextHelper
    {
        public static DataTable GetTextData(string file)
        {
            var table = new DataTable();
            table.Columns.Add("Column1");
            table.Columns.Add(Path.GetFileNameWithoutExtension(file));
            using (var f = new StreamReader(file))
            {
                while (!f.EndOfStream)
                {
                    var readLine = f.ReadLine();
                    if (readLine != null)
                    {
                        string[] data = readLine.Split(new[]{','});
                        var row = table.NewRow(); 
                        row[0] = data[0];
                        row[1] = data[1];
                        table.Rows.Add(row);
                    }
                }
            }
            return table;
        }

        public static DataTable GetTextData(string file,bool isAll)
        {

            var table = new DataTable();
            try
            {
                if (isAll)
                {
                    table.Columns.Add("0",typeof(double));
                }
                table.Columns.Add(Directory.GetParent(file).Name,typeof(double));
                #region 添加列头
                var head = table.NewRow();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    head[i] = table.Columns[i].ColumnName;
                }
                table.Rows.Add(head); 
                #endregion
                using (var f = new StreamReader(file))
                {
                    while (!f.EndOfStream)
                    {
                        var line = f.ReadLine();
                        if (line != null)
                        {
                            string[] data = line.Split(new[] { ',' });
                            var row = table.NewRow();
                            if (isAll)
                            {
                                row[0] = Convert.ToDouble(data[0]);
                                row[1] = Convert.ToDouble(data[1]);
                            }
                            else
                            {
                                row[0] = Convert.ToDouble(data[1]);
                            }
                            table.Rows.Add(row);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
            return table;
        }
    }
}
