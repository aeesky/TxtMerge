using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HPSF;
namespace TextMerge.Helpers
{
    public class ExcelHelper
    {
        private HSSFWorkbook _hssfworkbook;
        public HSSFWorkbook ActiveWorkBook
        {
            get { return _hssfworkbook; }
            set { _hssfworkbook = value; }
        }

        private ISheet _sheet;
        public ISheet ActiveSheet
        {
            get { return _sheet; }
        }

        public string FilePath
        {
            get;
            set;
        }

        /// <summary>
        /// 打开Excel文件
        /// </summary>
        /// <param name="filePath">文件地址</param>
        public bool Open(string filePath)
        {
            try
            {
                FilePath = filePath;
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    _hssfworkbook = new HSSFWorkbook(file);
                    file.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 新建Excel文件
        /// </summary>
        /// <param name="filepath">文件地址</param>
        public void Create(string filepath)
        {
            FilePath = filepath;
            _hssfworkbook = new HSSFWorkbook();
            _hssfworkbook.CreateSheet("sheet1");

            try
            {
                //新建文件
                using (var ms = new MemoryStream())
                {
                    _hssfworkbook.Write(ms);
                    ms.Flush();
                    ms.Position = 0;
                    using (var fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                    {
                        var data = ms.ToArray();
                        fs.Write(data, 0, data.Length);
                        fs.Flush();
                        fs.Close();
                        ms.Close();
                    }
                    //workbook.Dispose();//一般只用写这一个就OK了，他会遍历并释放所有资源，但当前版本有问题所以只释放_sheet
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                throw;
            }
        }

        /// <summary>
        /// 保存当前文档内容
        /// </summary>
        public void Save()
        {
            SaveAs(FilePath);
        }

        /// <summary>
        /// 保存当前文档内容
        /// </summary>
        public void SaveAs(string filename)
        {
            var dir = Path.GetDirectoryName(filename);
            Debug.Assert(dir != null, "保存目录为空");
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            using (var fs = new FileStream(string.IsNullOrEmpty(filename) ? FilePath : filename, FileMode.Create, FileAccess.Write))
            {
                ActiveWorkBook.Write(fs);
                fs.Close();
            }
        }


        /// <summary>
        /// 整表导出的DataTable
        /// </summary>
        /// <param name="index">_sheet页索引</param>
        /// <returns></returns>
        public DataTable ExportToDataTable(int index)
        {
            _sheet = _hssfworkbook.GetSheetAt(index);
            var table = new DataTable();
            var headerRow = ActiveSheet.GetRow(0);//第一行为标题行
            int cellCount = headerRow.LastCellNum;//LastCellNum = PhysicalNumberOfCells
            var rowCount = ActiveSheet.LastRowNum;//LastRowNum = PhysicalNumberOfRows - 1

            //handling header.
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                var column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }
            for (var i = (ActiveSheet.FirstRowNum + 1); i <= rowCount; i++)
            {
                var row = ActiveSheet.GetRow(i);
                var dataRow = table.NewRow();

                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                            dataRow[j] = GetCellValue(row.GetCell(j));
                    }
                }

                table.Rows.Add(dataRow);
            }
            return table;
        }

        /// <summary>
        /// 数据表导入到
        /// </summary>
        /// <param name="datatb"></param>
        /// <param name="shindex">_sheet页代码，从0开始计数</param>
        /// <param name="startcol"></param>
        /// <param name="startrow"></param>
        public void ImportFromTable(DataTable datatb, int shindex, int startcol, int startrow)
        {
            _sheet = _hssfworkbook.GetSheetAt(shindex);
            for (var i = 0; i < datatb.Rows.Count; i++)
            {
                for (var j = 0; j < datatb.Columns.Count; j++)
                {
                    SetCellValue(datatb.Rows[i][j], startcol + j, startrow + i);
                }
            }
            ActiveSheet.ForceFormulaRecalculation = true;
        }

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="shIndex">表索引</param>
        /// <param name="rIndex">插入起始行位置</param>
        /// <param name="rCount">插入行数</param>
        /// <param name="sIndex">新增行格式源所在行，行格式复制</param>
        public void InsertRows(int shIndex, int rIndex, int rCount, int sIndex)
        {
            try
            {
                _sheet = _hssfworkbook.GetSheetAt(shIndex);
                #region 批量移动行
                _sheet.ShiftRows(
                            rIndex,                               //--开始行
                            _sheet.LastRowNum,                    //--结束行
                            rCount,                               //移动行数-负数往下移动
                            true,                                 //是否复制行高
                            false                                 //是否重置行高
                );
                #endregion
                var source = _sheet.GetRow(sIndex); //新增行的格式源

                #region 对批量移动后空出的空行插，创建相应的行，并以rIndex的上一行为格式源(即：rIndex-1的那一行)
                for (var i = rIndex; i < rIndex + rCount - 1; i++)
                {
                    IRow targetRow = _sheet.CreateRow(i + 1);

                    for (int m = source.FirstCellNum; m < source.LastCellNum; m++)
                    {
                        ICell sourceCell = source.GetCell(m);
                        if (sourceCell == null)
                            continue;
                        ICell targetCell = targetRow.CreateCell(m);
                        targetCell.CellStyle = sourceCell.CellStyle;
                        targetCell.SetCellType(sourceCell.CellType);
                    }
                }

                var firstTargetRow = _sheet.GetRow(rIndex);
                for (int m = source.FirstCellNum; m < source.LastCellNum; m++)
                {
                    ICell firstSourceCell = source.GetCell(m);
                    if (firstSourceCell == null)
                        continue;
                    ICell firstTargetCell = firstTargetRow.CreateCell(m);
                    firstTargetCell.CellStyle = firstSourceCell.CellStyle;
                    firstTargetCell.SetCellType(firstSourceCell.CellType);
                }
                #endregion

                ActiveSheet.ForceFormulaRecalculation = true;
                HSSFFormulaEvaluator.EvaluateAllFormulaCells(_hssfworkbook);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                // throw;
            }
        }


        /// <summary>
        /// 获取单元格数据
        /// </summary>
        /// <param name="cindex">列索引</param>
        /// <param name="rindex">行索引</param>
        /// <returns></returns>
        public string GetCellValue(int cindex, int rindex)
        {
            var row = ActiveSheet.GetRow(rindex);
            var cell = row.GetCell(cindex);
            return GetCellValue(cell);
        }

        /// <summary>
        /// 根据Excel列类型获取列的值
        /// </summary>
        /// <param name="cell">Excel单元格，可通过row.GetCell()获取单元格</param>
        /// <returns></returns>
        private string GetCellValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.BLANK:
                    return string.Empty;
                case CellType.BOOLEAN:
                    return cell.BooleanCellValue.ToString();
                case CellType.ERROR:
                    return cell.ErrorCellValue.ToString(CultureInfo.InvariantCulture);
                //case CellType.NUMERIC:
                //case CellType.Unknown:
                default:
                    return cell.ToString();//This is a trick to get the correct value of the cell. NumericCellValue will return a numeric value no matter the cell value is a date or a number
                case CellType.STRING:
                    return cell.StringCellValue;
                case CellType.FORMULA:
                    try
                    {
                        var e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                    }
            }
        }

        public void SetCellValue(object value, int col, int row)
        {
            if (ActiveSheet != null)
            {
                //添加行列判断
                var newr = ActiveSheet.GetRow(row) ?? ActiveSheet.CreateRow(row);
                if (newr.GetCell(col) == null)
                    newr.CreateCell(col);
                if (value == null || value == DBNull.Value)
                    return;
                try
                {
                    switch (value.GetType().ToString())
                    {
                        case "System.String":
                            ActiveSheet.GetRow(row).GetCell(col).SetCellType(CellType.STRING);
                            ActiveSheet.GetRow(row).GetCell(col).SetCellValue(value.ToString());
                            break;
                        case "System.Double":
                            ActiveSheet.GetRow(row).GetCell(col).SetCellType(CellType.NUMERIC);
                            ActiveSheet.GetRow(row).GetCell(col).SetCellValue(Convert.ToDouble(value));
                            break;
                        case "System.Int32":
                            ActiveSheet.GetRow(row).GetCell(col).SetCellType(CellType.NUMERIC);
                            ActiveSheet.GetRow(row).GetCell(col).SetCellValue((int)value);
                            break;
                        case "System.DateTime":
                            ActiveSheet.GetRow(row).GetCell(col).SetCellType(CellType.STRING);
                            ActiveSheet.GetRow(row).GetCell(col).SetCellValue((DateTime)value);
                            break;
                        default:
                            ActiveSheet.GetRow(row).GetCell(col).SetCellType(CellType.NUMERIC);
                            ActiveSheet.GetRow(row).GetCell(col).SetCellValue(Convert.ToDouble(value));
                            break;
                    }
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.Message);
                }
            }
        }

        #region HSSFFormulaEvaluator

        #endregion

    }
}


