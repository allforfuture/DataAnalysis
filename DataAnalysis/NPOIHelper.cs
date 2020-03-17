using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.IO;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace DataAnalysis
{
    static class NPOIHelper
    {
        /// <summary>
        /// 将Excel导入DataTable
        /// </summary>
        /// <param name="filepath">导入的文件路径（包括文件名）</param>
        /// <param name="sheetname">工作表名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>DataTable</returns>
        public static DataTable Excel2DT(string filepath, string sheetname, bool isFirstRowColumn,ref dynamic workExcel)
        {
            ISheet sheet;//工作表
            DataTable data = new DataTable();

            int startrow;
            using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read))
            {
                
                   dynamic workbook = null;
                try
                {
                    if (filepath.IndexOf(".xlsx") > 0) // 2007版本
                        workbook = new XSSFWorkbook(fs);
                    else if (filepath.IndexOf(".xls") > 0) // 2003版本
                        workbook = new HSSFWorkbook(fs);
                    if (sheetname != null)
                    {
                        sheet = workbook.GetSheet(sheetname);
                        if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                        {
                            sheet = workbook.GetSheetAt(0);
                        }
                    }
                    else
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                    if (sheet != null)
                    {
                        IRow firstrow = sheet.GetRow(0);
                        int cellCount = firstrow.LastCellNum; //行最后一个cell的编号 即总的列数
                        if (isFirstRowColumn)
                        {
                            for (int i = firstrow.FirstCellNum; i < cellCount; i++)
                            {
                                ICell cell = firstrow.GetCell(i);
                                if (cell != null)
                                {
                                    string cellvalue = cell.StringCellValue;
                                    if (cellvalue != null)
                                    {
                                        DataColumn column = new DataColumn(cellvalue);
                                        data.Columns.Add(column);
                                    }
                                }
                            }
                            startrow = sheet.FirstRowNum + 1;
                        }
                        else
                        {
                            startrow = sheet.FirstRowNum;
                        }
                        //读数据行
                        int rowcount = sheet.LastRowNum;
                        for (int i = startrow; i <= rowcount; i++)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row == null)
                            {
                                continue; //没有数据的行默认是null
                            }
                            DataRow datarow = data.NewRow();//具有相同架构的行
                            for (int j = row.FirstCellNum; j < cellCount; j++)
                            {
                                if (row.GetCell(j) != null)
                                {
                                    string a = row.GetCell(j).ToString();
                                    datarow[j] = row.GetCell(j).ToString();
                                }
                            }
                            data.Rows.Add(datarow);
                            if (i >= Main.main.MaxRows)
                                break;
                        }
                    }
                    return data;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception: " + ex.Message, "ReadFile", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                finally
                {
                    fs.Close();
                    fs.Dispose();
                    workExcel = workbook;
                }
            }
        }

        /// <summary>
        /// 将DataTable导入到Excel
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="filepath">导入的文件路径（包含文件名称）</param>
        /// <param name="sheename">要导入的表名</param>
        /// <param name="iscolumwrite">是否写入列名</param>
        /// <returns>导入Excel的行数</returns>
        public static int DT2Excel(DataTable data, string filepath, string sheename, bool iscolumwrite)
        {
            int i, j, count;
            ISheet sheet;
            using (FileStream fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                dynamic workbook = null;
                //根据Excel不同版本实例不同工作簿
                if (filepath.IndexOf(".xlsx") > 0) // 2007版本
                {
                    workbook = new XSSFWorkbook();
                }
                else if (filepath.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook();

                try
                {
                    if (workbook != null)
                    {
                        sheet = workbook.CreateSheet(sheename);
                    }
                    else
                    {
                        return -1;
                    }

                    if (iscolumwrite == true) //写入DataTable的列名
                    {
                        IRow row = sheet.CreateRow(0);
                        for (j = 0; j < data.Columns.Count; ++j)
                        {
                            row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                        }
                        count = 1;
                    }
                    else
                    {
                        count = 0;
                    }

                    for (i = 0; i < data.Rows.Count; ++i)
                    {
                        IRow row = sheet.CreateRow(count);
                        for (j = 0; j < data.Columns.Count + 3; ++j)
                        {
                            if (j < data.Columns.Count && data.Rows[i][j]!="")
                                row.CreateCell(j).SetCellValue(Convert.ToDouble(data.Rows[i][j]));
                            //写入公式（极差）
                            else if (j == data.Columns.Count)
                            {
                                int Int_A = 65;//A的ascii码
                                string startC = NPOIHelper.Chr(Int_A + 1);//B开始
                                string endC = NPOIHelper.Chr((Int_A + j) - 1);
                                string formula = String.Format("MAX({0}{2}:{1}{2})-MIN({0}{2}:{1}{2})", startC, endC, i + 2);
                                row.CreateCell(j).CellFormula = formula;
                            }
                            //写入公式（平均数）
                            else if (j == data.Columns.Count + 1)
                            {
                                int Int_A = 65;
                                string startC = NPOIHelper.Chr(Int_A + 1);
                                string endC = NPOIHelper.Chr((Int_A + j) - 2);
                                string formula = String.Format("AVERAGE({0}{2}:{1}{2})", startC, endC, i + 2);
                                row.CreateCell(j).CellFormula = formula;
                            }
                            //写入公式（异常显示）
                            else if (j == data.Columns.Count + 2)
                            {
                                int Int_A = 65;
                                //string startC = NPOIHelper.Chr(Int_A + 1);
                                string endC = NPOIHelper.Chr((Int_A + j)-2);
                                //IF(T2>5,"Range>5","")
                                string formula = String.Format("IF({0}{1}>5,\"Range > 5\",\"\")", endC,  i + 2);
                                row.CreateCell(j).CellFormula = formula;
                            }
                        }
                        count++;
                    }
                    workbook.Write(fs); //写入到excel
                    return count;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception: " + ex.Message);
                    return -1;
                }
                finally { fs.Close(); fs.Dispose(); }
            }
        }

        /// <summary>
        /// 将DataTable导入到Excel
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="filepath">导入的文件路径（包含文件名称）</param>
        /// <param name="sheename">要导入的表名</param>
        /// <param name="iscolumwrite">是否写入列名</param>
        /// <returns>导入Excel的行数</returns>
        public static int DT2Excel_Cover(DataTable data, dynamic workbook, string filepath, string sheename, bool iscolumwrite)
        {
            int i, j, count;
            ISheet sheet;
            using (FileStream fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                try
                {
                    if (workbook != null)
                    {
                        sheet = workbook.CreateSheet(sheename);
                    }
                    else
                    {
                        return -1;
                    }

                    if (iscolumwrite == true) //写入DataTable的列名
                    {
                        IRow row = sheet.CreateRow(0);
                        for (j = 0; j < data.Columns.Count; ++j)
                        {
                            row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                        }
                        count = 1;
                    }
                    else
                    {
                        count = 0;
                    }

                    for (i = 0; i < data.Rows.Count; ++i)
                    {
                        IRow row = sheet.CreateRow(count);
                        for (j = 0; j < data.Columns.Count + 3; ++j)
                        {
                            if (j < data.Columns.Count && data.Rows[i][j] != "")
                                row.CreateCell(j).SetCellValue(Convert.ToDouble(data.Rows[i][j]));
                            //写入公式（极差）
                            else if (j == data.Columns.Count)
                            {
                                int Int_A = 65;//A的ascii码
                                string startC = NPOIHelper.Chr(Int_A + 1);//B开始
                                string endC = NPOIHelper.Chr((Int_A + j) - 1);
                                string formula = String.Format("MAX({0}{2}:{1}{2})-MIN({0}{2}:{1}{2})", startC, endC, i + 2);
                                row.CreateCell(j).CellFormula = formula;
                            }
                            //写入公式（平均数）
                            else if (j == data.Columns.Count + 1)
                            {
                                int Int_A = 65;
                                string startC = NPOIHelper.Chr(Int_A + 1);
                                string endC = NPOIHelper.Chr((Int_A + j) - 2);
                                string formula = String.Format("AVERAGE({0}{2}:{1}{2})", startC, endC, i + 2);
                                row.CreateCell(j).CellFormula = formula;
                            }
                            //写入公式（异常显示）
                            else if (j == data.Columns.Count + 2)
                            {
                                int Int_A = 65;
                                //string startC = NPOIHelper.Chr(Int_A + 1);
                                string endC = NPOIHelper.Chr((Int_A + j) - 2);
                                //IF(T2>5,"Range>5","")
                                string formula = String.Format("IF({0}{1}>5,\"Range > 5\",\"\")", endC, i + 2);
                                row.CreateCell(j).CellFormula = formula;
                            }
                        }
                        count++;
                    }
                    workbook.Write(fs); //写入到excel
                    return count;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception: " + ex.Message);
                    return -1;
                }
                finally { fs.Close(); fs.Dispose(); }
            }
        }

        /// <summary>
        /// 将DataTable导出到Execl文档
        /// </summary>
        /// <param name="dt">传入一个DataTable数据集</param>
        /// <returns>返回一个Bool类型的值，表示是否导出成功</returns>
        /// True表示导出成功，Flase表示导出失败
        public static bool DT2Excel_(DataTable dt, string Outpath)
        {
            bool result = false;
            IWorkbook workbook;
            FileStream fs = null;
            IRow row;
            ISheet sheet;
            ICell cell;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new HSSFWorkbook();
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数

                    //设置列头
                    row = sheet.CreateRow(0);//excel第一行设为列头
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    //向outPath输出数据
                    using (fs = File.OpenWrite(Outpath))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据
                        result = true;
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        static string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
            }
        }
    }
}
