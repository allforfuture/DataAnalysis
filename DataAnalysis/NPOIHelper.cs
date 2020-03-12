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
    class NPOIHelper
    {
        /// <summary>
        /// 将Excel导入DataTable
        /// </summary>
        /// <param name="filepath">导入的文件路径（包括文件名）</param>
        /// <param name="sheetname">工作表名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>DataTable</returns>
        public DataTable ExcelToDataTable(string filepath, string sheetname, bool isFirstRowColumn)
        {
            ISheet sheet = null;//工作表
            DataTable data = new DataTable();

            var startrow = 0;
            using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read))
            {
                try
                {
                    dynamic workbook=null;
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
                        }
                    }
                    return data;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception: " + ex.Message, "读取文件", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                finally { fs.Close(); fs.Dispose(); }
            }
        }
    }
}
