using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.Data;

namespace DataAnalysis.Data
{
    static class Helper
    {
        //public List<List<double>> cell;

        /// <summary>
        /// 大于1400小于301的数据掩盖
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DataTable changeDT(DataTable dt)
        {
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 1; c < dt.Columns.Count; c++)
                {
                    double cell = Convert.ToDouble(dt.Rows[r][c].ToString());
                    if (cell < 301 || cell > 1400)
                        //null识别不了
                        //dt.Rows[r][c] = null;
                        dt.Rows[r][c] = "";
                }
            }
            return dt;
        }
    }
}
