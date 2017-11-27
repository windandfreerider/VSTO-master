using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Layered.Application.Services;
using Layered.Application.Services.Dto;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Layered.UI
{
    public partial class Sheet1
    {
        private IAverageScoreService _averageScoreService = new AverageScoreService();
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            #region 读(Read)
            
            //从Excel表格读取数据
            var dto = GetData();
            
            #endregion

            #region 处理数据(Action)

            //调用业务逻辑层处理数据，并获取得返回结果
            var result = _averageScoreService.Calculation(GetData()); 

            #endregion

            #region 写(Write)

            //将结果写入Excel表格
            Cells[2, 7].Value = result.First(s => s.Subject == "语文").Score;
            Cells[3, 7].Value = result.First(s => s.Subject == "数学").Score; 

            #endregion
        }

        /// <summary>
        /// 加载Excel数据
        /// </summary>
        private List<AverageScoreDto> GetData()
        {
            //创建数据参数对象
            var data = new List<AverageScoreDto>();
            // 添加语文数据
            var rng = (Range)Range[Cells[2, 2], Cells[5, 2]];
            data.AddRange(rng.Cast<Range>().Select(s => new AverageScoreDto
            {
                Subject = "语文",
                Score = decimal.Parse(s.Value.ToString())
            }));
            // 添加数学数据
            rng = (Range)Range[Cells[2, 3], Cells[5, 3]];
            data.AddRange(rng.Cast<Range>().Select(s => new AverageScoreDto
            {
                Subject = "数学",
                Score = decimal.Parse(s.Value.ToString())
            }));
            return data;
        }
    }
}
