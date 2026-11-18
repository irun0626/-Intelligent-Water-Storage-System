using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zhaoxi.IntelligentWaterStorageMonitorSystem.Models
{
    /// <summary>
    /// 曲线数据方法定义
    /// </summary>
    public class TrendData
    {

        public DateTime CurTime { get; set; }
        public float MainOutPressure { get; set; }
        public float MainOutFlow { get; set; }
    }
}
