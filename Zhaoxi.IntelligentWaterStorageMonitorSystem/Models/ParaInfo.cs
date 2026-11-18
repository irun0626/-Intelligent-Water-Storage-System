using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zhaoxi.IntelligentWaterStorageMonitorSystem.Models
{
    /// <summary>
    /// 段落信息方法，用来存储获取的表格字段
    /// </summary>
    public class ParaInfo
    {
        public string ParaName { get; set; }

        public ushort Address { get; set; } 

        public ushort FunctionCode { get; set; }

        public string DataType { get; set; }

        public int DCount { get; set; }
    }
}
