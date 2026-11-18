using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zhaoxi.IntelligentWaterStorageMonitorSystem.Models
{
    /// <summary>
    /// 通信套件方法：用来定义存储从站ip、端口属性
    /// </summary>
    public class CommunicationSet
    {
        public string IP { get; set; } //IP属性
        public int Port { get; set; }//端口属性
    }
}
