// 引入系统基础命名空间
using System;
// 引入集合类（Dictionary、List等）
using System.Collections.Generic;
// 调试输出 Debug.WriteLine 需要这个
using System.Diagnostics;
// LINQ查询语法
using System.Linq;
// 字符串处理
using System.Text;
// 异步Task相关
using System.Threading.Tasks;
// 第三方自定义控件库，UButton就在这个里面
using Zhaoxi.CustControls;

// 工具类所在命名空间，和主窗体项目对应
namespace Zhaoxi.IntelligentWaterStorageMonitorSystem.Utils
{
    /// <summary>
    /// 项目相关的通用工具类，全部是静态方法，直接调用，不用new实例
    /// </summary>
    public static class Utility
    {
        /// <summary>
        /// 静态字典：存储【按钮Name -> 按钮原始样式副本】
        /// 键string：按钮的Name名称（btnConnect、btnDisconnect）
        /// 值UButton：预先复制好的、正常启用状态的按钮样式备份对象
        /// 作用：EnableBtn恢复按钮原始颜色时从这里读取备份
        /// </summary>
        public static Dictionary<string, UButton> disBtnSets = new Dictionary<string, UButton>();

        /// <summary>
        /// 扩展方法：禁用UButton自定义按钮
        /// this UButton btn 代表给UButton这个类新增这个方法，可以直接 btn.UnableBtn() 调用
        /// </summary>
        /// <param name="btn">要禁用的按钮对象</param>
        public static void UnableBtn(this UButton btn)
        {
            // 设置按钮本身不可点击
            btn.Enabled = false;
            // 背景主色改为灰色
            btn.BgColor = Color.Gray;
            // 渐变第二色设置透明
            btn.BgColor2 = Color.Transparent;
            // 文字前景色改为浅灰色
            btn.ForeColor = Color.LightGray;
            // 鼠标聚焦时文字颜色
            btn.FocusForeColor = Color.LightGray;
            // 鼠标聚焦背景色
            btn.FocusBgColor = Color.Gray;
        }

        /// <summary>
        /// 扩展方法：启用UButton自定义按钮，并且恢复备份的原始样式
        /// </summary>
        /// <param name="btn">要启用恢复样式的按钮</param>
        public static void EnableBtn(this UButton btn)
        {
            // 按钮恢复可点击
            btn.Enabled = true;
            // 根据按钮Name，从静态字典取出之前备份好的原始按钮副本
            UButton btnCopy = disBtnSets[btn.Name];
            // 把备份里保存的全部颜色，赋值回当前按钮
            btn.BgColor = btnCopy.BgColor;
            btn.BgColor2 = btnCopy.BgColor2;
            btn.ForeColor = btnCopy.ForeColor;
            btn.FocusForeColor = btnCopy.FocusForeColor;
            btn.FocusBgColor = btnCopy.FocusBgColor;
        }

        /// <summary>
        /// 将2个ushort（Modbus读取到的2个寄存器）拼接转换成C# float单精度浮点数
        /// 注意：这里的字节顺序是【高寄存器在前、低寄存器在后】的调换拼接
        /// Modbus浮点数一般占用连续2个保持寄存器
        /// </summary>
        /// <param name="ushorts">长度必须等于2，两个寄存器ushort数组</param>
        /// <returns>转换完成的float值；数组异常时返回0</returns>
        public static float GetUshortToFloat(ushort[] ushorts)
        {
            // 健壮性校验：数组为空 或者 长度不足2，直接报错日志返回0，防止崩溃
            if (ushorts == null || ushorts.Length < 2)
            {
                Debug.WriteLine($"❌ GetUshortToFloat：传入数组长度不足2，当前长度={ushorts?.Length}");
                return 0f; // 异常返回0，不抛出异常中断程序
            }
            // 准备字节列表，用来组装4个字节（float固定4字节）
            List<byte> bytes = new List<byte>();
            // 把第2个ushort(ushorts[1])转成2个byte，先加入集合
            byte[] bytes1 = BitConverter.GetBytes(ushorts[1]);
            bytes.AddRange(bytes1);
            // 再把第1个ushort(ushorts[0])转成2个byte，追加到后面
            byte[] bytes2 = BitConverter.GetBytes(ushorts[0]);
            bytes.AddRange(bytes2);
            // 把组装好的4字节数组，转成float单精度浮点数，从索引0开始读取
            return BitConverter.ToSingle(bytes.ToArray(), 0);
        }

        /// <summary>
        /// 将2个ushort寄存器拼接成一个uint无符号32位整数
        /// 同样是调换两个寄存器顺序拼接字节，和上面float的字节排布逻辑一致
        /// </summary>
        /// <param name="ushorts">长度为2的ushort寄存器数组</param>
        /// <returns>32位无符号整数uint</returns>
        public static uint GetUshortToUint(ushort[] ushorts)
        {
            List<byte> bytes = new List<byte>();
            // 先取第二个寄存器的字节
            byte[] bytes1 = BitConverter.GetBytes(ushorts[1]);
            bytes.AddRange(bytes1);
            // 再取第一个寄存器的字节
            byte[] bytes2 = BitConverter.GetBytes(ushorts[0]);
            bytes.AddRange(bytes2);
            // 4字节转uint
            return BitConverter.ToUInt32(bytes.ToArray(), 0);
        }
    }
}
