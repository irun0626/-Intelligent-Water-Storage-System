using Modbus.Device;
using System.Data;
using System.Diagnostics;
using System.Net.Sockets;
using System.Windows.Forms.DataVisualization.Charting;
using Zhaoxi.Components;
using Zhaoxi.CustControls;
using Zhaoxi.IntelligentWaterStorageMonitorSystem.Models;
using Zhaoxi.IntelligentWaterStorageMonitorSystem.Utils;
using Zhaoxi.Utils;
using Timer = System.Windows.Forms.Timer;
using System.Runtime.Versioning;

namespace Zhaoxi.IntelligentWaterStorageMonitorSystem
{
    [SupportedOSPlatform("windows")]
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        #region 窗口拖动（自定义标题栏拖动）
        Point p = new Point();//暂存首次按下鼠标时的坐标位置
        bool isMove = false;//标识是否处于拖动状态
        private void panelTop_MouseDown(object sender, MouseEventArgs e)
        {
            p = e.Location;//记录鼠标按下瞬间，在panelTop内部的坐标
            isMove = true;//标记：开启拖动
        }
        private void panelTop_MouseMove(object sender, MouseEventArgs e)
        {
            if (isMove && e.Button == MouseButtons.Left)//判断：正在拖动 并且 左键保持按下
            {
                // 窗体位置 = 原位置 + 鼠标移动偏移量（当前坐标 - 按下时的起始坐标）
                this.Location += new Size(e.Location.X - p.X, e.Location.Y - p.Y);
            }
        }
        private void panelTop_MouseUp(object sender, MouseEventArgs e)
        {
            isMove = false;//标记：结束拖动
        }
        #endregion

        // 配置文件路径
        string txtPath = "Communication.txt";
        string excelPath = "DataAddressFile.xls";

        CommunicationSet txtSet = null;//通信设置对象
        List<ParaInfo> paraList = new List<ParaInfo>();//参数点位列表，从Excel加载
        Timer timer = new Timer();//UI定时器，用来刷新管道流动动画
        bool isConnected = false;//PLC连接状态标记
        TcpClient tcpClient = null; //Tcp客户端对象
        IModbusMaster master = null; //Modbus主站对象
        private readonly object _modbusLock = new object();//Modbus读写锁，防止多线程并发冲突

        Dictionary<string, decimal> dicDatas = new Dictionary<string, decimal>();//实时数值字典
        Dictionary<string, bool> dicStates = new Dictionary<string, bool>();//开关状态字典
        List<TrendData> trendDatas = new List<TrendData>(); //曲线历史数据缓存列表

        #region 辅助取值方法
        /// <summary>
        /// 安全读取数值字典，key不存在返回0
        /// </summary>
        private decimal GetDec(string key)
        {
            if (dicDatas.TryGetValue(key, out decimal v))
            {
                return v;
            }
            Debug.WriteLine($"⚠️ 数值key不存在：{key}");
            return 0;
        }

        /// <summary>
        /// 安全读取状态字典，key不存在返回false
        /// </summary>
        private bool GetBool(string key)
        {
            if (dicStates.TryGetValue(key, out bool v))
                return v;
            Debug.WriteLine($"⚠️ 状态key不存在：{key}");
            return false;
        }
        #endregion

        #region 窗体加载事件：全部控件初始化、加载配置
        private void FrmMain_Load(object sender, EventArgs e)
        {
            #region 组态区控件初始化
            //组态区的数据与状态初始化
            //遍历panelDatas面板里面所有子控件
            foreach (Control c in panelDatas.Controls)
            {
                //如果控件是圆形开关 UEllipseSwitch
                if (c is UEllipseSwitch)
                {
                    UEllipseSwitch sw = (UEllipseSwitch)c; //强制类型转换
                    sw.Checked = false;                    //开关置为关闭状态
                }
                //如果控件是旋转开关 URotarySwitch
                else if (c is URotarySwitch)
                {
                    URotarySwitch rs = (URotarySwitch)c;   //强制类型转换
                    rs.SwitchState = false;                //旋转开关置为关闭
                }
                //如果控件是参数输入框 ParaTextBox
                else if (c is ParaTextBox)
                {
                    ParaTextBox txt = (ParaTextBox)c;      //强制类型转换
                    txt.Value = 0;                         //参数值清零
                }
                //如果控件是分组面板 UGroupPanel2（里面还嵌套控件）
                else if (c is UGroupPanel2)
                {
                    //遍历这个分组面板内部所有子控件
                    foreach (Control cc in c.Controls)
                    {
                        //内部子控件如果是参数输入框
                        if (cc is ParaTextBox)
                        {
                            ParaTextBox txt = (ParaTextBox)cc;
                            txt.Value = 0; //参数清零
                        }
                    }
                }
            }
            #endregion

            #region 状态检测区控件初始化
            //状态监测区的初始化
            //遍历gpStates分组内所有控件
            foreach (Control c in gpStates.Controls)
            {
                //如果是点动开关 UJogSwitch
                if (c is UJogSwitch)
                {
                    UJogSwitch js = (UJogSwitch)c;
                    js.IsOn = false; //点动开关默认关闭
                }
                //如果是Label标签，并且Tag不为空（Tag存参数标识）
                if (c is Label && c.Tag != null)
                {
                    Label lbl = (Label)c;
                    string tagStr = lbl.Tag.ToString();//取出标签上绑定的参数标识字符串
                    //标签标识包含BreakState：故障状态
                    if (tagStr.Contains("BreakState"))
                    {
                        lbl.Text = "无"; //故障默认：无故障
                    }
                    //标签标识包含PumpState：泵状态
                    else if (tagStr.Contains("PumpState"))
                    {
                        lbl.Text = "已停止"; //泵默认停止
                    }
                    //标签标识包含ValveState：阀门状态
                    else if (tagStr.Contains("ValveState"))
                    {
                        lbl.Text = "已关闭"; //阀门默认关闭
                    }
                    lbl.ForeColor = Color.Gray; //文字灰色，代表初始未刷新状态
                }
            }
            #endregion

            #region 用水区控件初始化
            wbWaters.Value = 0;        //用水累计控件清零
            lblActualWaters.Value = 0; //实际用水显示控件清零
            #endregion

            #region 趋势曲线区域控件初始化
            chart1.Series.Clear();     //清空图表所有曲线序列
            AddDefaultSeries();        // 调用方法：添加默认曲线序列
            #endregion

            #region 环境数据监测区域控件初始化
            //遍历环境监测面板gpMeters内所有控件
            foreach (Control c in gpMeters.Controls)
            {
                //参数输入框清零
                if (c is ParaTextBox)
                {
                    ParaTextBox txt = (ParaTextBox)c;
                    txt.Value = 0;
                }
                //横向仪表 UMeter 清零
                else if (c is UMeter)
                {
                    ((UMeter)c).Value = 0;
                }
                //圆形仪表盘 UCircleMeter 清零
                else if (c is UCircleMeter)
                {
                    ((UCircleMeter)c).Value = 0;
                }
            }
            #endregion

            #region 通信控制区域初始化
            //保存按钮的副本（把按钮样式配置存入全局字典，用于切换启用/禁用样式）
            Utility.disBtnSets.Add(btnConnect.Name, CopyBtnsets(btnConnect));
            Utility.disBtnSets.Add(btnDisconnect.Name, CopyBtnsets(btnDisconnect));
            IblConnection.Text = "等待连接";       //通信状态文字
            IblConnection.ForeColor = Color.Red;   //红色：未连接
            btnConnect.EnableBtn();                //【连接按钮】启用
            btnDisconnect.UnableBtn();             //【断开按钮】禁用
            LoadCommunication();                   //加载通信参数配置（串口/IP、波特率等）
            LoadDataAddressConfig();               //加载寄存器地址配置（Modbus地址表）
            #endregion

            //初始化定时器（轮询PLC、刷新管道动画的定时器）
            initTimer();
        }
        #endregion

        #region 趋势曲线区所需的自定义方法
        /// <summary>
        /// 给图表添加默认的两条曲线（主管压力、主管流量）
        /// </summary>
        private void AddDefaultSeries()
        {
            // 调用添加曲线方法：曲线名称【主管压力】，绑定参数标识MainOutPressure，线条颜色天蓝色
            AddSplineSeries("主管压力", "MainOutPressure", Color.DodgerBlue);
            // 调用添加曲线方法：曲线名称【主管流量】，绑定参数标识MainOutFlow，线条颜色橙色
            AddSplineSeries("主管流量", "MainOutFlow",Color.Orange);
        }

        /// <summary>
        /// 配置样条曲线Series并添加到chart1图表
        /// 创建一条平滑样条曲线，加入到Winform的Chart控件
        /// </summary>
        /// <param name="seriesName">曲线显示名称（图例上看到的名字）</param>
        /// <param name="paraName">参数标识，存到Tag，后面定时读数据时用来匹配这条曲线</param>
        /// <param name="color">曲线颜色</param>
        private void AddSplineSeries(string seriesName, string paraName, Color color)
        {
            // ========== Series 来源说明 ==========
            // Series 来自命名空间：System.Windows.Forms.DataVisualization.Charting
            // 是微软Chart控件内置的类，代表【一条曲线/一组数据序列】
            Series series = new Series(seriesName);

            // 设置图表类型：Spline = 样条曲线（平滑曲线，不是折线）
            series.ChartType = SeriesChartType.Spline;

            // Tag：自定义附加字段，把参数名字存进去，后面刷新数据的时候，用来识别这条曲线对应哪个PLC变量
            series.Tag = paraName;

            // 线条宽度，像素
            series.BorderWidth = 2;

            // 曲线颜色
            series.Color = color;

            // X轴值类型设置为时间，代表横轴是时间轴
            series.XValueType = ChartValueType.DateTime;

            // 把建好的曲线序列，添加到界面chart1控件里面，界面图例就会出现这条曲线
            chart1.Series.Add(series);
        }
        #endregion

        #region 通信控制区域所需的自定义方法
        //创建一个指定按钮的副本--保存基本样式设置（用于切换启用禁用样式）
        private UButton CopyBtnsets(UButton btn)
        {
            UButton btnCopy = new UButton()
            {
                BgColor = btn.BgColor,
                BgColor2 = btn.BgColor2,
                ForeColor = btn.ForeColor,
                FocusBgColor = btn.FocusBgColor,
                FocusForeColor = btn.FocusForeColor
            };
            return btnCopy;
        }

        //加载通信配置文本 Communication.txt
        private void LoadCommunication()
        {
            if (File.Exists(txtPath))
            {
                string[] lines = File.ReadAllLines(txtPath);
                if (lines.Length == 2)
                {
                    string ip = lines[0].Split(':')[1];//取出ip地址
                    int port = lines[1].Split(":")[1].GetInt();//取出端口号
                    textPort.Text = port.ToString();//把值传给控件
                    textIP.Text = ip.ToString();
                    txtSet = new CommunicationSet()
                    {
                        IP = ip,
                        Port = port,
                    };
                }
            }
            else
            {
                textIP.Text = "127.0.0.1";
                textPort.Text = "502";
            }
        }

        //加载Excel点位配置表 DataAddressFile.xls
        private void LoadDataAddressConfig()
        {
            DataTable dt = ExcelHelper.ExcelToDataTable(excelPath, "Sheet1", true);
            if (dt != null && dt.Rows.Count > 0)
            {
                paraList.Clear();
                foreach (DataRow dr in dt.Rows)
                {
                    ParaInfo para = new ParaInfo();
                    para.ParaName = dr["参数名"].ToString();
                    para.Address = dr["地址"].ToString().GetUShort();
                    para.FunctionCode = dr["功能区"].ToString().GetByte();
                    para.DataType = dr["数据类型"].ToString();
                    para.DCount = dr["小数位数"].ToString().GetInt();
                    paraList.Add(para);
                }
            }
            Debug.WriteLine($"✅ 一共加载 {paraList.Count} 个点位配置");
        }

        /// <summary>
        /// 初始化管道动画定时器
        /// 作用：定时刷新所有UPipe管道控件的Active状态（控制管道是否高亮/通水动画）
        /// </summary>
        private void initTimer()
        {
            // 设置定时器触发间隔：100毫秒 = 0.1秒刷新一次管道状态
            timer.Interval = 100;

            // 绑定定时器Tick事件：每到100ms就执行下面这一段匿名方法
            timer.Tick += (obj, ev) =>//定时器的tick事件绑定一个匿名事件，参数列表为obj, ev，方法体为=>后面的方法提=体
            {
                // 遍历panelDatas面板里面所有放着的控件
                foreach (Control c in panelDatas.Controls)
                {
                    // 判断这个控件是不是自定义管道控件UPipe
                    if (c is UPipe)
                    {
                        // 强制转换为UPipe管道对象
                        UPipe pipe = (UPipe)c;
                        // Invalidate()：标记管道控件画面失效，触发重绘，刷新通水动画
                        pipe.Invalidate();
                    }
                }

                // ========== 管道状态逻辑：Active=true 管道通水高亮；false 灰色关闭 ==========
                // 总阀门控制管道：总阀门开关swValve_M的勾选状态直接决定管道是否通水
                uPipe3.Active = swValve_M.Checked;

                // 第一段管道控制
                // 1号泵的运行状态决定uPipe10管道通水
                uPipe10.Active = rsPump_1.SwitchState;
                // 1号分支阀勾选状态决定uPipe5管道通水
                uPipe5.Active = swValve_1.Checked;

                // 第二段管道控制
                // 2号泵运行状态决定uPipe13管道通水
                uPipe13.Active = rsPump_2.SwitchState;
                // 2号分支阀勾选状态决定uPipe6管道通水
                uPipe6.Active = swValve_2.Checked;

                // 第三段管道控制
                // 3号泵运行状态，同时控制 uPipe11 和 uPipe16 两根管道通水
                uPipe11.Active = uPipe16.Active = rsPump_3.SwitchState;
                // 3号分支阀勾选状态，同时控制 uPipe9 和 uPipe7 两根管道通水
                uPipe9.Active = uPipe7.Active = swValve_3.Checked;

                // 主管：只要1、2、3任意一个分支阀门打开，主管uPipe4就通水
                bool mainBranchActive = swValve_1.Checked || swValve_2.Checked || swValve_3.Checked;
                uPipe4.Active = mainBranchActive;

                // 主管1：只要2号、3号分支阀任意一个打开，主管uPipe8就通水
                bool mainBranchActive1 = swValve_2.Checked || swValve_3.Checked;
                uPipe8.Active = mainBranchActive1;

                // 公共泵管道：三台水泵里面任意一台泵运行，uPipe14公共管道通水
                bool PublicWaterPumpPipe = rsPump_1.SwitchState || rsPump_2.SwitchState || rsPump_3.SwitchState;
                uPipe14.Active = PublicWaterPumpPipe;

                // 公共泵管道1：2号、3号泵任意一台运行，uPipe15公共管道通水
                bool PublicWaterPumpPipe1 = rsPump_2.SwitchState || rsPump_3.SwitchState;
                uPipe15.Active = PublicWaterPumpPipe1;
            };
        }
        #endregion

        /// <summary>
        /// 【连接按钮点击事件】
        /// 点击连接：建立Tcp连接，创建ModbusIP主站，启动采集和动画定时器
        /// sender：触发事件的控件（就是btnConnect按钮）
        /// e：事件参数，这里没有用到
        /// </summary>
        private void btnConnect_Click(object sender, EventArgs e)
        {
            // 判断：当前还没有连接成功，才执行连接逻辑，防止重复连接
            if (isConnected == false)
            {
                try
                {
                    // 获取界面文本框里填写的IP地址
                    string ip = textIP.Text;
                    // 获取端口文本，调用扩展方法GetInt()，把文本转成int数字
                    int port = textPort.Text.GetInt();

                    // 创建TcpClient，主动连接指定IP和端口的服务端（PLC/Modbus网关）
                    tcpClient = new TcpClient(ip, port);
                    // 基于已经连上的tcpClient，创建Modbus TCP主站对象（NModbus库）
                    master = ModbusIpMaster.CreateIp(tcpClient);

                    // 判断TCP通道是否成功连上
                    if (tcpClient.Connected)
                    {
                        isConnected = true; // 全局标记：已经连接
                        IblConnection.Text = "已连接，监测中..."; // 修改状态标签文字
                        IblConnection.ForeColor = Color.DeepSkyBlue; // 文字改成天蓝色

                        textIP.Enabled = false;    // IP输入框禁用，不能再修改
                        textPort.Enabled = false;  // 端口输入框禁用，不能再修改

                        btnConnect.UnableBtn();    // 自定义扩展方法：连接按钮置灰、不可点击
                        btnDisconnect.EnableBtn(); // 自定义扩展方法：断开连接按钮启用，可以点击

                        // 开启后台数据采集线程，用来循环读取Modbus寄存器
                        StartDataMonitor();
                        // 启动管道动画定时器（就是前面那段刷新UPipe管道Active状态的timer）
                        timer.Start();
                    }
                }
                catch (Exception ex)
                {
                    // 捕获连接过程所有异常（连不上、IP错误、端口不通、网络不通等）
                    MessageHelper.Error("异常提示", $"异常：{ex.Message}");
                    // 在VS调试输出窗口打印异常信息，方便调试排查
                    Debug.WriteLine($"连接异常：{ex.Message}");
                    return; // 连接失败，直接退出这个点击方法
                }
            }
        }


        /// <summary>
        /// 通信断开按钮点击事件，释放资源
        /// </summary>
        private void btnDisConnect_Click(object sender, EventArgs e)
        {
            // 判断：当前是已连接状态才执行断开逻辑，避免重复断开
            if (isConnected)
            {
                try
                {
                    // 把连接标记置为未连接
                    isConnected = false;
                    // 关闭Tcp客户端套接字，断开网络通道
                    tcpClient.Close();
                    // 释放Modbus主站对象占用的资源
                    master.Dispose();
                    // 把主站对象置空，方便GC回收
                    master = null;
                    // tcpClient对象置空
                    tcpClient = null;

                    // 修改状态标签文字
                    IblConnection.Text = "等待连接";
                    // 文字颜色改成红色，提示未连接
                    IblConnection.ForeColor = Color.Red;

                    // IP输入框恢复可用，可以修改IP
                    textIP.Enabled = true;
                    // 端口输入框恢复可用，可以修改端口
                    textPort.Enabled = true;

                    // 启用【连接】按钮
                    btnConnect.EnableBtn();
                    // 禁用【断开】按钮
                    btnDisconnect.UnableBtn();

                    // 停止管道动画定时器，不再刷新管道通水动画
                    timer.Stop();
                }
                catch (Exception ex)
                {
                    // 捕获断开时的异常，弹窗提示
                    MessageHelper.Error("异常提示", $"异常：{ex.Message}");
                    // 调试窗口打印异常信息
                    Debug.WriteLine($"断开异常：{ex.Message}");
                    return;
                }
            }
        }


        /// <summary>
        /// 启动后台采集循环任务
        /// </summary>
        private void StartDataMonitor()
        {
            Task.Run(async () =>
            {
                while (true)
                {
                    // 断开连接 / 窗体已经销毁，退出循环
                    if (isConnected == false || this.IsDisposed)
                        break;
                    try
                    {
                        await ReadAndLoadDaas();// 执行一次Modbus读取、解析、更新UI
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"采集循环异常：{ex.Message}");
                    }
                    await Task.Delay(1000);// 等待1000ms（1秒），再下一轮
                }
            });
        }

        /// <summary>
        /// 读取Modbus寄存器、线圈、离散输入，解析数据并刷新界面
        /// 异步采集方法：读保持寄存器、线圈、离散输入；解析数值；跨线程刷新UI；追加图表趋势点
        /// </summary>
        private async Task ReadAndLoadDaas()
        {
            try
            {
                ushort[] udatas;       // 存放【保持寄存器（功能码03）】读取结果
                bool[] states;         // 存放【线圈（功能码01）】读取结果
                bool[] breakStates;    // 存放【离散输入（功能码02）】读取结果

                // Modbus读写加锁，防止并发冲突
                // 多线程同时调用master读取会造成报文错乱、异常，锁保证同一时刻只有一处执行读写
                lock (_modbusLock)
                {
                    if (master == null) return; // Modbus主站对象已经释放，直接退出本次采集

                    // ===== 读取38个保持寄存器 起始地址0，一共38个 =====
                    udatas = master.ReadHoldingRegisters(1, 0, 38);
                    // 2.设备与阀门的状态 线圈 起始0，读取7个线圈
                    states = master.ReadCoils(1, 0, 7);
                    // 3.故障状态 离散输入 起始0，读取3个离散输入
                    breakStates = master.ReadInputs(1, 0, 3);
                }

                // 数据解析与存储：处理保持寄存器（功能码3）
                if (udatas.Length > 0)
                {
                    // 从参数配置列表筛选出功能码=3（保持寄存器）的参数点位
                    var paras = paraList.Where(p => p.FunctionCode == 3).ToList();
                    foreach (var para in paras)
                    {
                        ushort addr = para.Address;// 当前参数对应的寄存器起始地址
                        decimal val = 0.0m; // 存放解析后的数值

                        // 根据参数配置的数据类型做不同解析
                        switch (para.DataType)
                        {
                            case "float":
                                // float在Modbus里占用连续2个ushort寄存器，校验边界防止数组越界
                                if (addr + 2 > udatas.Length)
                                {
                                    Debug.WriteLine($"⚠️浮点点位{para.ParaName} 地址{addr}越界，跳过");
                                    continue;
                                }
                                // 从数组addr位置取连续2个寄存器
                                ushort[] datas = udatas.Skip(addr).Take(2).ToArray();
                                if (datas.Length < 2)
                                {
                                    Debug.WriteLine($"⚠️datas长度不足2，跳过点位：{para.ParaName}");
                                    continue;
                                }
                                // 自定义工具方法：2个ushort拼成IEEE754单精度浮点数
                                float fVal = Utility.GetUshortToFloat(datas);
                                val = (decimal)fVal;
                                break;

                            case "uint":
                                // uint同样占用连续2个寄存器
                                if (addr + 2 > udatas.Length)
                                {
                                    Debug.WriteLine($"⚠️uint点位{para.ParaName} 地址{addr}越界，跳过");
                                    continue;
                                }
                                ushort[] datas1 = udatas.Skip(addr).Take(2).ToArray();
                                val = Utility.GetUshortToUint(datas1);
                                break;

                            case "ushort":
                                // ushort只占用单个寄存器
                                if (addr >= udatas.Length)
                                {
                                    Debug.WriteLine($"⚠️ushort点位{para.ParaName} 地址{addr}越界，跳过");
                                    continue;
                                }
                                ushort udata = udatas[addr];
                                val = udata;
                                break;
                        }

                        // 更新全局数据字典 dicDatas：存放所有采集到的参数值
                        if (dicDatas.ContainsKey(para.ParaName))
                            dicDatas[para.ParaName] = val; // 存在就覆盖更新
                        else
                            dicDatas.Add(para.ParaName, val); // 不存在就新增
                    }
                }

                // 解析线圈、离散输入，交给统一处理方法DealStates
                if (states.Length > 0)
                {
                    DealStates(states, 1);// 传入线圈读取结果，功能码标记1（01读线圈）
                }
                if (breakStates.Length > 0)
                {
                    DealStates(breakStates, 2);// 传入离散输入读取结果，功能码标记2（02读离散输入）
                }

                // ===== 界面刷新
                // this.Invoke：后台Task线程切回UI主线程执行UI赋值逻辑，避免跨线程控件异常
                this.Invoke(new Action(() =>
                {
                    // ===== 开关状态刷新：从dicDatas读取布尔状态赋值给开关控件 =====
                    rsPump_1.SwitchState = GetBool(rsPump_1.Text);
                    rsPump_2.SwitchState = GetBool(rsPump_2.Text);
                    rsPump_3.SwitchState = GetBool(rsPump_3.Text);
                    swValve_M.Checked = GetBool(swValve_M.Text);
                    swValve_1.Checked = GetBool(swValve_1.Text);
                    swValve_2.Checked = GetBool(swValve_2.Text);
                    swValve_3.Checked = GetBool(swValve_3.Text);

                    // ===== 分管1 条件显示：水泵+阀门同时开启才显示压力流量，否则清零 =====
                    bool state1 = rsPump_1.SwitchState && swValve_1.Checked;
                    if (state1)
                    {
                        lblPressure_1.Value = GetDec(lblPressure_1.ParaName);
                        lblFlow_1.Value = GetDec(lblFlow_1.ParaName);
                    }
                    else
                    {
                        lblPressure_1.Value = 0;
                        lblFlow_1.Value = 0;
                    }

                    // ===== 分管2 =====
                    bool state2 = rsPump_2.SwitchState && swValve_2.Checked;
                    if (state2)
                    {
                        lblPressure_2.Value = GetDec(lblPressure_2.ParaName);
                        lblFlow_2.Value = GetDec(lblFlow_2.ParaName);
                    }
                    else
                    {
                        lblPressure_2.Value = 0;
                        lblFlow_2.Value = 0;
                    }

                    // ===== 分管3 =====
                    bool state3 = rsPump_3.SwitchState && swValve_3.Checked;
                    if (state3)
                    {
                        lblPressure_3.Value = GetDec(lblPressure_3.ParaName);
                        lblFlow_3.Value = GetDec(lblFlow_3.ParaName);
                    }
                    else
                    {
                        lblPressure_3.Value = 0;
                        lblFlow_3.Value = 0;
                    }

                    // ===== 主管压力、流量：总阀开启 并且 任意一个分支开启 =====
                    decimal mainPressureVal = 0;
                    decimal mainFlowVal = 0;
                    bool mainEnable = swValve_M.Checked && (state1 || state2 || state3);
                    if (mainEnable)
                    {
                        mainPressureVal = GetDec(lblMainPressure.ParaName);
                        mainFlowVal = GetDec(lblMainFlow.ParaName);
                        lblMainPressure.Value = mainPressureVal;
                        lblMainFlow.Value = mainFlowVal;
                    }
                    else
                    {
                        lblMainPressure.Value = 0;
                        lblMainFlow.Value = 0;
                    }

                    // 当前水位（GetDec做安全读取，防止字典key不存在崩溃）
                    lblCurpos.Value = GetDec(lblCurpos.ParaName);

                    // 加载水泵面板详情数据
                    LoadPumpData(gpPumpData1, rsPump_1);
                    LoadPumpData(gpPumpData2, rsPump_2);
                    LoadPumpData(gpPumpData3, rsPump_3);

                    // 状态监测区：遍历容器内所有控件，动态刷新开关、状态标签文字颜色
                    foreach (Control c in gpStates.Controls)
                    {
                        if (c is UJogSwitch)
                        {
                            UJogSwitch js = (UJogSwitch)c;
                            js.IsOn = GetBool(js.Text);
                        }
                        if (c is Label && c.Tag != null)
                        {
                            Label lbl = (Label)c;
                            string paraName = lbl.Tag.ToString();
                            bool state = GetBool(paraName);

                            if (paraName.Contains("PumpState"))
                            {
                                if (!state)
                                {
                                    lbl.Text = "已停止";
                                    lbl.ForeColor = Color.Red;
                                }
                                else
                                {
                                    lbl.Text = "运行中";
                                    lbl.ForeColor = Color.Orange;
                                }
                            }
                            else if (paraName.Contains("ValveState"))
                            {
                                if (!state)
                                {
                                    lbl.Text = "已关闭";
                                    lbl.ForeColor = Color.Gray;
                                }
                                else
                                {
                                    lbl.Text = "已打开";
                                    lbl.ForeColor = Color.Orange;
                                }
                            }
                            else if (paraName.Contains("BreakState"))
                            {
                                if (!state)
                                {
                                    int id = paraName.Split('-')[1].GetInt();
                                    // 水泵在运行才显示正常，水泵停机显示无
                                    if (GetBool("PumpState-" + id))
                                    {
                                        lbl.Text = "正常";
                                        lbl.ForeColor = Color.DodgerBlue;
                                    }
                                    else
                                    {
                                        lbl.Text = "无";
                                        lbl.ForeColor = Color.Gray;
                                    }
                                }
                                else
                                {
                                    lbl.Text = "故障";
                                    lbl.ForeColor = Color.Red;
                                }
                            }
                        }
                    }

                    // 用水情况区：计划用水量、实际用水量
                    lblPlanWaters.Value = GetDec(lblPlanWaters.ParaName);
                    lblActualWaters.Value = GetDec(lblActualWaters.ParaName);
                    // 计算用水比例，判断除数非0，避免除零异常
                    if (lblPlanWaters.Value != 0)
                    {
                        int value = (int)((lblActualWaters.Value / lblPlanWaters.Value) * 100);
                        wbWaters.Value = value;
                    }
                    else
                    {
                        wbWaters.Value = 0;
                    }

                    // 环境数据监测区：遍历仪表容器，刷新各类自定义仪表控件
                    foreach (Control c in gpMeters.Controls)
                    {
                        if (c is ParaTextBox)
                        {
                            ParaTextBox txt = (ParaTextBox)c;
                            string paraName = txt.ParaName;
                            txt.Value = GetDec(paraName);
                        }
                        else if (c is UMeter)
                        {
                            UMeter meter = (UMeter)c;
                            meter.Value = (double)GetDec(meter.Text);
                        }
                        else if (c is UCircleMeter)
                        {
                            UCircleMeter meter = (UCircleMeter)c;
                            meter.Value = (double)GetDec(meter.Text);
                        }
                    }

                    // ===== 曲线数据：无论是否为0都存入，曲线连续不中断（重点设计）
                    TrendData trendData = new TrendData()
                    {
                        CurTime = DateTime.Now,        // 当前时间戳
                        MainOutPressure = (float)mainPressureVal,  // 主管压力
                        MainOutFlow = (float)mainFlowVal           // 主管流量
                    };
                    trendDatas.Add(trendData);
                    if (trendDatas.Count > 100) // 环形缓存，只保留最近100个采集点，防止内存无限上涨
                    {
                        trendDatas.RemoveAt(0);
                    }
                    // 调用方法：把trendDatas里的数据推到chart图表上渲染
                    RefreshChart();
                }));
            }
            catch (IndexOutOfRangeException ex)
            {
                // 单独捕获数组越界（寄存器解析地址配置错误最容易触发）
                Debug.WriteLine($"💥索引越界：{ex.Message}");
            }
            catch (Exception ex)
            {
                // 捕获其余所有未知异常，只打印日志，不向上抛出；外层循环继续下一轮采集
                Debug.WriteLine($"💥其他异常：{ex.Message}");
            }
        }


        /// <summary>
        /// 解析开关状态，存入状态字典dicStates
        /// </summary>
        private void DealStates(bool[] states, byte functionCode)
        {
            var paras = paraList.Where(p => p.FunctionCode == functionCode).ToList();
            foreach (var para in paras)
            {
                ushort addr = para.Address;
                // 地址不能超过数组最大下标
                if (addr >= states.Length)
                {
                    Debug.WriteLine($"⚠️【功能码{functionCode}】点位 {para.ParaName} 地址{addr}超出数组范围，跳过");
                    continue;
                }
                bool state = states[addr];
                if (dicStates.ContainsKey(para.ParaName))
                {
                    dicStates[para.ParaName] = state;
                }
                else
                    dicStates.Add(para.ParaName, state);
            }
        }

        /// <summary>
        /// 刷新曲线：主管压力(蓝)、主管流量(橙) Spline样条曲线
        /// </summary>
        private void RefreshChart()
        {
            // 新增：没有数据直接返回，防止Min/Max报错
            if (trendDatas.Count == 0)
                return;

            // 遍历图表里面所有的曲线序列（Series）
            foreach (var ser in chart1.Series)
            {
                ser.Points.Clear(); // 清空当前这条曲线上面所有旧的坐标点
            }
            // 循环遍历我们缓存好的每一组历史采样数据
            foreach (var item in trendDatas)
            {
                // 对每一条曲线，填充对应数据
                foreach (var ser in chart1.Series)
                {
                    string tag = ser.Tag.ToString(); // 取出这条曲线预先绑定的标记字符串
                    double val = 0;
                    if (tag == "MainOutPressure")
                    {
                        val = item.MainOutPressure;
                    }
                    if (tag == "MainOutFlow")
                    {
                        val = item.MainOutFlow;
                    }
                    // 向曲线添加一个坐标点：X=时间，Y=压力/流量数值
                    ser.Points.AddXY(item.CurTime, val);
                }
            }
            // X轴自动滚动
            chart1.ChartAreas[0].AxisX.Minimum = trendDatas.Min(d => d.CurTime).ToOADate();
            chart1.ChartAreas[0].AxisX.Maximum = trendDatas.Max(d => d.CurTime).ToOADate();
            // X轴时间标签格式：时:分:秒
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";
        }

        /// <summary>
        /// 加载水泵数据文本框的数据方法
        /// </summary>
        private void LoadPumpData(UGroupPanel2 gp, URotarySwitch js)
        {
            foreach (Control c in gp.Controls) //遍历UGroupPanel2里面的控件
            {
                if (c is ParaTextBox)
                {
                    ParaTextBox txt = (ParaTextBox)c;
                    string paraName = txt.ParaName;
                    if (js.SwitchState)
                    {
                        txt.Value = GetDec(paraName);
                    }
                    else
                        txt.Value = 0;
                }
            }
        }

        /// <summary>
        /// UI界面水泵的启停控制下发
        /// </summary>
        private void rePump_StateChange(object sender, EventArgs e)
        {
            if (isConnected)
            {
                //获取控件
                URotarySwitch js = (URotarySwitch)sender;
                string paraName = js.Text;
                bool state = js.SwitchState;
                ParaInfo para = paraList.Find(p => p.ParaName == paraName);
                //修复：判断点位是否找到，防止空引用
                if (para == null)
                {
                    MessageHelper.Error("错误提示", $"未在配置表找到点位：{paraName}");
                    js.SwitchState = !state;
                    return;
                }
                ushort addr = para.Address;
                lock (_modbusLock)
                {
                    master.WriteSingleCoil(1, addr, state);
                }
                dicStates[paraName] = state;
            }
            else
            {
                MessageHelper.Error("错误提示", "当前未连接，不能进行水泵控制！");
                //回滚控件状态，保持和PLC一致
                URotarySwitch js = (URotarySwitch)sender;
                js.SwitchState = !js.SwitchState;
                return;
            }
        }

        /// <summary>
        /// Ui界面的阀门启停控制下发
        /// </summary>
        private void swValve_CheckedChanged(object sender, EventArgs e)
        {
            if (isConnected)
            {
                //获取控件
                UEllipseSwitch sw = (UEllipseSwitch)sender;
                string paraName = sw.Text;
                bool state = sw.Checked;
                ParaInfo para = paraList.Find(p => p.ParaName == paraName);
                if (para == null)
                {
                    MessageHelper.Error("错误提示", $"未在配置表找到点位：{paraName}");
                    sw.Checked = !state;
                    return;
                }
                ushort addr = para.Address;
                lock (_modbusLock)
                {
                    master.WriteSingleCoil(1, addr, state);
                }
                dicStates[paraName] = state;
            }
            else
            {
                MessageHelper.Error("错误提示", "当前未连接，不能进行阀门控制！");
                //回滚控件状态
                UEllipseSwitch sw = (UEllipseSwitch)sender;
                sw.Checked = !sw.Checked;
                return;
            }
        }

        //窗口退出按钮
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //窗体关闭事件，资源释放
        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageHelper.Question("退出提示", "你确定要退出系统么？") == DialogResult.OK)
            {
                if (isConnected)
                {
                    timer.Stop();
                    lock (_modbusLock)
                    {
                        tcpClient?.Close();
                        master?.Dispose();
                    }
                    isConnected = false;
                }
                Application.Exit();
            }
            else
            {
                e.Cancel = true;//取消退出
            }
        }
    }
}
