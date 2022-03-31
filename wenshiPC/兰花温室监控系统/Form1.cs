using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PortControlDemo;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.IO.Ports;
using System.Data.SqlClient;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace 兰花温室监控系统
{
    public partial class Form1 : Form
    {
        //#（土壤温度2）（土壤湿度2）（环境温度2）（环境湿度2）（光强4）（霍尔电流2）（外遮阳1）（内遮阳1）（顶开窗1）（侧开窗1）（湿帘1）（轴流风机1）（环流风机1）（雾化1）（灌溉1）
        //串口通信部分
        int[] BaudRateArr = new int[] { 115200,4800,2400,1200,300,100 };
        int[] DataBitArr = new int[] { 8,7,6 };
        int[] StopBitArr = new int[] { 1, 2, 3 };
        int[] TimeoutArr = new int[] { 500, 1000, 2000, 5000, 10000 };
        object[] CheckBitArr = new object[] { "None" };
        private bool ReceiveState = false;
        private PortControlHelper pchSend;
        private PortControlHelper pchReceive;
        //TCP通信部分
        Thread getmess = null;//负责定时抓取最新数据；
        Thread threadWatch = null; // 负责监听客户端连接请求的 线程；
        Socket socketWatch = null;
        Dictionary<string, Socket> dict = new Dictionary<string, Socket>();
        Dictionary<string, Thread> dictThread = new Dictionary<string, Thread>();
        private Queue<double> dataQueue = new Queue<double>(100);
        private int num = 1;//每次删除增加几个点
        //private int Wendu = 0;
        public string portstring = "#00000000000000000000000";
        public string mystring1 = "";
        //阈值设置部分&指令下发
        public string waizheyang_order = "0";
        public string neizheyang_order = "0";
        public string dingkaichuang_order = "0";
        public string cekaichuang_order = "0";
        public string shilian_order = "0";
        public string zhouliufengji_order = "0";
        public string huanliufengji_order = "0";
        public string wuhua_order = "0";
        public string guangai_order = "0";
        string dingkaichuang_chongfupanduan = "0";
        //终端控制
        public string yuzhiwenduxia = "00";
        public string yuzhiwendushang = "00";
        public string waizheyang_state = "1";
        public string neizheyang_state = "1";
        public string dingkaichuang_state = "1";
        public string cekaichuang_state = "1";
        public string shilian_state = "1";
        public string zhouliufengji_state = "1";
        public string huanliufengji_state = "1";
        public string wuhua_state = "1";
        public string guangai_state = "1";
        //数据库部分
        private SqlConnection sqlConnection;
        string wenduxia_receive = "00", wendushang_receive = "00", waizheyangshang_receive = "0", waizheyangxia_receive = "0";
        string dingkaichuang_receive = "0", cekaichuang_receive = "0", shilianfengji_receive = "0", zhouliufengji_receive = "0";
        string huanliufengji_receive = "0", wuhuaxitong_receive = "0", guangaixitong_receive = "0";
        string dingkaichuang_control = "1011";
        //节点位
        string Tu_wendu = "00";
        string Tu_shidu = "00";
        string Huan_wendu = "00";
        string Huan_shidu = "00";
        string Guangqiang = "0000";
        string guangqiang_send = "0000";
        string Dianliu = "00";
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            this.skinEngine1 = new Sunisoft.IrisSkin.SkinEngine(((System.ComponentModel.Component)(this)));
            this.skinEngine1.SkinFile = Application.StartupPath + "//office2007.ssk";
            pchSend = new PortControlHelper();
            pchReceive = new PortControlHelper();
            String hostName = Dns.GetHostName();
            IPHostEntry iPHostEntry = Dns.GetHostEntry(hostName);
            foreach (IPAddress ipAddress in iPHostEntry.AddressList)
            {
                if (ipAddress.AddressFamily == AddressFamily.InterNetwork)
                {
                    cbLocalIP.Items.Add(ipAddress.ToString());
                }
            }
            cbLocalIP.SelectedIndex = 0;
            InitView();
            //初始化图表
            InitChart(charttuwen, "土壤温度", Color.Blue, 0, 100);
            InitChart(charttushi, "土壤湿度", Color.Red, 0, 100);
            InitChart(charthuanwen, "环境温度", Color.Green, 0, 100);
            InitChart(charthuanshi, "环境湿度", Color.Tomato, 0, 100);
            InitChart(chartguangqiang, "光照强度", Color.Pink, 0, 10000);
            InitChart(chartdianliu, "电流", Color.Yellow, 0, 100);
            //初始化控制部分
            ID.Text = "room1";
            this.waizheyangshang_bt.SelectedIndex = 0;
            this.waizheyangxia_bt.SelectedIndex = 0;
            this.dingkaichuang_bt.SelectedIndex = 0;
            this.cekaichuang_bt.SelectedIndex = 0;
            this.shilianfengji_bt.SelectedIndex = 0;
            this.zhouliufengji_bt.SelectedIndex = 0;
            this.huanliufengji_bt.SelectedIndex = 0;
            this.wuhua_bt.SelectedIndex = 0;
            this.guangai_bt.SelectedIndex = 0;

        }
        private void InitView()
        {

            cb_portNameReceive.DataSource = pchReceive.PortNameArr;
            cb_baudRate.DataSource = BaudRateArr;
            cb_dataBit.DataSource = DataBitArr;
            cb_stopBit.DataSource = StopBitArr;
            cb_checkBit.DataSource = CheckBitArr;
            cb_timeout.DataSource = TimeoutArr;
            FreshBtnState(pchReceive.PortState);
        }
        private void FreshBtnState(bool state)
        {
            if (state)
            {
                Btn_open.Text = "关闭接收串口";

                Btn_receive.Enabled = true;
            }
            else
            {
                Btn_open.Text = "打开接收串口";

                Btn_receive.Enabled = false;
            }
        }
        private void ComReceiveData(string data)
        {
            this.Invoke(new EventHandler(delegate
            {
                tb_receive.AppendText(data);
                portstring = data;
            }));
        }



       

        private void Btn_open_Click(object sender, EventArgs e)
        {

            if (pchReceive.PortState)
            {              
                pchReceive.ClosePort();
            }
            else
            {
                pchReceive.OpenPort(cb_portNameReceive.Text, int.Parse(cb_baudRate.Text),
               int.Parse(cb_dataBit.Text), int.Parse(cb_stopBit.Text),
               int.Parse(cb_timeout.Text));
            }

            FreshBtnState(pchReceive.PortState);
            pchReceive.OnComReceiveDataHandler += new PortControlHelper.ComReceiveDataHandler(ComReceiveData);
            Btn_receive.Text = "停止接收";
            ReceiveState = true;
        }

        private void Btn_receive_Click(object sender, EventArgs e)
        {
            if (ReceiveState)
            {
                pchReceive.OnComReceiveDataHandler -= new PortControlHelper.ComReceiveDataHandler(ComReceiveData);
                Btn_receive.Text = "开始接收";
                ReceiveState = false;
            }
            else
            {

                pchReceive.OnComReceiveDataHandler += new PortControlHelper.ComReceiveDataHandler(ComReceiveData);
                Btn_receive.Text = "停止接收";
                ReceiveState = true;
            }
        }

     
        void WatchConnecting()
        {
            string shijian = System.DateTime.Now.ToString();//系统时间
            while (true)  // 持续不断的监听客户端的连接请求；
            {
                // 开始监听客户端连接请求，Accept方法会阻断当前的线程；
                Socket sokConnection = socketWatch.Accept(); // 一旦监听到一个客户端的请求，就返回一个与该客户端通信的 套接字；
                // 想列表控件中添加客户端的IP信息；
                lbOnline.Items.Add(sokConnection.RemoteEndPoint.ToString());
                // 将与客户端连接的 套接字 对象添加到集合中；
                dict.Add(sokConnection.RemoteEndPoint.ToString(), sokConnection);
                ShowMsg(shijian + "    客户端连接成功！");
                Thread thr = new Thread(RecMsg);
                thr.IsBackground = true;
                thr.Start(sokConnection);
                dictThread.Add(sokConnection.RemoteEndPoint.ToString(), thr);  //  将新建的线程 添加 到线程的集合中去。
            }
        }
        void RecMsg(object sokConnectionparn)
        {
            Socket sokClient = sokConnectionparn as Socket;
            while (true)
            {
                // 定义一个2M的缓存区；
                byte[] arrMsgRec = new byte[1024 * 1024 * 2];
                // 将接受到的数据存入到输入  arrMsgRec中；
                int length = -1;
                try
                {
                    length = sokClient.Receive(arrMsgRec); // 接收数据，并返回数据的长度；
                }
                catch (SocketException se)
                {
                    ShowMsg("异常：" + se.Message);
                    // 从 通信套接字 集合中删除被中断连接的通信套接字；
                    dict.Remove(sokClient.RemoteEndPoint.ToString());
                    // 从通信线程集合中删除被中断连接的通信线程对象；
                    dictThread.Remove(sokClient.RemoteEndPoint.ToString());
                    // 从列表中移除被中断的连接IP
                    lbOnline.Items.Remove(sokClient.RemoteEndPoint.ToString());
                    break;
                }
                catch (Exception e)
                {
                    ShowMsg("异常：" + e.Message);
                    // 从 通信套接字 集合中删除被中断连接的通信套接字；
                    dict.Remove(sokClient.RemoteEndPoint.ToString());
                    // 从通信线程集合中删除被中断连接的通信线程对象；
                    dictThread.Remove(sokClient.RemoteEndPoint.ToString());
                    // 从列表中移除被中断的连接IP
                    lbOnline.Items.Remove(sokClient.RemoteEndPoint.ToString());
                    break;
                }
                string strMsg = System.Text.Encoding.UTF8.GetString(arrMsgRec, 0, length);// 将接受到的字节数据转化成字符串；
                strMsg = strMsg + "\n";
                Doupdate(strMsg);
            }
        }

        void ShowMsg(string str)//显示函数
        {
            richTextBox2.AppendText(str + "\r\n");
        }
        private void Doupdate(string arrMsg)//希望被执行的函数（被委托）                          接收的数据并解析
        {

            //信号识别解析 
            
           //richTextBox2.AppendText(Convert.ToString(arrMsg.Length));


            if (arrMsg.Length==13) {
                richTextBox2.AppendText("手机端下达指令：\n");
                /*
                wenduxia_receive = arrMsg.Substring(1, 2);
                wendushang_receive = arrMsg.Substring(3, 2);
                waizheyangshang_receive = arrMsg.Substring(5, 1);
                waizheyangxia_receive = arrMsg.Substring(6, 1);
                dingkaichuang_receive = arrMsg.Substring(7, 1);
                cekaichuang_receive = arrMsg.Substring(8, 1);
                shilianfengji_receive = arrMsg.Substring(9, 1);
                zhouliufengji_receive = arrMsg.Substring(10, 1);
                huanliufengji_receive = arrMsg.Substring(11, 1);
                wuhuaxitong_receive = arrMsg.Substring(12, 1);
                guangaixitong_receive = arrMsg.Substring(13, 1);
                */
                wenduxia_receive = arrMsg.Substring(1, 2);
                wendushang_receive = arrMsg.Substring(3, 2);
                guangaixitong_receive = arrMsg.Substring(5, 1);
                dingkaichuang_receive = arrMsg.Substring(6, 1);
                shilianfengji_receive = arrMsg.Substring(7, 1);               
                wuhuaxitong_receive = arrMsg.Substring(8, 1);
                waizheyangshang_receive = arrMsg.Substring(9, 1);
                //从手机端接收数据后发给串口
                string order = "";
                /*
                order = "#" + wenduxia_receive + wendushang_receive + waizheyangshang_receive + waizheyangxia_receive
                    + dingkaichuang_receive + cekaichuang_receive + shilianfengji_receive + zhouliufengji_receive
                    + huanliufengji_receive + wuhuaxitong_receive + guangaixitong_receive + "1#";
                    dingkaichuang_control
                */
                // order = "N4" + shilianfengji_receive + "N5" + wuhuaxitong_receive + guangaixitong_receive + "N6" + dingkaichuang_receive + "\r\n";
                //order = "N1" + wenduxia_receive + wendushang_receive + "N4" + shilianfengji_receive + "N5" + wuhuaxitong_receive + "N6" + dingkaichuang_receive + "N7" + waizheyangshang_receive+"\r\n";




                //waizheyang_state +neizheyang_state + dingkaichuang_state + cekaichuang_state + shilian_state + zhouliufengji_state+ huanliufengji_state + wuhua_state + guangai_state
                Yuzhi_xia_wendu.Text = wenduxia_receive;//温度下限
            Yuzhi_shang_wendu.Text = wendushang_receive;//温度上限                                                      //这边控制位要改
            //外开窗
                if (waizheyangshang_receive == "1")
                 {
                     waizheyangshang_bt.Text = "关";
                     state_waizheyangshang.Text = "关";
                     state_waizheyangxia.Text = "关";
                     waizheyang_state = "1";
                     neizheyang_state = "1";
                 }
                 else if (waizheyangshang_receive == "0")
                {
                     waizheyangshang_bt.Text = "开";
                     state_waizheyangshang.Text = "开";
                     state_waizheyangxia.Text = "开";
                     waizheyang_state = "0";
                     neizheyang_state = "0";
                }
            //内开窗
            if (waizheyangxia_receive == "1")
            {
                waizheyangxia_bt.Text = "关";
            }
            else if (waizheyangxia_receive == "0")
            {
                waizheyangxia_bt.Text = "开";
            }
            //顶开窗
             if (dingkaichuang_receive == "0")
              {
                 //   dingkaichuang_control = "1011";
                 //   if (dingkaichuang_control == dingkaichuang_chongfupanduan)
                 //       dingkaichuang_control = "1111";
                 //   dingkaichuang_chongfupanduan = "1011";
                    dingkaichuang_bt.Text = "关";
                    dingkaichuang_control = "1011";
                     state_dingkaichuang.Text = "关";
                     state_cekaichuag.Text = "关";
                    dingkaichuang_state = "1";
                    cekaichuang_state = "1";
                }
                else if (dingkaichuang_receive == "1")
                {
                     dingkaichuang_bt.Text = "开";
                     dingkaichuang_control = "1110";
                   // dingkaichuang_chongfupanduan = "1110";
                    state_dingkaichuang.Text = "开";
                     state_cekaichuag.Text = "开";
                    dingkaichuang_state = "1";
                    cekaichuang_state = "1";
                }
            //侧开窗
            if (cekaichuang_receive == "1")
            {
                cekaichuang_bt.Text = "关";
            }
            else if (cekaichuang_receive == "0")
            {
                cekaichuang_bt.Text = "开";
            }
            //湿帘风机
            if (shilianfengji_receive == "1")
            {
                shilianfengji_bt.Text = "关";
                state_shilianfengji.Text = "关";
                state_zholiufengji.Text = "关";
                state_huanliufengji.Text = "关";
                    shilian_state = "1";
                    zhouliufengji_state = "1";
                    huanliufengji_state = "1";
                }
            else if (shilianfengji_receive == "0")
            {
                shilianfengji_bt.Text = "开";
                state_shilianfengji.Text = "开";
                state_zholiufengji.Text = "开";
                state_huanliufengji.Text = "开";
                    shilian_state = "0";
                    zhouliufengji_state = "0";
                    huanliufengji_state = "0";
                }
            //轴流风机
            if (zhouliufengji_receive == "1")
            {
                zhouliufengji_bt.Text = "关";
            }
            else if (zhouliufengji_receive == "0")
            {
                zhouliufengji_bt.Text = "开";
            }
            //环流风机
            if (huanliufengji_receive == "1")
            {
                huanliufengji_bt.Text = "关";
            }
            else if (huanliufengji_receive == "0")
            {
                huanliufengji_bt.Text = "开";
            }
            //雾化系统
            if (wuhuaxitong_receive == "1")
                {
                wuhua_bt.Text = "关";
                state_wuhua.Text = "关";
                    wuhua_state = "1";
                }
            else if (wuhuaxitong_receive == "0")
            {
                wuhua_bt.Text = "开";
                state_wuhua.Text = "开";
                    wuhua_state = "0";
                }
            //灌溉系统
              if (guangaixitong_receive == "1")
                 {
                    guangai_bt.Text = "关";
                    state_guangai.Text = "关";
                    guangai_state = "1";
                }
                 else if (guangaixitong_receive == "0")
                {
                     guangai_bt.Text = "开";
                     state_guangai.Text = "开";
                    guangai_state = "0";
                }
                order = "N4" + shilianfengji_receive + "N5" + wuhuaxitong_receive + guangaixitong_receive + "N6" + dingkaichuang_control +"N7"+ waizheyangshang_receive+ "\r\n";
                pchReceive.SendData(order);
            }
        }
        private void UpdateQueueValue()
        {
            Thread.Sleep(100);
            if (dataQueue.Count > 100)
            {
                //先出列
                for (int i = 0; i < num; i++)
                {
                    dataQueue.Dequeue();
                }
            }
           
        }
        void send()//发送函数
        {
            string strMsg = mystring1.ToString();
            byte[] arrMsg = System.Text.Encoding.UTF8.GetBytes(strMsg); // 将要发送的字符串转换成Utf-8字节数组；
            foreach (Socket s in dict.Values)
            {
                s.Send(arrMsg);
            }
            //ShowMsg(strMsg);

        }
        //定时抓取串口数据的函数
        private void GetMess()
        {
            while (true)
            {
                //#（土壤温度2）（土壤湿度2）（环境温度2）（环境湿度2）（光强4）（霍尔电流2）（外遮阳1）（内遮阳1）（顶开窗1）（侧开窗1）（湿帘1）（轴流风机1）（环流风机1）（雾化1）（灌溉1）
                string str = portstring;    //我们抓取当前字符当中的123
                string result = System.Text.RegularExpressions.Regex.Replace(str, @"[^0-9]+", "");
                //richTextBox2.AppendText(result);
                //上面取数字位
                string result3 = result;
                //string result3 = portstring;                            
                double[][] Five_data = new double[1][];//我们要二维数组的形式
                Five_data[0] = new double[7];
                string[] TimeStamp = new string[1];//时间戳数组              
                //mystring1 = result3 + "\n";                      //手机端的消息先注释掉
                //mystring1 = result3 + "\n";

                if (result3 != "")
                {
                    string jiedianwei = result3.Substring(0, 1);
                    int jiedianwei_int = int.Parse(jiedianwei);                   //用来判断哪一个节点发送的数据
                    if (jiedianwei_int == 1&& result3.Length==5)
                    {
                        Huan_wendu = result3.Substring(1, 2);
                        Huan_shidu = result3.Substring(3, 2);
                    }
                    else if (jiedianwei_int == 2 && result3.Length == 5)
                    {
                        Tu_wendu = result3.Substring(1, 2);
                        Tu_shidu = result3.Substring(3, 2);
                    }
                    else if (jiedianwei_int == 3)
                    {
                        int gqchangdu = 0;
                        gqchangdu = result3.Length;
                        if (gqchangdu == 5)
                        {
                            Guangqiang = result3.Substring(1, 4);
                            guangqiang_send = Guangqiang;
                        }
                        else if (gqchangdu == 4)
                        {
                            Guangqiang = result3.Substring(1, 3);
                            guangqiang_send = "0" + Guangqiang;
                        }

                    }
                    else if (jiedianwei_int == 4)
                    {
                        shilian_state = result3.Substring(1, 1);
                        zhouliufengji_state = result3.Substring(1, 1);
                        huanliufengji_state = result3.Substring(1, 1);
                    }
                    else if (jiedianwei_int == 5)
                    {
                        wuhua_state = result3.Substring(1, 1);
                        guangai_state = result3.Substring(1, 1);
                    }
                    else if (jiedianwei_int == 6)
                    {
                        dingkaichuang_state = result3.Substring(1, 1);
                        cekaichuang_state = result3.Substring(1, 1);
                    }
                    else if (jiedianwei_int == 7)
                    {
                        waizheyang_state = result3.Substring(1, 1);
                        neizheyang_state = result3.Substring(1, 1);
                    }

                    mystring1 = "#" + Tu_wendu + Tu_shidu + Huan_wendu + Huan_shidu + guangqiang_send + Dianliu + waizheyang_state
                        + neizheyang_state + dingkaichuang_state + cekaichuang_state + shilian_state + zhouliufengji_state
                        + huanliufengji_state + wuhua_state + guangai_state + "\n";


                }

               

                //外开窗
                if (waizheyang_state == "1")
                {
                    state_waizheyangshang.Text = "关";
                }
                else if (waizheyang_state == "0")
                {
                    state_waizheyangshang.Text = "开";
                }
                //内开窗
                if (neizheyang_state == "1")
                {
                    state_waizheyangxia.Text = "关";
                }
                else if (neizheyang_state == "0")
                {
                    state_waizheyangxia.Text = "开";
                }
                //顶开窗
                if (dingkaichuang_state == "1")
                {
                    state_dingkaichuang.Text = "关";
                }
                else if (dingkaichuang_state == "0")
                {
                    state_dingkaichuang.Text = "开";
                }
                //侧开窗
                if (cekaichuang_state == "1")
                {
                    state_cekaichuag.Text = "关";
                }
                else if (cekaichuang_state == "0")
                {
                    state_cekaichuag.Text = "开";
                }
                //湿帘风机
                if (shilian_state == "1")
                {
                    state_shilianfengji.Text = "关";
                }
                else if (shilian_state == "0")
                {
                    state_shilianfengji.Text = "开";
                }
                //轴流风机
                if (zhouliufengji_state == "1")
                {
                    state_zholiufengji.Text = "关";
                }
                else if (zhouliufengji_state == "0")
                {
                    state_zholiufengji.Text = "开";
                }
                //环流风机
                if (huanliufengji_state == "1")
                {
                    state_huanliufengji.Text = "关";
                }
                else if (huanliufengji_state == "0")
                {
                    state_huanliufengji.Text = "开";
                }
                //雾化系统
                if (wuhua_state == "1")
                {
                    state_wuhua.Text = "关";
                }
                else if (wuhua_state == "0")
                {
                    state_wuhua.Text = "开";
                }
                //灌溉系统
                if (guangai_state == "1")
                {
                    state_guangai.Text = "关";
                }
                else if (guangai_state == "0")
                {
                    state_guangai.Text = "开";
                }
                //发送数据
                send();
                //数据显示
                this.turangwendu.Text = Tu_wendu;//土壤温度
                this.turangshidu.Text = Tu_shidu;//土壤湿度
                this.huanjingwendu.Text = Huan_wendu;//环境温度
                this.huanjingshidu.Text = Huan_shidu;//环境湿度
                this.guangqiang.Text = Guangqiang;//光强
                this.huoerdianliu.Text = Dianliu;//霍尔电流                             
                UpdateQueueValue();
                //图表显示
                int inttuwendu = int.Parse(Tu_wendu);
                int inttushidu = int.Parse(Tu_shidu);
                int inthuanwendu = int.Parse(Huan_wendu);
                int inthuanshidu = int.Parse(Huan_shidu);
                int intguangqiang = int.Parse(Guangqiang);
                int intdianliu = int.Parse(Dianliu);
                AddPoint(charttuwen, inttuwendu);
                AddPoint(charttushi, inttushidu);
                AddPoint(charthuanwen, inthuanwendu);
                AddPoint(charthuanshi, inthuanshidu);
                AddPoint(chartguangqiang, intguangqiang);
                AddPoint(chartdianliu, intdianliu);
                //阈值报警处理
                yuzhiwenduxia = Yuzhi_xia_wendu.Text.ToString().Trim();
                yuzhiwendushang = Yuzhi_shang_wendu.Text.ToString().Trim();
                if (!(yuzhiwenduxia == null || yuzhiwenduxia == ""))              
                     if (inthuanwendu < int.Parse(yuzhiwenduxia))
                     {
                          richTextBox2.AppendText("当前温室温度低于阈值\r\n");
                     }
                if (!(yuzhiwendushang == null || yuzhiwendushang == ""))
                    if (inthuanwendu > int.Parse(yuzhiwendushang))
                     {
                         richTextBox2.AppendText("当前温室温度高于阈值\r\n");
                     }
                 
                if (intdianliu==0)
                {
                    //richTextBox2.AppendText("温室停电！\r\n");
                }
                //数据库增加记录
                string shebeihao="0";
                if (ID.Text.ToString().Trim()!=null)
                    shebeihao= ID.Text.ToString().Trim();
                string sqltuwendu = Tu_wendu + "℃";
                string sqltushidu = Tu_shidu + "%";
                string sqlhuanwendu = Huan_wendu + "℃";
                string sqlhuanshidu = Huan_shidu + "%";
                string sqlquangqiang = Guangqiang + "Lux";
                string sqldianliu = Dianliu + "A";
                string shijian = System.DateTime.Now.ToString();//系统时间
                AddRecord(shebeihao, sqltuwendu, sqltushidu, sqlhuanwendu, sqlhuanshidu, sqlquangqiang, sqldianliu, shijian);
                DispDatabase();
                Thread.Sleep(2000);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string shijian = System.DateTime.Now.ToString();//系统时间
            // 创建负责监听的套接字，注意其中的参数；
            socketWatch = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            // 获得文本框中的IP对象；
            IPAddress address = IPAddress.Parse(cbLocalIP.Text.Trim());
            // 创建包含ip和端口号的网络节点对象；
            IPEndPoint endPoint = new IPEndPoint(address, int.Parse(COM.Text.Trim()));
            //连接SQL服务器
            String sqlStr = "server='.\\SQLwenshi';" +
               "database='wenshi';" +
               "uid='sa';" +
               "pwd='123456';";
            sqlConnection = new SqlConnection(sqlStr);
            sqlConnection.Open();
            ClearRecord();
            DispDatabase();


            try
            {
                // 将负责监听的套接字绑定到唯一的ip和端口上；
                socketWatch.Bind(endPoint);
            }
            catch (SocketException se)
            {
                MessageBox.Show("异常：" + se.Message);
                return;
            }
            // 设置监听队列的长度；
            socketWatch.Listen(10);
            // 创建负责监听的线程；
            threadWatch = new Thread(WatchConnecting);
            threadWatch.IsBackground = true;
            threadWatch.Start();
            //这里用于抓取数据
            getmess = new Thread(GetMess);
            getmess.IsBackground = true;
            getmess.Start();
            
            ShowMsg(shijian + "    服务器启动监听成功！");
            button1.Text = "关闭";
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sqlConnection.Close();
            sqlConnection.Dispose();
            this.Close();
        }

        private void 重启ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("要重新启动吗？", "提示", MessageBoxButtons.YesNoCancel,
 MessageBoxIcon.Question) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location);
                System.Environment.Exit(0);
            }
        }

        

        private void button4_Click(object sender, EventArgs e)  //阈值和控制指令下达
        {
            /*
            N1：空气阈值
            N4：风机系统 1位 N41
            N5：灌溉雾化 1位
            N6：开窗 1位
            N7：遮阳 1位
            */
            yuzhiwenduxia = Yuzhi_xia_wendu.Text.ToString().Trim();  //2
            yuzhiwendushang = Yuzhi_shang_wendu.Text.ToString().Trim();  //2
            //外遮阳
            if (waizheyangshang_bt.Text=="关")                                               //控制位
            {
                waizheyang_order = "1";                            
                state_waizheyangshang.Text = "关";            
                state_waizheyangxia.Text = "关";
                waizheyang_state = "1";
                neizheyang_state = "1";
            }
            else if (waizheyangshang_bt.Text == "开")
            {
                waizheyang_order = "0";
                state_waizheyangshang.Text = "开";
                state_waizheyangxia.Text = "开";
                waizheyang_state = "0";
                neizheyang_state = "0";
            }
            //waizheyang_state +neizheyang_state + dingkaichuang_state + cekaichuang_state + shilian_state + zhouliufengji_state+ huanliufengji_state + wuhua_state + guangai_state
            //内遮阳
            /*
            if (waizheyangxia_bt.Text == "关")
            {
                neizheyang_order = "0";
            }
            else if (waizheyangxia_bt.Text == "开")
            {
                neizheyang_order = "1";
            }
            */
            //顶开窗
            if (dingkaichuang_bt.Text == "关")
            {
                dingkaichuang_order = "1011";
              //  if (dingkaichuang_order == dingkaichuang_chongfupanduan)
              //      dingkaichuang_order = "1111";
            //dingkaichuang_chongfupanduan = "1011";             
                state_dingkaichuang.Text = "关";
                state_cekaichuag.Text = "关";
                dingkaichuang_state = "1";
                cekaichuang_state = "1";
            }
            else if (dingkaichuang_bt.Text == "开")
            {
               // dingkaichuang_chongfupanduan = "1110";
                dingkaichuang_order = "1110";
                state_dingkaichuang.Text = "开";
                state_cekaichuag.Text = "开";
                dingkaichuang_state = "0";
                cekaichuang_state = "0";
            }
            //侧开窗
            /*
            if (cekaichuang_bt.Text == "关")
            {
                cekaichuang_order = "0";
            }
            else if (cekaichuang_bt.Text == "开")
            {
                cekaichuang_order = "1";
            }
            */
            //湿帘风机
            if (shilianfengji_bt.Text == "关")
            {
                shilian_order = "1";
                state_shilianfengji.Text = "关";
                state_zholiufengji.Text = "关";
                state_huanliufengji.Text = "关";
                shilian_state = "1";
                zhouliufengji_state = "1";
                huanliufengji_state = "1";
            }
            else if (shilianfengji_bt.Text == "开")
            {
                shilian_order = "0";
                state_shilianfengji.Text = "开";
                state_zholiufengji.Text = "开";
                state_huanliufengji.Text = "开";
                shilian_state = "0";
                zhouliufengji_state = "0";
                huanliufengji_state = "0";
            }
           
            //雾化系统
            if (wuhua_bt.Text == "关")
            {
                wuhua_order = "1";
                state_wuhua.Text = "关";
                wuhua_state = "1";
            }
            else if (wuhua_bt.Text == "开")
            {
                wuhua_order = "0";
                state_wuhua.Text = "开";
                wuhua_state = "0";
            }
            //灌溉系统
            
            if (guangai_bt.Text == "关")
            {
                guangai_order = "1";
                state_guangai.Text = "关";
                guangai_state = "1";
            }
            else if (guangai_bt.Text == "开")
            {
                guangai_order = "0";
                state_guangai.Text = "开";
                guangai_state = "0";
            }
            
            string order = "";
           

            order = "N4" + shilian_order + "N5" + wuhua_order + guangai_order + "N6" + dingkaichuang_order+"N7"+ waizheyang_order + "\r\n";
            pchReceive.SendData(order);
            richTextBox2.AppendText("指令下达成功\r\n");
        }

        //图表部分
        private void InitChart(Chart chart, String title, Color color, int minY, int maxY)
        {
            Series series = chart.Series[0];
            series.ChartType = SeriesChartType.FastLine;
            series.BorderWidth = 2;
            series.Color = color;

            chart.Legends[0].Enabled = false;

            chart.Titles.Clear();
            chart.Titles.Add(title);
            chart.Titles[0].Text = title;

            ChartArea chartArea = chart.ChartAreas[0];
            chartArea.AxisX.Minimum = 0;
            chartArea.AxisY.Minimum = minY;
            chartArea.AxisY.Maximum = maxY;

            chartArea.AxisX.ScrollBar.IsPositionedInside = false;
            chartArea.AxisX.ScrollBar.Enabled = true;
            chartArea.AxisX.ScaleView.Position = 0;
            chartArea.AxisX.LabelStyle.ForeColor = Color.White;
            chartArea.AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None;

        }
        private delegate void DispLineChartDelegate(Chart chart, double value);
        private void AddPoint(Chart chart, double value)
        {

            if (chart.InvokeRequired)
            {
                DispLineChartDelegate dispLineChartDelegate = new DispLineChartDelegate(AddPoint);
                chart.Invoke(dispLineChartDelegate, new Object[] { chart, value });
            }
            else
            {
                Series series = chart.Series[0];
                series.Points.AddXY(series.Points.Count, value);

                ChartArea chartArea = chart.ChartAreas[0];
                chartArea.AxisX.ScaleView.Position = series.Points.Count - 30;
                chartArea.AxisX.ScaleView.Size = 30;

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void ClearRecord()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "truncate table SensorInfo";
            cmd.Connection = sqlConnection;
            cmd.ExecuteNonQuery();
        }
        private delegate void DispDatabaseDelegate();
        private void DispDatabase()
        {
            if (dataGridView1.InvokeRequired)
            {
                DispDatabaseDelegate dispDatabaseDelegate = new DispDatabaseDelegate(DispDatabase);
                dataGridView1.Invoke(dispDatabaseDelegate, null);
            }
            else
            {
                RrefreshDataView();
            }
        }
        private void AddRecord(String device, String trwendu, String trshidu, String hwendu, String hshidu, String guangq, String dianliu, String shijian)
        {
           
            string sql = string.Format("INSERT INTO SensorInfo(device,tuwendu,tushidu,huanwendu,huanshidu,light,dianliu,time)VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')", device, trwendu, trshidu, hwendu, hshidu, guangq, dianliu, shijian);
            SqlCommand cmd = new SqlCommand(sql, sqlConnection);      
            cmd.ExecuteNonQuery();
        }
        private void RrefreshDataView()
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter("select device as 设备,tuwendu as 土壤温度,tushidu as 土壤湿度,huanwendu as 环境温度,huanshidu as 环境湿度,light as 光强,dianliu as 电流,time as 时间 from SensorInfo", sqlConnection);            
                DataSet set = new DataSet();
                adapter.Fill(set, "SensorInfo");
                dataGridView1.DataSource = set.Tables["SensorInfo"];
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void dayinbaobiao_Click(object sender, EventArgs e)
        {
            ExportToExcel d = new ExportToExcel();
            d.OutputAsExcelFile(dataGridView1);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            
            //sqlConnection.Open();
            string id = comboBox1.SelectedItem.ToString().Trim();
            string sql = String.Format("delete from SensorInfo where device='{0}'", id);
            SqlCommand command = new SqlCommand(sql, sqlConnection);
            command.ExecuteNonQuery();
            RrefreshDataView();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox2.Clear();                           //发送信息提示栏
        }

       
    }
}
