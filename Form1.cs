using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using System.Management;
using System.Text.RegularExpressions;
using ClosedXML.Excel;


namespace serial_com
{
    public partial class Form1 : Form
    {
        private bool button1_flag = false;
        private bool button2_flag = false;
        private bool parent_manager_called = false;
        private CancellationTokenSource _cancellationtoken = new CancellationTokenSource();
        public Dictionary<string, string> my_all_data = new Dictionary<string, string>();    //グラフに書かれているデータ（すべてここに入る）
        private MyCom class_my_com;
        private MyGraph class_my_gra;
        private DateTime class_my_starttime;    //記録開始ボタンを押した直後の時間が入る
        //private bool class_check_box;

        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle; //autolayoutが実装できるまでこのサイズを保つ
            this.Text = "Vacuum Gauge Controller";
            button1.Click += this.Button1_Click;
            button2.Click += this.Button2_Click;
            saveFileDialog1.Filter = "xlsxファイル(*.xlsx)|*.xlsx";

            checkBox1.Checked = true;   //チェックされた状態からスタート
            
            //メニューバーイベントハンドラー
            開くOToolStripMenuItem.Click += toolstrip_open_click;
            名前を付けて保存AToolStripMenuItem.Click += toolstrip_save_as_click;
            終了ToolStripMenuItem.Click += toolstrip_exit_click;
            プロットエリアの設定ToolStripMenuItem.Click += toolstrip_plot_config;
            上書き保存SToolStripMenuItem.Click += toolstrip_save_samename_click;
            上書き保存SToolStripMenuItem.Enabled = false;    //上書きしたくない
            積分ToolStripMenuItem.Click += toolstrip_integral;
            開くOToolStripMenuItem.Enabled = false;
            /*foreach (var value in SerialPort.GetPortNames())
            {      //コンボボックスにcom port 追加
                ToolStripMenuItem tmp = new ToolStripMenuItem();
                tmp.Text = value;
                comポートToolStripMenuItem.DropDownItems.Add(tmp);
                this.my_coms.Add(tmp);
                tmp.Click += toolstrip_com_click;
            }*/
            ManagementClass mg = new ManagementClass("Win32_SerialPort");
            foreach (ManagementObject port in mg.GetInstances()) {
                ToolStripMenuItem tmp = new ToolStripMenuItem();
                tmp.Text = port.GetPropertyValue("Caption").ToString();
                comポートToolStripMenuItem.DropDownItems.Add(tmp);
                this.my_coms.Add(tmp);
                tmp.Click += toolstrip_com_click;
            }

        }

        private void Button2_Click(object sender, EventArgs e) {
            if (this.selected_com == null)
            {
                MessageBox.Show("comポートを選んでください");
                return;
            }
            if (!(button2_flag) && !(this.parent_manager_called)) {
                this.class_my_com = new MyCom(this.selected_com);
                this.button2_flag = true;
                class_my_gra = new MyGraph(this.chart1);
                class_my_starttime = DateTime.Now;  //エラー回避のため
                this.Parent_ManagerAsync(class_my_com, 1, class_my_gra);
            }
        }

        struct MyPoints {
            public double X { get; }
            public double Y { get; }
            public MyPoints(double xx, double yy) {
                this.X = xx;
                this.Y = yy;
            }
        }

        private void Button1_Click(object sender, EventArgs e) {
            if (this.selected_com == null) {
                MessageBox.Show("comポートを選んでください");
                return;
            }
            if (this.button1_flag == false && this.my_all_data.Count > 0)
            {
                DialogResult my_result = MessageBox.Show("今のグラフに描かれているデータは消えますがよろしいですか?", "question", MessageBoxButtons.YesNo);
                if (my_result == DialogResult.Yes)
                {                   //すべて初期化(エラー起きやすい)
                    this.my_all_data.Clear();
                    //this.chart1.Series.Remove(this.class_my_gra.which_main_series());
                    //this.class_my_gra = new MyGraph(this.chart1);
                    this._cancellationtoken.Cancel();
                    Thread.Sleep(300);
                    this._cancellationtoken.Dispose();
                    this._cancellationtoken = new CancellationTokenSource();
                    this.parent_manager_called = false;
                    this.class_my_starttime = DateTime.Now;
                }
                else
                {
                    return;
                }
            }

            this.button1_flag = !this.button1_flag;
            if (this.button1_flag)
            {
 
                this.button1.Text = "STOP";
                if (this.parent_manager_called == false)
                {
                    this.class_my_com = new MyCom(this.selected_com);
                    this.class_my_gra = new MyGraph(this.chart1);
                    this.class_my_starttime = DateTime.Now;
                    this.Parent_ManagerAsync(class_my_com, 1, class_my_gra);
                }
                else {
                    this.class_my_starttime = DateTime.Now; //よく考えてない
                }
                //graph code
                //MyGraph mygra = new MyGraph(this.chart1);
                //List<MyPoints> mydatas = Form1.MyGraph.make_data();
                //mygra.Plot(mydatas);
            }
            else
            {
                this.button1.Text = "記録開始";
                //_cancellationtoken.Cancel();        //button2用にparent_managerAsync が死んでしまうので再起動しておく必要がある(不要かも)
                //_cancellationtoken.Dispose();
                //ここにデータをファイルに保存するかを決める
                if (this.checkBox1.Checked == true) {
                    名前を付けて保存AToolStripMenuItem.PerformClick();
                    return;
                }
                DialogResult __result = MessageBox.Show("ファイルを保存しますか?","question",MessageBoxButtons.YesNo);
                if (__result == DialogResult.Yes) {
                    名前を付けて保存AToolStripMenuItem.PerformClick();
                }
            }


        }
        private async Task Parent_ManagerAsync(MyCom myc, int intervaltime, MyGraph myg) { //データの書き込み、グラフ作成を行う
            this.parent_manager_called = true;
            await Task.Run(() => {
                while (!this._cancellationtoken.Token.IsCancellationRequested) {
                    string pressure = myc.Manager();
                    if (!this._cancellationtoken.Token.IsCancellationRequested)
                    {
                        this.UpdateText(pressure);
                    }

                    if (this.button1_flag) {            //記録開始ボタンが押されているとき
                        DateTime nowtime = DateTime.Now;
                        this.Write_csv_and_graph(pressure, nowtime, myg, 1);
                    }
                    //Thread.Sleep(intervaltime * 1000);
                }

            });

            //ここにデータをファイルに保存するようなコードを書く(cancellationtokenを使う場合のみ)
        }
        private void Write_csv_and_graph(string pressure, DateTime nowtime, MyGraph myg, int intervaltime) { //my_all_dataに代入とグラフを描く関数
            //graph code
            int passed_time = this.time_to_int(this.class_my_starttime, nowtime);
            MyPoints tmp_point = new MyPoints(passed_time, double.Parse(pressure));
            if (!this.my_all_data.ContainsKey(passed_time.ToString()))
            {
                this.my_all_data[passed_time.ToString()] = pressure;
                myg.slowly_plot(tmp_point, this);
            }
            //write_csv
            //MessageBox.Show(passed_time.ToString());

            //using (StreamWriter sw = new StreamWriter(save_name, true, Encoding.GetEncoding("shift_jis"))) {    //例外とらえていない
            //    sw.WriteLine(nowtime.ToString() + "," + pressure);
            //}

        }
        private void UpdateText(string text) {
            if (this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate () { UpdateText(text); });
                return;
            }
            else
            {
                this.label1.Text = text;
            }
        }
        private int time_to_int(DateTime starttime, DateTime nowtime) {  //経過時間、分、秒を秒換算でintにして返す
            //return time_str.Hour * 3600 + time_str.Minute * 60 + time_str.Second;
            TimeSpan tms = nowtime - starttime;
            //Console.WriteLine(tms.TotalSeconds);
            return int.Parse((tms.Ticks / 10000000).ToString());     //intで足りるのか?

        }
        class MyGraph {
            private Series graphdata;
            private Chart chartcontrol;
            public MyGraph(Chart chartc) {      //chartcontrolインスタンスを入れる
                this.graphdata = new Series();
                this.graphdata.ChartType = SeriesChartType.Line;
                this.graphdata.Name = "graph1";
                this.chartcontrol = chartc;
                this.chartcontrol.Series.Clear();
                init();
               
            }
            private void init() {
                this.chartcontrol.Series.Add(this.graphdata);
                this.chartcontrol.Series[0].MarkerStyle = MarkerStyle.Circle;
                this.chartcontrol.Series[0].MarkerSize = 8;
                this.chartcontrol.Series[0].MarkerColor = Color.Red;
                this.chartcontrol.Series[0].IsVisibleInLegend = false;
                this.chartcontrol.Series[0].IsValueShownAsLabel = false;
                
            }
            public void slowly_plot(MyPoints point, object ob) {
                Form1 tmp = (Form1)ob;

                string[] tmpstr = { };
                foreach (var value in this.chartcontrol.ChartAreas) {
                    tmpstr = value.ToString().Split('-');
                    //this.chartcontrol.ChartAreas[tmpstr[1]].AxisX.Maximum += 1;   //グラフエリアのx軸の拡張
                   
                }
               

                if (tmp.InvokeRequired)
                {
                    tmp.Invoke((MethodInvoker)delegate () { this.slowly_plot(point, ob); });
                    return;
                }
                else
                {
                    this.graphdata.Points.AddXY(point.X, point.Y);
                    //if (this.chartcontrol.ChartAreas[tmpstr[1]].AxisX.Maximum <= point.X)
                    //{
                    //  this.chartcontrol.ChartAreas[tmpstr[1]].AxisX.Maximum += 10;
                    //}
                    if (tmp.class_plot_xmax != null && tmp.class_plot_xmin != null)
                    {   //エラーハンドリング必要
                        this.chartcontrol.ChartAreas[tmpstr[1]].AxisX.Maximum = (double)tmp.class_plot_xmax;
                        this.chartcontrol.ChartAreas[tmpstr[1]].AxisX.Minimum = (double)tmp.class_plot_xmin;
                       
                    }
                    if (tmp.class_plot_ymax != null && tmp.class_plot_ymin != null)
                    {
                        this.chartcontrol.ChartAreas[tmpstr[1]].AxisY.Maximum = (double)tmp.class_plot_ymax;
                        this.chartcontrol.ChartAreas[tmpstr[1]].AxisY.Minimum = (double)tmp.class_plot_ymin;
                    }



                }
            }
            public void my_axis_changed(Form1 tmp,double? xmin,double? xmax) {
                string[] tmpstr = { };
                foreach (var value in this.chartcontrol.ChartAreas)
                {
                    tmpstr = value.ToString().Split('-');
                   
                }
                if (tmp.InvokeRequired)
                {
                    tmp.Invoke((MethodInvoker)delegate () { this.my_axis_changed(tmp,xmin,xmax); });
                    return;
                }
                else
                {
                    if (xmin != null) {
                        tmp.chart1.ChartAreas[tmpstr[1]].AxisX.Minimum = (double)xmin;
                        return;
                    }
                    if (xmax != null) {
                        tmp.chart1.ChartAreas[tmpstr[1]].AxisX.Maximum = (double)xmax;
                    }
                    

                    if (tmp.class_plot_xmax != null && tmp.class_plot_xmin != null)
                    {   //エラーハンドリング必要
                        tmp.chart1.ChartAreas[tmpstr[1]].AxisX.Maximum = (double)tmp.class_plot_xmax;
                        tmp.chart1.ChartAreas[tmpstr[1]].AxisX.Minimum = (double)tmp.class_plot_xmin;
                    }
                    if (tmp.class_plot_ymax != null && tmp.class_plot_ymin != null) {
                        this.chartcontrol.ChartAreas[tmpstr[1]].AxisY.Maximum = (double)tmp.class_plot_ymax;
                        this.chartcontrol.ChartAreas[tmpstr[1]].AxisY.Minimum = (double)tmp.class_plot_ymin;
                    }
                }
            }
            public void Plot(My_Read_struct myrs,Form1 myform) {      //Tはdouble or float
                this.my_axis_changed(myform,0,null);
                foreach (var _val in myrs.dc){
                    //this.graphdata.Points.AddXY(_val.Key, _val.Value);
                    this.safety_points_addxy(_val.Key,_val.Value,myform);
                }
                //this.chartcontrol.Series.Add(this.graphdata);

            }
            private void safety_points_addxy(string x,string y,Form1 self) {
                if (self.InvokeRequired)
                {
                    self.Invoke((MethodInvoker)delegate () { this.safety_points_addxy(x, y, self); });
                    return;
                }
                else {
                    this.graphdata.Points.AddXY(x,y);
                }
            }
            public static List<MyPoints> make_data() {
                double dx = Math.PI / 180.0;
                MyPoints tmp;
                var result = new List<MyPoints>();
                for (int i = 0; i < 360; i++) {
                    tmp = new MyPoints(i * dx, Math.Sin(i * dx));
                    result.Add(tmp);
                }
                return result;
            }
            public Series which_main_series() {
                return this.graphdata;
            }

        }

        class MyCom {
            private string PortName;
            private int BaudRate = 19200;
            private Parity _Parity = Parity.None;
            private int DataBit = 8;
            StopBits _StopBits = StopBits.One;
            private SerialPort myPort;
            private byte[] send_data;

            public MyCom(string com) {
                this.PortName = com;    //comが選択orない場合、エラーハンドリングが必要
                this.myPort = new SerialPort(PortName, BaudRate, _Parity, DataBit, _StopBits);
                string tmp = @"#01RDIG";
                byte tmpb = 0x0d;
                byte[] data = Encoding.GetEncoding("shift_jis").GetBytes(tmp);
                byte[] data1 = new byte[data.Length + 1];
                for (int i = 0; i < data.Length; i++)
                {
                    data1[i] = data[i];
                }
                data1[data.Length] = tmpb;
                this.send_data = data1;
            }


            public string Manager() {       //高速化可能

                this.myPort.Open();
                myPort.Write(this.send_data, 0, this.send_data.Length); //send部分



                Thread.Sleep(98);   //送るのと受信にかかる時間
                //Read部分
                int rbyte = myPort.BytesToRead;
                int read = 0;
                byte[] buffer = new byte[13];
                while (read < rbyte) {
                    int length = myPort.Read(buffer, read, rbyte - read);
                    read += length;

                }
                //変換
                string result = Encoding.GetEncoding("shift_jis").GetString(buffer);
                myPort.Close();
                string[] tmp_str = result.Split(' ');
                if (tmp_str.Length > 1)
                {
                    result = tmp_str[1];
                }
                else {
                    result = "0";       //エラー回避
                }
                //仮に2.23E-6を出しておく(return result)
                result = result.Replace("\r","").Replace("\n","");
                return result;      //改行文字入っている可能性

            }


        }
        class My_Read_struct {
            public DateTime dt;
            public Dictionary<string,string> dc;
            public My_Read_struct(DateTime _a,Dictionary<string,string> _b) {
                this.dt = _a;
                this.dc = _b;
            }
        }
        class MyIOmanager{  //書き込み読み込みクラス
            public static void Write_CSV(DateTime dt,Dictionary<string,string> dc,string save_name) {
                try
                {
                    using (StreamWriter sw = new StreamWriter(save_name, false, Encoding.GetEncoding("shift_jis")))
                    {
                        sw.WriteLine(dt.ToString());
                        foreach (var val_dic in dc)
                        {
                            sw.WriteLine(val_dic.Key + "," + val_dic.Value);
                        }

                    }
                }
                catch (Exception e) {
                    MessageBox.Show(e.Message + "\n書き込めませんでした");
                }
                
            }
            public static My_Read_struct Read_CSV(string open_file_name) {  //ファイルが違った場合のエラーハンドリングありだが弱い(もし違うファイルを選ぶと落ちる)
                var tmpdc = new Dictionary<string, string>();
                DateTime result_dt = new DateTime();    //必ず何か入れなきゃだから
                try
                {
                    using (StreamReader sr = new StreamReader(open_file_name, Encoding.GetEncoding("shift_jis")))
                    {
                        int i = 0;
                        
                        while (sr.Peek() != -1) {
                            string[] tmp_str = null;
                            if (i==0) {
                                string rs = sr.ReadLine();
                                result_dt = DateTime.Parse(rs);
                                i++;
                                continue;
                            }
                            tmp_str = sr.ReadLine().Split(',');
                            if (tmp_str != null && tmp_str.Length != 2) {
                                if (tmp_str.Length == 1 && tmp_str[0] == "") {  //空白行は読み飛ばし
                                    continue;
                                }
                                throw new Exception("ファイルがフォーマットに沿ってない可能性があります");
                            }
                            tmpdc[tmp_str[0]] = tmp_str[1];
                            
                        }

                    }
                }
                catch(Exception e) {
                    MessageBox.Show(e.Message + "\n読み込めませんでした");
                    return null;
                }
                return new My_Read_struct(result_dt,tmpdc);
            }
            public static void Write_excel(DateTime dt,Dictionary<string,string> dc,string save_name) {
                XLWorkbook workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Sheet1");
                worksheet.Cell(1, 1).Value = dt.ToString();
                worksheet.Cell(2,1).Value = "Time(second)";
                worksheet.Cell(2,2).Value = "pressure(Pa)";
                int i = 3;
                foreach (var value in dc) {
                    worksheet.Cell(i,1).Value = value.Key;
                    worksheet.Cell(i,2).Value = value.Value;
                    i++;
                }
                workbook.SaveAs(save_name);
                workbook.Dispose();
                worksheet.Dispose();
            }
            /*public static My_Read_struct Read_excel(string open_file_name) {
                try
                {
                    XLWorkbook book = new XLWorkbook(open_file_name);
                    IXLTable tmp = book.Worksheet(1).RangeUsed().AsTable();
                    var first_row = tmp.Rows().ToList();
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return null;
                }
                
            }
            */
        }

        //メニューバー
        private string open_file_name;
        private string save_as;
        private List<ToolStripMenuItem> my_coms = new List<ToolStripMenuItem>();
        private string selected_com;
        
        private void toolstrip_open_click(object sender,EventArgs e) { //メニューバー開く
            openFileDialog1.FileName = "";  //気を付けて後で不具合につながるかも
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK) {
                open_file_name = openFileDialog1.FileName;
                if (this.button1_flag == true)
                {
                    MessageBox.Show("STOPボタンを押してからこの操作をしてください");
                }
                else {
                    DialogResult my_result = MessageBox.Show("今のグラフに描かれているデータは消えますがよろしいですか?", "question", MessageBoxButtons.YesNo);
                    if (my_result == DialogResult.Yes)
                    {                   //すべて初期化(エラー起きやすい)
                        this.my_all_data.Clear();
                        //this.chart1.Series.Remove(this.class_my_gra.which_main_series());
                        //this.class_my_gra = new MyGraph(this.chart1);
                        this._cancellationtoken.Cancel();
                        Thread.Sleep(300);
                        this._cancellationtoken.Dispose();
                        this._cancellationtoken = new CancellationTokenSource();
                        this.parent_manager_called = false;
                        var my_tmp = Form1.MyIOmanager.Read_CSV(this.open_file_name);
                        this.class_my_starttime = my_tmp.dt;
                        this.class_my_gra = new MyGraph(chart1);
                        this.class_my_gra.Plot(my_tmp, this);
                        this.my_all_data = my_tmp.dc;
                        
                        
                    }
                    else
                    {
                        return;
                    }
                }
            }
        }
       
        private void toolstrip_save_as_click(object sender,EventArgs e) { //メニューバー名前を付けて保存
            if (checkBox1.Checked == true) {
                //自動的に保存
                try
                {
                    System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(@".\Chamber_Pressure");
                    string save_file_tmp = System.IO.Directory.GetCurrentDirectory() + @"\Chamber_Pressure\" + this.class_my_starttime.ToString().Replace("/","_").Replace(" ","_").Replace(":","_") + ".xlsx";
                    
                    //Form1.MyIOmanager.Write_CSV(this.class_my_starttime,this.my_all_data,save_file_tmp);
                    Form1.MyIOmanager.Write_excel(this.class_my_starttime,this.my_all_data,save_file_tmp);
                }
                catch (Exception ex){
                    MessageBox.Show("アクセス権限がありません\n" + ex.Message);
                }
                return;
            }
            DialogResult dr = saveFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                save_as = saveFileDialog1.FileName;
               // Form1.MyIOmanager.Write_CSV(this.class_my_starttime,this.my_all_data,save_as);
                Form1.MyIOmanager.Write_excel(this.class_my_starttime,this.my_all_data,save_as);
                MessageBox.Show("保存完了");

            }
            
        }
        private void toolstrip_save_samename_click(object sender,EventArgs e) {//メニューバー上書き保存
            if (save_as != null)
            {
                //一度名前を付けて保存が押されている必要あり
                Form1.MyIOmanager.Write_CSV(this.class_my_starttime, this.my_all_data,this.save_as);
                MessageBox.Show("保存完了");
            }
            else {
                名前を付けて保存AToolStripMenuItem.PerformClick();
            }
        }

        private void toolstrip_com_click(object sender,EventArgs e) {   //メニューバーcom選択
            foreach (ToolStripMenuItem item in this.my_coms) {
                item.Checked = object.ReferenceEquals(item, sender) ? true : false;
                if (item.Checked) {
                    selected_com = Regex.Match(item.Text,@"COM[0-9]*").ToString();  //正規表現
                }
            }
           
        }
        private void toolstrip_exit_click(object sender,EventArgs e) {  //メニューバー閉じる
            this._cancellationtoken.Cancel();
            Thread.Sleep(1000);
            this.Close();
        }

        
        public int? class_plot_xmin = null; //form2から値を受け渡された場合これらに入る
        public int? class_plot_xmax = null;
        public double? class_plot_ymin = null;
        public double? class_plot_ymax = null;

        private void toolstrip_plot_config(object sender,EventArgs e) {
            Form2 f = new Form2(this);
            f.ShowDialog(this);
        }
        private void toolstrip_integral(object sender,EventArgs e) {
            Form3 f = new Form3(this);
            f.ShowDialog();
        }
    }
}
