using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
//using System.Data.SqlClient;
using System.Timers;
using System.IO;
using System.Configuration;
using tomlib;
using Dapper;
using System.Data.SqlClient;

namespace form
{
    public partial class Form1 : Form
    {
        string startDate;
        string endDate;
        int hour = 09; //設定時間 0-23 hour
        int minute = 06; //設定時間 0-59 minute
        List<string[]> be5 = new List<string[]>();
        List<double[]> be5n = new List<double[]>();
        public Form1()
        {
            InitializeComponent();
            TimerInsert();
            timer1.Enabled = true;
        }
        class LogRecord 
        {
            public static void WriteLog(string message)
            {
                string DIRNAME = Application.StartupPath + @"\Log\";
                string FILENAME = DIRNAME + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";

                if (!Directory.Exists(DIRNAME))
                    Directory.CreateDirectory(DIRNAME);

                if (!File.Exists(FILENAME))
                {
                    File.Create(FILENAME).Close();
                }
                using (StreamWriter sw = File.AppendText(FILENAME))
                {
                    Log(message, sw);
                }
            }
            private static void Log(string logMessage, TextWriter w)
            {
                w.Write("\r\nLog Enter:");
                w.WriteLine(" {0}  {1}  {2}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString(), logMessage);
            }

            private static void DumpLog(StreamReader r)
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    Console.WriteLine(line);
                }
            }
        }       


        public void TimerInsert()
        {
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Enabled = true;
            timer.Interval = 6000; //每幾毫秒執行一次
            timer.Start();
            timer.Elapsed += new System.Timers.ElapsedEventHandler(Insertbe5);
            //timer.Elapsed += new System.Timers.ElapsedEventHandler(Insertbe6);
        }

        public void Insertbe5(object source, ElapsedEventArgs e)
        {
            string connetionString = @"Data Source=202.132.16.178,5050;Initial Catalog=intranet2;User ID=intramgr;Password=tymis1224";
            SqlConnection conn;
            conn = new SqlConnection(connetionString);
            conn.Open();
            SqlCommand cmd;

            cmd = new SqlCommand("select count(*) from be5", conn);
            Int32 countall = (Int32)cmd.ExecuteScalar();
            conn.Close();

            bool time;
            if (DateTime.Now.Hour == hour && DateTime.Now.Minute == minute) 
                time = true;
            else
                time = false;

            if (time == true)
            {
                string url = @"https://api.tomsworld.com.cn/api/legacy/erp/1?sd=2020-10-01&ed=2021-10-01";

                var json = new WebClient().DownloadString(url);
                var Datajson = JsonConvert.DeserializeObject(json);
                string ja = Convert.ToString(Datajson);
                JObject o = JObject.Parse(ja);
                JArray jar = JArray.Parse(o["data"].ToString().Trim());

                int length = jar.Count;
                string test = length.ToString();
                int i = 0;
                for (i = countall; i < length; i++)
                {
                    JObject j = JObject.Parse(jar[i].ToString().Trim());
                    string bb1no = j["bb1no"].ToString().Trim();
                    string pl1no = j["pl1no"].ToString().Trim();
                    string online_date = j["online_date"].ToString().Trim();
                    string num = j["num"].ToString().Trim();
                    string price = j["price"].ToString().Trim();
                    string YM = j["online_date"].ToString().Trim().Substring(0, 7);
                    string YMdate = YM.Replace("-", "");
                    
                   
                    conn.Open();
                    SqlTransaction tran;
                    JArray jA = new JArray();
                    JObject jO1 = new JObject();

                    //cmd = new SqlCommand("select bb1no from be5 where bb1no = @bb1no", conn);
                    //cmd.Parameters.AddWithValue("@bb1no", bb1no);
                    //dr = cmd.ExecuteReader();

                    //if (dr.HasRows)
                    //{
                    //    dr.Close();
                    //    cmd = new SqlCommand("update be5 set bb1no=@bb1no,pl1no=@pl1no,bh1ym2=@YMdate,bh1date=@online_date,be1qty=@num,be1price=@price where bb1no=@bb1no", conn);
                    //}
                    //else
                    //{
                    //dr.Close();
                    tran = conn.BeginTransaction();
                    cmd = new SqlCommand("insert into be5(bb1no,pl1no,bh1ym2,bh1date,be1qty,be1price) values(@bb1no,@pl1no,@YMdate,@online_date,@num,@price)", conn);
                    cmd.Transaction = tran;
                    tran.Commit();
                    tran = null;

                    //}
                    cmd.Parameters.AddWithValue("@bb1no", bb1no);
                    cmd.Parameters.AddWithValue("@pl1no", pl1no);
                    cmd.Parameters.AddWithValue("@YMdate", YMdate);
                    cmd.Parameters.AddWithValue("@online_date", online_date);
                    cmd.Parameters.AddWithValue("@num", num);
                    cmd.Parameters.AddWithValue("@price", price);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
                LogRecord.WriteLog(": 新增be5成功");
            }
            else
            {
                if (time == true)
                {
                    LogRecord.WriteLog(": 新增be5失敗");
                }
            }
        }
        public void Insertbe6(object source, ElapsedEventArgs e)
        {
            bool time;
            if (DateTime.Now.Hour == hour && DateTime.Now.Minute == minute)
                time = true;
            else
                time = false;

            if (time == true)
            {
                //string url = @"https://api.tomsworld.com.cn/api/legacy/erp/2?sd=2020-10-01&ed=2021-10-01";

                //var json = new WebClient().DownloadString(url);
                
                var json = "{'success':true,'data':[{'pl1no':'CZJM','online_date':'2021-01-07','num':'100','price':'0.00'}]}";
                var Datajson = JsonConvert.DeserializeObject(json);
                string ja = Convert.ToString(Datajson);
                JObject o = JObject.Parse(ja);
                JArray jar = JArray.Parse(o["data"].ToString().Trim());

                int length = jar.Count;
                int i = 0;
                for (i = 0; i < length; i++)
                {
                    JObject j = JObject.Parse(jar[i].ToString().Trim());
                    string pl1no = j["pl1no"].ToString().Trim();
                    string online_date = j["online_date"].ToString().Trim();
                    string num = j["num"].ToString().Trim();
                    string price = j["price"].ToString().Trim();
                    string YM = j["online_date"].ToString().Trim().Substring(0, 7);
                    string YMdate = YM.Replace("-", "");
                    string connetionString;
                    SqlConnection conn;
                    connetionString = @"Data Source=192.168.1.51;Initial Catalog=intranet2;User ID=sa;Password=23460371";
                    conn = new SqlConnection(connetionString);
                    conn.Open();
                    SqlCommand cmd;
                    //SqlDataReader dr;
                    SqlTransaction tran;
                    JArray jA = new JArray();
                    JObject jO1 = new JObject();

                    tran = conn.BeginTransaction();
                    cmd = new SqlCommand("insert into be6(pl1no,bh1ym2,bh1date,be1qty,be1price) values(@pl1no,@YMdate,@online_date,@num,@price)", conn);
                    cmd.Transaction = tran;
                    tran.Commit();
                    tran = null;

                    cmd.Parameters.AddWithValue("@pl1no", pl1no);
                    cmd.Parameters.AddWithValue("@YMdate", YMdate);
                    cmd.Parameters.AddWithValue("@online_date", online_date);
                    cmd.Parameters.AddWithValue("@num", num);
                    cmd.Parameters.AddWithValue("@price", price);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
                LogRecord.WriteLog(": 新增be6成功");
            }
            else
            {
                if (time == true)
                {
                    LogRecord.WriteLog(": 新增be6失敗");
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
//            string connetionString;
            string strDb = ConfigurationManager.ConnectionStrings["DbData2"].ConnectionString;
            IDbConnection conn;
            IDataReader dr;
            DynamicParameters para;
            DataTable dt = new DataTable();
            //connetionString = @"Data Source=192.168.1.51;Initial Catalog=intranet2;User ID=sa;Password=23460371";
            conn = new SqlConnection(strDb);
            conn.Open();

            if (startDate != null & endDate != null)
            {
                string selectbe = "select [pl1no],[bh1date],[bb1no],[be1qty],[be1price],[bh1ym2] from be5 where bh1date >= @startDate and  bh1date <= @endDate";
                para = new DynamicParameters();
                para.Add("@startDate", startDate);
                para.Add("@endDate", endDate);
                dr = conn.ExecuteReader(selectbe, para);
                dt.Load(dr, LoadOption.OverwriteChanges);
                dataGridView1.DataSource = dt;
                //using (SqlCommand cmd = new SqlCommand(selectbe, conn))
                //{
                //    cmd.Parameters.AddWithValue("@startDate", startDate);
                //    cmd.Parameters.AddWithValue("@endDate", endDate);
                //using (SqlDataReader dr = conn.ExecuteReader())
                //{
                //    DataTable dt = new DataTable();
                //    dt.Load(dr, LoadOption.OverwriteChanges);
                //    dataGridView1.DataSource = dt;
                //}
                //}
            }
            else
            {
                //string selectbe = "select [pl1no],[bh1date],[bb1no],[be1qty],[be1price],[bh1ym2] from be5 ";
                //using (SqlCommand cmd = new SqlCommand(selectbe, conn))
                //{
                //    using (SqlDataReader dr = cmd.ExecuteReader())
                //    {
                //        DataTable dt = new DataTable();
                //        dt.Load(dr, LoadOption.OverwriteChanges);
                //        dataGridView1.DataSource = dt;
                //    }
                //}
            }
            dataGridView1.Columns["pl1no"].HeaderText = "店號";
            dataGridView1.Columns["bh1date"].HeaderText = "日期";
            dataGridView1.Columns["bb1no"].HeaderText = "機台代號";
            dataGridView1.Columns["be1qty"].HeaderText = "點數";
            dataGridView1.Columns["be1price"].HeaderText = "幣值";
            dataGridView1.Columns["bh1ym2"].HeaderText = "年月";
            conn.Close();
        }
        private void button2_Click(object sender, EventArgs e)//------select
        {
            string strDb = ConfigurationManager.ConnectionStrings["DbData2"].ConnectionString;
            IDbConnection conn;
            IDataReader dr;
            DynamicParameters para;
            DataTable dt = new DataTable();
            DateTime bt, et;
            try
            {
                bt = dateTimePicker1.Value;
                et = dateTimePicker2.Value;
                conn = new SqlConnection(strDb);
                conn.Open();

                string selectbe = "select pl1no,[bh1ym2],[bh1date],[be1qty],[be1price] from be6 where bh1date >= @startDate and  bh1date <= @endDate";
                para = new DynamicParameters();
                para.Add("@startDate", bt);
                para.Add("@endDate", et);
                dr = conn.ExecuteReader(selectbe, para);
                dt.Load(dr, LoadOption.OverwriteChanges);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns["pl1no"].HeaderText = "店號";
                dataGridView1.Columns["bh1date"].HeaderText = "日期";
                dataGridView1.Columns["be1qty"].HeaderText = "點數";
                dataGridView1.Columns["be1price"].HeaderText = "幣值";
                dataGridView1.Columns["bh1ym2"].HeaderText = "年月";
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            startDate = dateTimePicker1.Value.ToString("yyyy-MM-dd");
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            endDate = dateTimePicker2.Value.ToString("yyyy-MM-dd");
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void label2_Click(object sender, EventArgs e)
        {
            
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToString("yyyy年MM月dd日 HH:mm:ss");
            timer1.Interval = 100;
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            save_data1();

        }
        public void save_data1()
        {
            SqlConnection conn;
            int i,n;
            string url;
            SqlTransaction tran;
            DynamicParameters para;
            string pl1no,bb1no,online_date,YM,YMdate,ja;
            string strDb = ConfigurationManager.ConnectionStrings["DbData2"].ConnectionString;
            double be1qty, be1price;
            DateTime dt;
            JObject j,o;
            JArray jar;

            try
            {
                conn = new SqlConnection(strDb);
                conn.Open();
                tran = conn.BeginTransaction();
                //----小程序營收
                url = @"https://api.tomsworld.com.cn/api/legacy/erp/1?sd=2021-01-26&ed=2021-10-01";

                var json = new WebClient().DownloadString(url);
                var Datajson = JsonConvert.DeserializeObject(json);
                ja = Convert.ToString(Datajson);
                o = JObject.Parse(ja);
                jar = JArray.Parse(o["data"].ToString().Trim());
                for (i = 0; i < jar.Count; i++)
                {
                    j = JObject.Parse(jar[i].ToString().Trim());
                    pl1no = j["pl1no"].ToString().Trim();
                    online_date = j["online_date"].ToString().Trim();
                    para = new DynamicParameters();
                    para.Add("@pl1no", pl1no);
                    para.Add("@bh1date", online_date);
                    conn.Execute("delete from be5 where pl1no=@pl1no and bh1date=@bh1date", para, tran);
                }
                for (i = 0; i < jar.Count; i++)
                {
                    j = JObject.Parse(jar[i].ToString().Trim());
                    pl1no = j["pl1no"].ToString().Trim();
                    bb1no = j["bb1no"].ToString().Trim();
                    online_date = j["online_date"].ToString().Trim();
                    be1qty = Convert.ToDouble(j["num"].ToString());
                    be1price = Convert.ToDouble(j["price"].ToString());
                    YM = j["online_date"].ToString().Trim().Substring(0, 7);
                    YMdate = YM.Replace("-", "");

                    save_Dm(pl1no, online_date, be1qty, be1price);
                    para = new DynamicParameters();
                    para.Add("@pl1no", pl1no);
                    para.Add("@bb1no", bb1no);
                    para.Add("@bh1date", online_date);
                    para.Add("@YMdate", YMdate);
                    para.Add("@be1qty", be1qty);
                    para.Add("@be1price", be1price);
                    conn.Execute("insert into be5 (bb1no,pl1no,bh1ym2,bh1date,be1qty,be1price) values(@bb1no,@pl1no,@YMdate,@bh1date,@be1qty,@be1price)", para, tran);
                }
                n = 0;
                foreach (string[] s1 in be5)
                {
                    dt = Convert.ToDateTime(s1[1]);
                    var in2m = conn.Query<in2M>("select in1num,pl1no,bk1date,bk1wek from in2 where pl1no=@pl1no and bk1date=@bk1date", new { pl1no = s1[0], bk1date = dt }, tran).ToList();
                    if (in2m.Count == 0)
                    {
                        para = new DynamicParameters();
                        para.Add("@pl1no", s1[0]);
                        para.Add("@bh1date", s1[1]);
                        para.Add("@bk1wek", TomLib.getWeek(dt));
                        para.Add("@be1qty", be5n[n][0]);
                        para.Add("@be1price", be5n[n][1]);
                        conn.Execute("insert into in2 (pl1no,bk1date,bk1wek,be1qty,be1price) values (@pl1no,@bh1date,@bk1wek,@be1qty,@be1price)", para, tran);
                    }
                    else
                    {
                        para = new DynamicParameters();
                        para.Add("@in1num", in2m[0].in1num);
                        para.Add("@be1qty", be5n[n][0]);
                        para.Add("@be1price", be5n[n][1]);
                        conn.Execute("update in2 set be1qty=@be1qty,be1price=@be1price where in1num=@in1num", para, tran);
                    }
                    n++;
                }
                tran.Commit();
                tran = null;
                conn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show("OK !!");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            save_data();
        }
        public void save_data()
        {
            SqlConnection conn;
            SqlTransaction tran;
            DynamicParameters para;
            string pl1no,bb1no, online_date, YM, YMdate,url,ja;
            string strDb = ConfigurationManager.ConnectionStrings["DbData2"].ConnectionString;
            double be1qty, be1price,be1price2;
            int length, i,n;
            DateTime dt;
            string bdate = "", edate = "";
            JObject o, j;
            JArray jar;

            bdate = "2021/01/01";
            edate = DateTime.Now.ToShortDateString();

//            MessageBox.Show(url);
//            return;
            try
            {
                conn = new SqlConnection(strDb);
                conn.Open();
                tran = conn.BeginTransaction();
                //----小程序營收,非特殊機台,be5---->type=1
                url = @"https://api.tomsworld.com.cn/api/legacy/erp/1?sd=" + bdate + "&ed=" + edate;

                var json = new WebClient().DownloadString(url);
                var Datajson = JsonConvert.DeserializeObject(json);
                ja = Convert.ToString(Datajson);
                o = JObject.Parse(ja);
                jar = JArray.Parse(o["data"].ToString().Trim());
                //---------------------------------------------------------------------------
                para = new DynamicParameters();
                para.Add("@bdate", bdate);
                para.Add("@edate", edate);
                conn.Execute("delete from be5 where bh1date between @bdate and @edate", para, tran);
                //for (i = 0; i < jar.Count; i++)
                //{
                //    j = JObject.Parse(jar[i].ToString().Trim());
                //    pl1no = j["pl1no"].ToString().Trim();
                //    online_date = j["online_date"].ToString().Trim();
                //    para = new DynamicParameters();
                //    para.Add("@pl1no", pl1no);
                //    para.Add("@bh1date", online_date);
                //    conn.Execute("delete from be5 where pl1no=@pl1no and bh1date=@bh1date and be1type=''", para, tran);
                //}
                for (i = 0; i < jar.Count; i++)
                {
                    j = JObject.Parse(jar[i].ToString().Trim());
                    pl1no = j["pl1no"].ToString().Trim();
                    bb1no = j["bb1no"].ToString().Trim();
                    online_date = j["online_date"].ToString().Trim();
                    be1qty = Convert.ToDouble(j["num"].ToString());
                    be1price = Convert.ToDouble(j["price"].ToString());
                    YM = j["online_date"].ToString().Trim().Substring(0, 7);
                    YMdate = YM.Replace("-", "");

                    save_Dm(pl1no, online_date, be1qty, be1price);
                    para = new DynamicParameters();
                    para.Add("@pl1no", pl1no);
                    para.Add("@bb1no", bb1no);
                    para.Add("@bh1date", online_date);
                    para.Add("@YMdate", YMdate);
                    para.Add("@be1qty", be1qty);
                    para.Add("@be1price", be1price);
                    conn.Execute("insert into be5 (bb1no,pl1no,bh1ym2,bh1date,be1qty,be1price,be1type) values(@bb1no,@pl1no,@YMdate,@bh1date,@be1qty,@be1price,'')", para, tran);
                }
                //----小程序營收,特殊機台-->type=3
                url = @"https://api.tomsworld.com.cn/api/legacy/erp/3?sd=2021-01-26&ed=2021-10-01";

                json = new WebClient().DownloadString(url);
                Datajson = JsonConvert.DeserializeObject(json);
                ja = Convert.ToString(Datajson);
                o = JObject.Parse(ja);
                jar = JArray.Parse(o["data"].ToString().Trim());
                for (i = 0; i < jar.Count; i++)
                {
                    j = JObject.Parse(jar[i].ToString().Trim());
                    pl1no = j["pl1no"].ToString().Trim();
                    bb1no = j["bb1no"].ToString().Trim();
                    online_date = j["online_date"].ToString().Trim();
                    be1qty = Convert.ToDouble(j["num"].ToString());
                    be1price = Convert.ToDouble(j["price_1"].ToString());
                    be1price2 = Convert.ToDouble(j["price_2"].ToString());
                    YM = j["online_date"].ToString().Trim().Substring(0, 7);
                    YMdate = YM.Replace("-", "");

                    save_Dm(pl1no, online_date, be1qty, be1price+be1price2);
                    para = new DynamicParameters();
                    para.Add("@pl1no", pl1no);
                    para.Add("@bb1no", bb1no);
                    para.Add("@bh1date", online_date);
                    para.Add("@YMdate", YMdate);
                    para.Add("@be1qty", be1qty);
                    para.Add("@be1price", be1price); //----小程序營收,特殊機台-->type=3
                    para.Add("@be1price2", be1price2); //----兌換券金額
                    conn.Execute("insert into be5 (bb1no,pl1no,bh1ym2,bh1date,be1qty,be1price,be1price2,be1type) values(@bb1no,@pl1no,@YMdate,@bh1date,@be1qty,@be1price,@be1price2,'Y')", para, tran);
                }
                //------------------------------------------------------------------------------------
                //-------兌幣機-be6,be1type="",--->type=2
                url = @"https://api.tomsworld.com.cn/api/legacy/erp/2?sd="+bdate+"&ed="+edate;
                var json2 = new WebClient().DownloadString(url);

                var Datajson2 = JsonConvert.DeserializeObject(json2);
                ja = Convert.ToString(Datajson2);
                o = JObject.Parse(ja);
                jar = JArray.Parse(o["data"].ToString().Trim());
                length = jar.Count;
                //--------------------刪除舊資料
                para = new DynamicParameters();
                para.Add("@bdate", bdate);
                para.Add("@edate", edate);
                conn.Execute("delete from be6 where bh1date between @bdate and @edate", para, tran);
                //---------------------先把營收歸=0----------------------------------------
                para = new DynamicParameters();
                para.Add("@bdate", bdate);
                para.Add("@edate", edate);
                conn.Execute("update in2 set be1qty=0,be1price=0 where bk1date between @bdate and @edate", para, tran);

                for (i = 0; i < length; i++)
                {
                    j = JObject.Parse(jar[i].ToString().Trim());
                    pl1no = j["pl1no"].ToString().Trim();
                    online_date = j["online_date"].ToString().Trim();
                    be1qty = Convert.ToDouble(j["num"].ToString());
                    be1price = Convert.ToDouble(j["price"].ToString());
                    YM = j["online_date"].ToString().Trim().Substring(0, 7);
                    YMdate = YM.Replace("-", "");
                    if (pl1no != null && pl1no.Trim() != "")
                    {
                        save_Dm(pl1no,online_date, be1qty, be1price); //---統計營收
                        para = new DynamicParameters();
                        para.Add("@pl1no", pl1no);
                        para.Add("@YMdate", YMdate);
                        para.Add("@bh1date", online_date);
                        para.Add("@be1qty", be1qty);
                        para.Add("@be1price", be1price);//----兌幣機-be6,be1type="",--->type=2
                        conn.Execute("insert into be6 (pl1no,bh1ym2,bh1date,be1qty,be1price,be1type) values (@pl1no,@YMdate,@bh1date,@be1qty,@be1price,'')", para, tran);
                    }
                }
                n = 0;
                //foreach (string[] s1 in be5)
                //{
                //    dt = Convert.ToDateTime(s1[1]);
                //    var in2m = conn.Query<in2M>("select in1num,pl1no,bk1date,bk1wek from in2 where pl1no=@pl1no and bk1date=@bk1date", new { pl1no = s1[0], bk1date = dt },tran).ToList();
                //    if (in2m.Count==0)
                //    {
                //        para = new DynamicParameters();
                //        para.Add("@pl1no", s1[0]);
                //        para.Add("@bh1date", s1[1]);
                //        para.Add("@bk1wek", TomLib.getWeek(dt));
                //        para.Add("@be1qty", be5n[n][0]);
                //        para.Add("@be1price", be5n[n][1]);
                //        conn.Execute("insert into in2 (pl1no,bk1date,bk1wek,be1qty,be1price) values (@pl1no,@bh1date,@bk1wek,@be1qty,@be1price)", para, tran);
                //    }
                //    else
                //    {
                //        para = new DynamicParameters();
                //        para.Add("@in1num", in2m[0].in1num);
                //        para.Add("@be1qty", be5n[n][0]);
                //        para.Add("@be1price", be5n[n][1]);
                //        conn.Execute("update in2 set be1qty=be1qty+@be1qty,be1price=be1price+@be1price where in1num=@in1num", para, tran);
                //    }
                //    n++;
                //}
                //-------櫃台核銷be6,be1type="Y"-->type=4
                url = @"https://api.tomsworld.com.cn/api/legacy/erp/4?sd=" + bdate + "&ed=" + edate;
                json2 = new WebClient().DownloadString(url);

                Datajson2 = JsonConvert.DeserializeObject(json2);
                ja = Convert.ToString(Datajson2);
                o = JObject.Parse(ja);
                jar = JArray.Parse(o["data"].ToString().Trim());

                length = jar.Count;
                for (i = 0; i < length; i++)
                {
                    j = JObject.Parse(jar[i].ToString().Trim());
                    pl1no = j["pl1no"].ToString().Trim();
                    online_date = j["online_date"].ToString().Trim();
                    be1qty = Convert.ToDouble(j["num"].ToString());
                    be1price = Convert.ToDouble(j["price"].ToString());
                    YM = j["online_date"].ToString().Trim().Substring(0, 7);
                    YMdate = YM.Replace("-", "");
                    if (pl1no != null && pl1no.Trim() != "")
                    {
                        save_Dm(pl1no, online_date, be1qty, be1price);
                        para = new DynamicParameters();
                        para.Add("@pl1no", pl1no);
                        para.Add("@YMdate", YMdate);
                        para.Add("@bh1date", online_date);
                        para.Add("@be1qty", be1qty);
                        para.Add("@be1price", be1price);//-------櫃台核銷be6,be1type="Y"-->type=4
                        conn.Execute("insert into be6 (pl1no,bh1ym2,bh1date,be1qty,be1price,be1type) values (@pl1no,@YMdate,@bh1date,@be1qty,@be1price,'Y')", para, tran);
                    }
                }
                n = 0;
                //--------------把資料存到營收檔-----------------------------------------------
                foreach (string[] s1 in be5)
                {
                    dt = Convert.ToDateTime(s1[1]);
                    var in2m = conn.Query<in2M>("select in1num,pl1no,bk1date,bk1wek from in2 where pl1no=@pl1no and bk1date=@bk1date", new { pl1no = s1[0], bk1date = dt }, tran).ToList();
                    if (in2m.Count == 0)
                    {
                        para = new DynamicParameters();
                        para.Add("@pl1no", s1[0]);
                        para.Add("@bh1date", s1[1]);
                        para.Add("@bk1wek", TomLib.getWeek(dt));
                        para.Add("@be1qty", be5n[n][0]);
                        para.Add("@be1price", be5n[n][1]);
                        conn.Execute("insert into in2 (pl1no,bk1date,bk1wek,be1qty,be1price) values (@pl1no,@bh1date,@bk1wek,@be1qty,@be1price)", para, tran);
                    }
                    else
                    {
                        para = new DynamicParameters();
                        para.Add("@in1num", in2m[0].in1num);
                        para.Add("@be1qty", be5n[n][0]);
                        para.Add("@be1price", be5n[n][1]);
                        conn.Execute("update in2 set be1qty=@be1qty,be1price=@be1price where in1num=@in1num", para, tran);
                    }
                    n++;
                }
                tran.Commit();
                tran = null;
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show("OK !!");
        }
        public void save_Dm(string pl1no,string bh1date, double be1qty, double be1price)
        {
            int n=0;
            //DateTime dt2;

            //dt2 = Convert.ToDateTime(bh1date);
            foreach (string[] sc in be5)
            {
                if (sc[1] == bh1date && sc[0]==pl1no)
                {
                    be5n[n][0] += be1qty;
                    be5n[n][1] += be1price;
                    return;
                }
                n++;
            }
            //MessageBox.Show(dt2.ToShortDateString());
            be5.Add(new string[] { pl1no,bh1date });
            be5n.Add(new double[] { be1qty, be1price });
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string strDb = ConfigurationManager.ConnectionStrings["DbData2"].ConnectionString;
            DateTime dt= Convert.ToDateTime("2021/2/9");
            int n;
            //DbDataDataContext db2 = new DbDataDataContext(strDb);

            try
            {
                var conn = new SqlConnection(strDb);
                //conn.Open();
                var in2m = conn.Query<in2M>("select in1num,pl1no,bk1date,bk1wek from in2 where pl1no=@pl1no and bk1date=@bk1date", new { pl1no = "CCC",bk1date=dt }).ToList();
                //for (n = 0; n < in2m.Count; n++)
                //{
                //    MessageBox.Show(in2m[n].in1num.ToString());
                //}
                in2m.ForEach(mm => MessageBox.Show(mm.pl1no));
                //conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public class in2M
        {
            public Int32 in1num { get; set; }
            public string pl1no { get; set; }
            public DateTime bk1date { get; set; }
            //public string bk1wek { get; set; }
            public override string ToString()
            {
                return "{pl1no}{in1num}";
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            MessageBox.Show(TomLib.getWeek(dt));
        }
    }
}
