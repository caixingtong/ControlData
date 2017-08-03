using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.ListViewItem;

namespace ControlData
{
    public partial class Form1 : Form
    {

        public static bool IsContinue = true;
        public static object _Lock = new object();
        FileSystemWatcher watcher = new FileSystemWatcher();
        public Form1()
        {
            InitializeComponent();
            this.button2.Enabled = false;
            LoadListView();
            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
            this.panel1.Visible = false;
            GetProvinceName(this.cbx_province);
        }
        /// <summary>
        /// 初始化上传列表
        /// </summary>
        void LoadListView()
        {
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.Columns.Add("状态", 150, HorizontalAlignment.Center);
            listView1.Columns.Add("文件名", 300, HorizontalAlignment.Center);
            listView1.Columns.Add("路径", 150, HorizontalAlignment.Center);
            listView1.Columns.Add("地区", 150, HorizontalAlignment.Center);
            listView1.Columns.Add("长", 150, HorizontalAlignment.Center);
            listView1.Columns.Add("宽", 150, HorizontalAlignment.Center);
            listView1.Columns.Add("厚度", 150, HorizontalAlignment.Center);
            listView1.Columns.Add("面积", 150, HorizontalAlignment.Center);
            listView1.Columns.Add("时间", 150, HorizontalAlignment.Center);
        }
        /// <summary>
        /// 监视开启按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            IsContinue = true;
            if (this.textBox1.Text == "")
            {
                MessageBox.Show("请填写监控路径!");
                return;
            }
            if (this.textBox2.Text == "")
            {
                MessageBox.Show("请填写备份路径!");
                return;
            }
            this.button1.Enabled = false;
            this.button2.Enabled = true;
            IsContinue = true;
            //TaskFactory factory = new TaskFactory();
            //Action act = (() =>
            //{
            //    TestFileSystemWatcher(textBox1);
            //});
            //factory.StartNew(act);

            //  Task task = Task.Run(() =>
            //  {
            TestFileSystemWatcher(textBox1);
            //  });

            //new Action(() =>
            //{

            //}).BeginInvoke(null, null);
            //TestFileSystemWatcher(textBox1);
        }
        private void button5_Click(object sender, EventArgs e)
        {
            this.textBox1.Text = ReturnPath(this.textBox1);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.textBox2.Text = ReturnPath(this.textBox2);
        }

        public string ReturnPath(TextBox textbox)
        {
            string defaultPath = textbox.Text;
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择一个文件夹";
            //是否显示对话框左下角 新建文件夹 按钮，默认为 true  
            dialog.ShowNewFolderButton = false;
            //首次defaultPath为空，按FolderBrowserDialog默认设置（即桌面）选择  
            if (defaultPath != "")
            {
                //设置此次默认目录为上一次选中目录  
                dialog.SelectedPath = defaultPath;
            }
            //按下确定选择的按钮  
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //记录选中的目录  
                defaultPath = dialog.SelectedPath;
            }
            return defaultPath;
        }
        public void TestFileSystemWatcher(TextBox textbox)
        {
            try
            {
                string path = textbox.Text;
                watcher.Path = @path;
                //设置监视文件的哪些修改行为
                watcher.NotifyFilter = NotifyFilters.LastAccess
                    | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
                watcher.Filter = "*.jpg";
                //watcher.Changed += new FileSystemEventHandler(OnChanged);
                watcher.Created += new FileSystemEventHandler(OnChanged);
                watcher.Deleted += new FileSystemEventHandler(OnChanged);
                //watcher.Renamed += new RenamedEventHandler(OnRenamed);
                watcher.EnableRaisingEvents = IsContinue;
                int result = 10;
                Thread.Sleep(Int32.TryParse(this.textBox3.Text, out result) ? result : 10);
            }
            catch (ArgumentException e)
            {
                MessageBox.Show(e.Message);
                return;
            }
        }

        public void OnChanged(object source, FileSystemEventArgs e)
        {
            ShowListView(e.FullPath, e.ChangeType.ToString());
            //MessageBox.Show(string.Format("File：{0} {1}！", e.FullPath, e.ChangeType));
        }

        //static void OnRenamed(object source, RenamedEventArgs e)
        //{
        //    MessageBox.Show(string.Format("File：{0} renamed to\n{1}", e.OldFullPath, e.FullPath));
        //}


        /// <summary>
        /// 计算某字符串在另一字符串中，出现第N次的位置
        /// </summary>
        /// <param name="strA"></param>
        /// <param name="strB">匹配字符</param>
        /// <param name="n">第N次</param>
        /// <returns></returns>
        public int GetIndexN(string strA, string strB, int n)
        {
            string a = strA;
            string c;
            string find_c;
            find_c = strB;
            int count = 0;
            for (int i = 0; i < a.Length; i++)
            {
                c = a[i].ToString();
                if (c == find_c)//求得a中包含该字符的个数，以便遍历   
                {
                    count++;
                }
            }
            int index = 0;
            if (n > count)
            {
                return 0;
            }
            for (int j = 1; j <= count; j++)
            {
                index = a.IndexOf(find_c, index);
                if (j == n)
                {
                    break;
                }
                else
                {
                    index = a.IndexOf(find_c, index + 1);
                }
            }
            return index;
        }
        /// <summary>
        /// 时间戳转DateTime
        /// </summary>
        /// <param name="timestamp"></param>
        /// <returns></returns>
        private DateTime TimestampToDateTime(long timestamp)
        {
            DateTime dateTimeStart = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
            long lTime = timestamp * 10000000;
            TimeSpan nowTimeSpan = new TimeSpan(lTime);
            DateTime resultDateTime = dateTimeStart.Add(nowTimeSpan);
            return resultDateTime;
        }


        /// <summary>
        /// 执行导出数据
        /// </summary>
        public void ExportToExecl()
        {
            System.Windows.Forms.SaveFileDialog sfd = new SaveFileDialog();
            sfd.DefaultExt = "csv";
            sfd.Filter = "Excel文件|*.csv";
            //sfd.Filter = "Excel文件(*.xls)|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                if (ExportForListView(this.listView1, sfd.FileName, true))
                {
                    MessageBox.Show("导出成功！");
                }
                else
                {
                    MessageBox.Show("导出失败！");
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="listView">ListView</param>
        /// <param name="fileName"></param>
        /// <param name="isShowExcle"></param>
        /// <returns></returns>

        public static bool ExportForListView(ListView listView, string fileName, bool isShowExcle)
        {
            FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, Encoding.Unicode);
            try 
            {
                string excel = "";      //用于存放要写入的一行文本。 
                for (int i = 0; i < listView.Columns.Count; i++)
                {
                    excel = excel + listView.Columns[i].Text.ToString().Trim() + Convert.ToChar(9);
                }
                //写入DataGridView的标题行。 
                excel = "";
                for (int i = 0; i < listView.Items.Count; i++)
                {
                    for (int j = 0; j < listView.Columns.Count; j++)
                    {
                        if (listView.Items[i].SubItems[j].Text.ToString() == null)
                            excel = excel + "" + Convert.ToChar(9);    //循环写入每一行 
                        else
                            excel = excel + listView.Items[i].SubItems[j].Text.ToString() + Convert.ToChar(9);
                    }
                    sw.WriteLine(excel);
                    excel = "";
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                sw.Close();
                fs.Close();
                if (isShowExcle)
                {
                    System.Diagnostics.Process.Start(fileName);
                }
            }
            return true;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            watcher.EnableRaisingEvents = false;
            this.button2.Enabled = false;
            this.button1.Enabled = true;
        }


        public void ShowListView(string strPath, string strState)
        {
            try
            {
                ListViewItem lt = new ListViewItem();
                ListViewSubItem item1 = new ListViewSubItem(lt, strState == "Deleted" ? "删除" : "新增");
                ListViewSubItem item3 = new ListViewSubItem(lt, strPath);
                int index = strPath.LastIndexOf('\\');
                string strName = strPath.Substring(index + 1, strPath.Length - index - 5);
                ListViewSubItem item2 = new ListViewSubItem(lt, strName);
                int Length = strName.Length;
                //int index1 = strName.IndexOf("_");
                int index0 = GetIndexN(strName, "_", 3) - GetIndexN(strName, "_", 2);
                string Country = strName.Substring(GetIndexN(strName, "_", 2) + 1, index0 - 1);

                int index1 = GetIndexN(strName, "_", 3);
                int index2 = strName.IndexOf("x");
                int index3 = strName.IndexOf("x", index2 + 1);
                int index4 = strName.IndexOf("_", index3 + 1);
                string Col = strName.Substring(index1 + 1, index2 - index1 - 1);
                string Row = strName.Substring(index2 + 1, index3 - index2 - 1);
                string Land = strName.Substring(index3 + 1, index4 - index3 - 1);
                // string Areas = strName.Substring(index4 + 1, Length - index4 - 1);
                int index5 = GetIndexN(strName, "_", 5) - GetIndexN(strName, "_", 4);
                string Areas = strName.Substring(GetIndexN(strName, "_", 4) + 1, index5 - 1);

                string Time = strName.Substring(GetIndexN(strName, "_", 5) + 1, Length - GetIndexN(strName, "_", 5) - 1);
                long ltime = 0;
                ltime = long.TryParse(Time, out ltime) ? ltime : 0;

                //listView1.Items[i].Text = "长" + Col + "宽" + Row + "厚度" + Land + "面积" + Areas;      
                ListViewSubItem item4 = new ListViewSubItem(lt, Col);
                ListViewSubItem item5 = new ListViewSubItem(lt, Row);
                ListViewSubItem item6 = new ListViewSubItem(lt, Land);
                ListViewSubItem item7 = new ListViewSubItem(lt, Areas);
                ListViewSubItem item8 = new ListViewSubItem(lt, Country);
                ListViewSubItem item9 = new ListViewSubItem(lt, TimestampToDateTime(ltime).ToString());
                //ListViewSubItem item9 = new ListViewSubItem(lt, Time);
                lt.SubItems.Insert(0, item1);
                lt.SubItems.Insert(1, item2);
                lt.SubItems.Insert(2, item3);
                //地区    
                lt.SubItems.Insert(3, item8);
                lt.SubItems.Insert(4, item4);
                lt.SubItems.Insert(5, item5);
                lt.SubItems.Insert(6, item6);
                lt.SubItems.Insert(7, item7);
                lt.SubItems.Insert(8, item9);
                this.listView1.Items.Add(lt);
                //  CopyFile(this.textBox1.Text,this.textBox2.Text);

                //SaveData(this.txt_goodsname.Text, Row, Col, this.cbx_color.Text,
                //    this.cbx_area.Text, this.cbx_city.Text, "", this.cbx_quality.Text
                //    , this.txt_zj.Text, this.txt_lxr.Text, this.txt_tel.Text, strPath, "", this.cbx_wl.Text
                //    , this.txt_dj.Text, this.cbx_province.Text);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(string.Format("请确保文件名格式正确！{0}",ex.Message));
            }

        }

        /// <summary>
        ///  拷贝文件到指定文件夹 
        /// </summary>
        /// <param name="strPath">文件路径</param>
        /// <param name="strPath2">备份路径</param>
        public void CopyFile(string strPath, string strPath2)
        {
            FileInfo files = new FileInfo(strPath);
            files.CopyTo(strPath2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExecl();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (this.panel1.Visible == true)
            {
                this.panel1.Visible = false;
            }
            else
            {
                this.panel1.Visible = true;
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        public void  GetProvinceName(ComboBox cbx)
        {
            string strSql = string.Format("select provinceid,provincename from [dbo].[T_PROVINCE]");
            try
            {
                DbHelper helper = new DbHelper();
                DataTable db = helper.ExecuteDataTable(strSql);
                cbx.DataSource = db;
                cbx.DisplayMember = "provincename";
                cbx.ValueMember = "provinceid";
            }
            catch (Exception)
            {
                throw;
            }
          
        }

        public void GetCityName(ComboBox cbx,string ParentId)
        {
            string strSql = string.Format("select cityid,cityname from [dbo].[T_CITY] where parentid={0}", ParentId);
            try
            {
                DbHelper helper = new DbHelper();
                DataTable db = helper.ExecuteDataTable(strSql);
                cbx.DataSource = db;
                cbx.DisplayMember = "cityname";
                cbx.ValueMember = "cityid";
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void GetAreaName(ComboBox cbx, string ParentId)
        {
            string strSql = string.Format("select areaid,areaname from [dbo].[T_AREA]  where parentid={0}", ParentId);
            try
            {
                DbHelper helper = new DbHelper();
                DataTable db = helper.ExecuteDataTable(strSql);
                cbx.DataSource = db;
                cbx.DisplayMember = "areaname";
                cbx.ValueMember = "areaid";
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void cbx_province_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.cbx_province.ValueMember != "")
            {
                GetCityName(this.cbx_city, this.cbx_province.SelectedValue.ToString());
            }
        }

        private void cbx_city_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.cbx_city.ValueMember != "")
            {
                GetAreaName(this.cbx_area, this.cbx_city.SelectedValue.ToString());
            }
        }

        public void SaveData(string strName, string strRow, string strCol, string strColor,
            string strArea, string strCityName, string strRemark, string strQuality,
            string strPrice, string strChartMan, string strTel, string strImageUrl,
            string strStyle, string strGrain, string strMoney, string strProvinceName)
        {
            string strSql = string.Format(@"INSERT INTO [dbo].[Goods]
           ([Name]
           ,[Row]
           ,[Col]
           ,[Color]
           ,[Areas]
           ,[CityName]
           ,[Remark]
           ,[Quality]
           ,[CreateTime]
           ,[CreatorId]
           ,[LastModifierId]
           ,[LastModifyTime]
           ,[Price]
           ,[ChartMan]
           ,[Tel]
           ,[ImageUrl]
           ,[Style]
           ,[Grain]
           ,[Money]
           ,[Stock]
           ,[ProvinceName]
           ,[IsEnable])
     VALUES
           ('{0}'
           ,'{1}'
           ,'{2}'
           ,'{3}'
           ,'{4}'
           ,'{5}'
           ,'{6}'
           ,'{7}'
           ,''
           ,1
           ,1
           ,'{8}'
           ,'{9}'
           ,'{10}'
           ,'{11}'
           ,'{12}'
           ,'{13}'
           ,'{14}'
           ,'{15}'
           ,1
           ,'{16}'
           ,1",strName,strRow,strCol,strColor,strArea,
           strCityName,strRemark,strQuality,strPrice,strChartMan,strTel,strImageUrl
           ,strStyle,strGrain,DateTime.Now,strMoney,strProvinceName);
            DbHelper helper = new DbHelper();
            helper.ExecuteNonQuery(strSql);
        }
    }

}
