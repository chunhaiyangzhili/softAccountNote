using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace SoftAccountNote
{
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern bool SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_MOVE = 0xF010;
        public const int HTCAPTION = 0x0002;

        private AccountNoteData accountNoteData = new AccountNoteData();
        private int MonBuget = 2000;
        // <summary>
        /// 用来存放DGV单元格修改之前值
        /// </summary>
        private object cellTempValue = null;

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {

            ReleaseCapture();
            SendMessage(this.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0);
        }

        private void PboxMin_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void PboxClose_Click(object sender, EventArgs e)
        {
            DialogResult ds =  MessageBox.Show("确认退出吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (ds == DialogResult.Yes)
            {
                accountNoteData.Close();
                Application.ExitThread();
            }
            else
            {

            }

        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if(this.WindowState == FormWindowState.Normal)
            {
                notifyIcon1.Visible = false;
            }else if(this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                notifyIcon1.Visible = true;
            }
        }

        private void NotifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if(this.WindowState == FormWindowState.Minimized)
            {
                this.Show();
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void 显示主界面ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NotifyIcon1_MouseDoubleClick(null, null);  //托盘图标快捷菜单
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("确定退出吗？", "退出", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if(result == DialogResult.OK)
            {
                accountNoteData.Close();
                Application.ExitThread();
            }
            else
            {
                this.WindowState = FormWindowState.Minimized;
            }
        }

        private void PboxMax_Click(object sender, EventArgs e)
        {
            if(this.WindowState != FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else if(this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Minimized;
            }
        }

        private void ToolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void ToolStrip1_Paint(object sender, PaintEventArgs e)
        {
            if ((sender as ToolStrip).RenderMode == ToolStripRenderMode.System)
            {
                Rectangle rect = new Rectangle(0, 0, this.toolStrip1.Width, this.toolStrip1.Height - 2);
                e.Graphics.SetClip(rect);
            }
        }

        private void TsBtnConsumePlay_MouseEnter(object sender, EventArgs e)
        {
            tsBtnConsumePlay.BackColor = Color.Green;
        }

        private void TsBtnConsumePlay_MouseLeave(object sender, EventArgs e)
        {
            tsBtnConsumePlay.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(157)))), ((int)(((byte)(215)))));
        }

        private void TsBtnNoteOne_MouseEnter(object sender, EventArgs e)
        {
            tsBtnNoteOne.BackColor = Color.Green;
        }

        private void TsBtnNoteOne_MouseLeave(object sender, EventArgs e)
        {
            tsBtnNoteOne.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(157)))), ((int)(((byte)(215)))));
        }

        private void TsBtnConsumePlay_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Columns["btnDGVDelete"] != null)
            {
                dataGridView1.Columns.Remove("btnDGVDelete");
            }
            tbcNoteOne.Visible = false;
            tbcAllConsume.Visible = false;
            tbcConfig.Visible = false;
            tbcSQL.Visible = false;
            tbcComsumePlay.Visible = true;
            tbcComsumePlay.Dock = DockStyle.Fill;
        }

        private bool SubmitInput()
        {
            int tmp;
            if (txtConsume.Text.Trim() == "")
            {
                MessageBox.Show("请输入消费金额（单位：元）", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtConsume.Focus();
                return false;
            }
            else if (!int.TryParse(txtConsume.Text.Trim(), out tmp))
            {
                MessageBox.Show("请输入正常的数字而非其他字符", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtConsume.Focus();
                return false;
            }
            else if(cmbxKind.Text.Trim() == "")
            {
                MessageBox.Show("请选择分类", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cmbxKind.Focus();
                return false;
            }else if(txtRemark.Text.Trim() == "")
            {
                MessageBox.Show("请输入备注", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtConsume.Focus();
                return false;
            }
            return true;
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            if (SubmitInput())
            {
                bool AddResult = accountNoteData.Add(txtConsume.Text.ToString(), cmbxKind.Text.ToString(), txtRemark.Text.ToString(), dateTimePicker1.Text.ToString());
                if (AddResult != true)
                {
                    MessageBox.Show("提交失败", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("提交成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            if (dataGridView1.Columns["btnDGVDelete"] != null)
            {
                dataGridView1.Columns.Remove("btnDGVDelete");
            }
            tbcNoteOne.Visible = false;
            tbcAllConsume.Visible = false;
            tbcConfig.Visible = true;
            tbcSQL.Visible = false;
            tbcComsumePlay.Visible = false;
            tbcConfig.Dock = DockStyle.Fill;
        }

        private void TsBtnNoteOne_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Columns["btnDGVDelete"] != null)
            {
                dataGridView1.Columns.Remove("btnDGVDelete");
            }
            tbcAllConsume.Visible = false;
            tbcConfig.Visible = false;
            tbcComsumePlay.Visible = false;
            tbcSQL.Visible = false;
            tbcNoteOne.Visible = true;
            tbcNoteOne.Dock = DockStyle.Fill;
        }

        private void TbcNoteOne_DrawItem(object sender, DrawItemEventArgs e)
        {
            //Graphics g = e.Graphics;
            //Rectangle r = this.tbcNoteOne.GetTabRect(e.Index);
            
            //if(e.Index == this.tbcNoteOne.SelectedIndex)
            //{
            //    Brush selected_color = new SolidBrush(Color.FromArgb(0, 157, 215)); ; //选中的项的背景色;
            //    g.FillRectangle(selected_color, r); //改变选项卡标签的背景色;
            //                                        //this.Font = new System.Drawing.Font("宋体", 10.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            //    string title = this.tbcNoteOne.TabPages[e.Index].Text;
            //    g.DrawString(title, new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134))), new SolidBrush(Color.Black), new PointF(r.X + 3, r.Y + 6));//PointF选项卡标题的位置;
            //}
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void BtnNoteOneResult_MouseMove(object sender, MouseEventArgs e)
        {

        }

        private void BtnNoteOneResult_MouseLeave(object sender, EventArgs e)
        {
            btnNoteOneResult.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(148)))), ((int)(((byte)(121)))), ((int)(((byte)(104)))));
        }

        private void BtnNoteOneResult_MouseEnter(object sender, EventArgs e)
        {
            btnNoteOneResult.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(157)))), ((int)(((byte)(215)))));
        }

        private void BtnChooseOther_Click(object sender, EventArgs e)
        {
            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }
            string Mon = dateTimePicker2.Text;
            bool flag = ChooseMonPaint(Mon);
            if (!flag)
            {
                MessageBox.Show("您选择的月份没有消费记录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private bool ChooseMonPaint(string Mon)
        {
            Dictionary<string, int> dicConsumeDay = new Dictionary<string, int>();
            Dictionary<string, int> dicConsumeKind = new Dictionary<string, int>();
            DataTable aa = accountNoteData.Query("MainData");
            foreach (DataRow item in aa.Rows)
            {
                if (Mon.Split(new char[] { '月' })[0] == item[4].ToString().Split(new char[] { '月'})[0])
                {
                    if(dicConsumeDay.ContainsKey(item[4].ToString().Split(new char[] { '月' })[1].Split(new char[] { '日' })[0]))
                    {
                        dicConsumeDay[item[4].ToString().Split(new char[] { '月' })[1].Split(new char[] { '日' })[0]] = int.Parse(item[1].ToString()) + dicConsumeDay[item[4].ToString().Split(new char[] { '月' })[1].Split(new char[] { '日' })[0]];
                    }
                    else
                    {
                        dicConsumeDay.Add(item[4].ToString().Split(new char[] { '月' })[1].Split(new char[] { '日' })[0], int.Parse(item[1].ToString()));
                    }
                }
            }
            
            ChartArea chartArea1 = chart1.ChartAreas["ChartArea1"];
            chartArea1.AxisX.Title = "日期";
            chartArea1.AxisY.Title = "花费：（元）";
            foreach (KeyValuePair<string, int> kv in dicConsumeDay)
            {
                chart1.Series[0].Points.AddXY(Convert.ToInt32(kv.Key), Convert.ToInt32(kv.Value));
            }
            chart1.Series[0].ToolTip = "#VAL(元)";
            chart1.Series[0].Label = "#VAL(元)";

            Series series2 = chart1.Series["Series2"];
            foreach (DataRow item in aa.Rows)
            {
                if (Mon.Split(new char[] { '月' })[0] == item[4].ToString().Split(new char[] { '月' })[0])
                {
                    if (dicConsumeKind.ContainsKey(item[2].ToString()))
                    {
                        dicConsumeKind[item[2].ToString()] = int.Parse(item[1].ToString()) + dicConsumeKind[item[2].ToString()];
                    }
                    else
                    {
                        dicConsumeKind.Add(item[2].ToString(), int.Parse(item[1].ToString()));
                    }

                }
            }
            chart1.Series[1]["PieLabelStyle"] = "Outside";//将文字移到外侧
            chart1.Series[1]["PieLineColor"] = "Black";//绘制黑色的连线。

            foreach (KeyValuePair<string, int> kv in dicConsumeKind)
            {
                series2.Points.AddXY(kv.Key, kv.Value);
            }
            chart1.Series[1].ToolTip = "#VAL(元)";
            if (dicConsumeDay.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void Chart1_Click(object sender, EventArgs e)
        {

        }

        private bool ChooseDayPaint(string day)
        {
            Dictionary<string, int> dicConsumeKind = new Dictionary<string, int>();
            DataTable aa = accountNoteData.Query("MainData");
            foreach (DataRow item in aa.Rows)
            {
                if (day.Split(new char[] { '月' })[1] == item[4].ToString().Split(new char[] { '月' })[1])
                {
                    if (dicConsumeKind.ContainsKey(item[2].ToString()))
                    {
                        dicConsumeKind[item[2].ToString()] = int.Parse(item[1].ToString()) + dicConsumeKind[item[2].ToString()];
                    }
                    else
                    {
                        dicConsumeKind.Add(item[2].ToString(), int.Parse(item[1].ToString()));
                    }

                }
            }
            if(dicConsumeKind.Count == 0)
            {
                return false;
            }
            else
            {
                ChartArea chartArea1 = chart1.ChartAreas["ChartArea1"];
                //chart1.Series[0].IsValueShownAsLabel = true;
                chartArea1.AxisX.Title = "消费类型";
                chartArea1.AxisY.Title = "花费：（元）";
                foreach (KeyValuePair<string, int> kv in dicConsumeKind)
                {
                    chart1.Series[0].Points.AddXY(kv.Key, Convert.ToInt32(kv.Value));
                }
                chart1.Series[0].ToolTip = "#VAL(元)";
                chart1.Series[0].Label = "#VAL(元)";

                Series series2 = chart1.Series["Series2"];
                chart1.Series[1]["PieLabelStyle"] = "Outside";//将文字移到外侧
                chart1.Series[1]["PieLineColor"] = "Black";//绘制黑色的连线。

                foreach (KeyValuePair<string, int> kv in dicConsumeKind)
                {
                    series2.Points.AddXY(kv.Key, kv.Value);
                }
                chart1.Series[1].ToolTip = "#VAL(元)";
                return true;
            }
        }
        private void BtnChooseDay_Click(object sender, EventArgs e)
        {
            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }
            string day = dateTimePicker2.Text;
            bool flag = ChooseDayPaint(day);
            if (!flag)
            {
                MessageBox.Show("您选择的日子没有消费记录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Label8_Click(object sender, EventArgs e)
        {

        }

        private void TsBtnConfig_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Columns["btnDGVDelete"] != null)
            {
                dataGridView1.Columns.Remove("btnDGVDelete");
            }
            tbcNoteOne.Visible = false;
            tbcComsumePlay.Visible = false;
            tbcAllConsume.Visible = false;
            tbcSQL.Visible = false;
            tbcConfig.Visible = true;
            tbcConfig.Dock = DockStyle.Fill;
        }

        private void BtnConfigResult_Click(object sender, EventArgs e)
        {
            if (txtBudget.Text.Trim() == "")
            {
                MessageBox.Show("您没有设置预算，默认为2000元", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MonBuget = int.Parse(txtBudget.Text.Trim());
                DateTime now = DateTime.Now;
                string MonDate = now.GetDateTimeFormats('y')[0].ToString();//2005年11月
                DataTable bb = accountNoteData.MonQuery("BaseData", MonDate);

                if (bb.Rows.Count != 0)
                {
                    bool ChangeResult = accountNoteData.ChangeBuget(MonDate, MonBuget.ToString());
                    if (ChangeResult)
                    {
                        MessageBox.Show("更改预算金额成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    bool AddResult = accountNoteData.AddBuget(MonDate, MonBuget.ToString());
                    if (AddResult)
                    {
                        MessageBox.Show("设置预算金额成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                Console.WriteLine("1111");
            }
        }

        private void TsBtnAllConsume_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Columns["btnDGVDelete"] != null)
            {
                dataGridView1.Columns.Remove("btnDGVDelete");
            }
            tbcNoteOne.Visible = false;
            tbcComsumePlay.Visible = false;
            tbcSQL.Visible = false;
            tbcAllConsume.Visible = true;
            tbcConfig.Visible = false;
            tbcAllConsume.Dock = DockStyle.Fill;
            Dictionary<string, int> dicConsumeDay = new Dictionary<string, int>();
            Dictionary<string, int> dicConsumeKind = new Dictionary<string, int>();
            DataTable aa = accountNoteData.Query("MainData");
            //accountNoteData.Close();
            int momAllConsume = 0;
            int dayAllConsume = 0;
            DateTime now = DateTime.Now;
            string MonDate = now.GetDateTimeFormats('y')[0].ToString();//2005年11月
            int days = DateTime.DaysInMonth(int.Parse(now.Year.ToString()), int.Parse(now.Month.ToString()));
            days -= now.Day; 
            foreach (DataRow item in aa.Rows)
            {
                if (now.ToLongDateString().ToString().Split(new char[] { '月' })[0] == item[4].ToString().Split(new char[] { '月' })[0])
                {
                    momAllConsume += int.Parse(item[1].ToString()); 
                }
                if (now.ToLongDateString().ToString() == item[4].ToString())
                {
                    dayAllConsume += int.Parse(item[1].ToString());
                }
            }
            lblMonConsume.Text = momAllConsume.ToString();
            lblDayConsume.Text = dayAllConsume.ToString();
            lblEndMon.Text = days.ToString();
            foreach (var series in chart2.Series)
            {
                series.Points.Clear();
            }

            
            DataTable bb = accountNoteData.Query("BaseData");
            foreach (DataRow item in bb.Rows)
            {
                if (MonDate == item[1].ToString())
                {
                    MonBuget = int.Parse((item[2].ToString()));
                }
            }
            string[] x = {"已消费额度", "剩余额度" };
            int[] y = { momAllConsume, MonBuget-momAllConsume};
            chart2.Series[0].XValueType = ChartValueType.String;  //设置X轴上的值类型
            //chart2.Series[0].Label = "#PERCENT";
            chart2.Series[0].Label = "#PERCENT";
            chart2.Series[0].LegendText = "#VALX";
            chart2.Series[0].Font = new Font("微软雅黑", 10f, FontStyle.Regular);
            //chart2.Series[0].IsValueShownAsLabel = true;

            chart2.Series[0].CustomProperties = "DrawingStyle = Cylinder";
            //chart2.Series[0].CustomProperties = "PieLabelStyle = Outside";
            chart2.Legends[0].Position.Auto = true;
            chart2.Series[0].IsValueShownAsLabel = true;
            chart2.Series[0].IsVisibleInLegend = true;

            //绑定数据
            chart2.Series[0].Points.DataBindXY(x, y);

        }

        private void TsBtnConfig_MouseEnter(object sender, EventArgs e)
        {
            tsBtnConfig.BackColor = Color.Green;
        }

        private void TsBtnConfig_MouseLeave(object sender, EventArgs e)
        {
            tsBtnConfig.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(157)))), ((int)(((byte)(215)))));
        }

        private void TsBtnAllConsume_MouseEnter(object sender, EventArgs e)
        {
            tsBtnAllConsume.BackColor = Color.Green;
        }

        private void TsBtnAllConsume_MouseLeave(object sender, EventArgs e)
        {
            tsBtnAllConsume.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(157)))), ((int)(((byte)(215)))));
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            bool cc = accountNoteData.Del();
            DataTable dd = accountNoteData.Query("MainData");
            string sqlStr = "Alter TABLE MainData Alter COLUMN ID COUNTER (1, 1)";
            bool ee = accountNoteData.HandleSQL(sqlStr);
            if (cc || dd.Rows.Count==0)
            {
                MessageBox.Show("初始化用户数据成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnChooseDay_MouseEnter(object sender, EventArgs e)
        {
            btnChooseDay.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(157)))), ((int)(((byte)(215)))));
        }

        private void BtnChooseDay_MouseLeave(object sender, EventArgs e)
        {
            btnChooseDay.BackColor = Color.White;
        }

        private void BtnChooseOther_MouseEnter(object sender, EventArgs e)
        {
            btnChooseOther.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(157)))), ((int)(((byte)(215)))));
        }

        private void BtnChooseOther_MouseLeave(object sender, EventArgs e)
        {
            btnChooseOther.BackColor = Color.White;
        }

        private void BtnReset_MouseEnter(object sender, EventArgs e)
        {
            btnReset.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(157)))), ((int)(((byte)(215)))));
        }

        private void BtnReset_MouseLeave(object sender, EventArgs e)
        {
            btnReset.BackColor = Color.White;
        }

        private void BtnConfigResult_MouseEnter(object sender, EventArgs e)
        {
            btnConfigResult.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(157)))), ((int)(((byte)(215)))));
        }

        private void BtnConfigResult_MouseLeave(object sender, EventArgs e)
        {
            btnConfigResult.BackColor = Color.White;
        }

        private void RichTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void tsBtnSQL_Click(object sender, EventArgs e)
        {
            tbcNoteOne.Visible = false;
            tbcComsumePlay.Visible = false;
            tbcAllConsume.Visible = false;
            tbcConfig.Visible = false;
            tbcSQL.Visible = true;
            tbcSQL.Dock = DockStyle.Fill;
            dataGridView1.Dock = DockStyle.Fill;
            DataTable aa = accountNoteData.Query("MainData");
            dataGridView1.DataSource = aa;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.RowsDefaultCellStyle.WrapMode = (DataGridViewTriState.True);
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            DataGridViewButtonColumn btnDGVDelete = new DataGridViewButtonColumn();
            btnDGVDelete.Name = "btnDGVDelete";
            btnDGVDelete.HeaderText = "删除";
            btnDGVDelete.DefaultCellStyle.NullValue = "删除";
            dataGridView1.Columns.Insert(5, btnDGVDelete);
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(148, 121, 104);
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(148, 121, 104);
            dataGridView1.DefaultCellStyle.BackColor = Color.FromArgb(148, 121, 104);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int CIndex = this.dataGridView1.CurrentCell.ColumnIndex;
            if (CIndex == 5)
            {
                //说明点击的列是DataGridViewButtonColumn列
                string sqlStr = "delete from MainData where ID=" + Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value) + "";
                //执行指定的SQL命令语句,并返回命令所影响的行数
                bool aa = accountNoteData.HandleSQL(sqlStr);
                if (aa == true) MessageBox.Show("删除成功");
                dataGridView1.Columns.Remove("btnDGVDelete");
                DataTable bb = accountNoteData.Query("MainData");
                dataGridView1.DataSource = bb;
                DataGridViewButtonColumn btnDGVDelete = new DataGridViewButtonColumn();
                btnDGVDelete.Name = "btnDGVDelete";
                btnDGVDelete.HeaderText = "删除";
                btnDGVDelete.DefaultCellStyle.NullValue = "删除";
                btnDGVDelete.DefaultCellStyle.BackColor = Color.FromArgb(148, 121, 104);
                dataGridView1.Columns.Insert(5, btnDGVDelete);
            }
        }



        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            cellTempValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //判断编辑前后的值是否一样（是否修改了内容）
            if (Object.Equals(cellTempValue, dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
            {
                //如果没有修改，则返回
                return;
            }
            else
            {
                string sqlStr = string.Empty;
                //说明点击的列是DataGridViewButtonColumn列
                sqlStr = "update MainData set 花费='" + dataGridView1.CurrentRow.Cells[1].Value
                + "',分类='" + dataGridView1.CurrentRow.Cells[2].Value
                + "',备注='" + dataGridView1.CurrentRow.Cells[3].Value
                + "',日期='" + dataGridView1.CurrentRow.Cells[4].Value
                + "'where ID=" + Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value) + "";
                ////执行指定的SQL命令语句,并返回命令所影响的行数
                try 
                {
                    accountNoteData.HandleSQL(sqlStr);
                }
                catch(OleDbException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
