
after changing other repo
namespace PBSepartor
{
    public partial class Form1 : Form
    {
        #region VAR
        private string dbName;
        private string excellName;
        private string connectionString;
       
        private string st17;
        private string st18;
        private string st19;
        private string st20;
        pr
        private bool IsfileOpen;
        private string strSelectCommand;
        private Form1.stateProgram stProgramm;
        private static int m_dataRowCount;

        private enum stateProgram
        {
            openFile,
            CloseFile,
        }

        #endregion

        public Form1()
        {
            this.components = (IContainer)null;
            this.con = new OleDbConnection();
            this.da = new OleDbDataAdapter();
            this.dtSensor = new System.Data.DataTable();
            this.tblData = new System.Data.DataTable();
            this.strOrder = " order by id";
            this.stProgramm = Form1.stateProgram.openFile;
            //bas:ctor
            this.InitializeComponent();
        }

        public string GetCPUId()
        {
            string str = string.Empty;
            //foreach (ManagementObject managementObject in new ManagementClass("Win32_Processor").GetInstances())
            //{
            //    if (str == string.Empty)
            //        str = managementObject.Properties["ProcessorId"].Value.ToString();
            //}
            str = "BFEBFBFF000206A7";
            return str;
        }

shrterkljkl
rtkl;jkjhwerhwejklrjhkwer
        {
            string str = string.Empty;
            foreach (ManagementObject managementObject in new ManagementClass("Win32_Processor").GetInstances())
            {
                if (str == string.Empty)
                    str = managementObject.Properties["ProcessorId"].Value.ToString();
            }

            if (str == "BFEBFBFF000306C3")
                return true;
            else
                return false;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.setInitial();
        }

        private void setInitial()
        {
            this.IsfileOpen = false;
            this.radPanel1.Dock = DockStyle.Top;
            this.rdbGrid1.Left = 10;
            this.rdbGrid1.Width = this.radDock1.Width - 30;
            this.rdbGrid1.Height = this.radDock1.Height - 230;
            this.dtPlayBack.Format = DateTimePickerFormat.Custom;
            this.dtPlayBack.ShowUpDown = true;
            this.dtPlayBack.CustomFormat = "yyyy/MM/dd HH:mm:ss";
            this.dtPlayBack.Value = DateTime.Now;
            this.dtFrom.Format = DateTimePickerFormat.Custom;
            this.dtFrom.ShowUpDown = true;
            this.dtFrom.CustomFormat = "yyyy/MM/dd HH:mm:ss";
            this.dtFrom.Value = DateTime.Now;
            this.dtTo.Format = DateTimePickerFormat.Custom;
            this.dtTo.ShowUpDown = true;
            this.dtTo.CustomFormat = "yyyy/MM/dd HH:mm:ss";
            this.dtTo.Value = DateTime.Now;
            this.queryCounter = 0;
            this.rdbGrid1.MasterGridViewTemplate.AllowAddNewRow = false;
            this.rdbGrid1.MasterGridViewTemplate.AllowDeleteRow = false;
            this.rdbGrid1.MasterGridViewTemplate.AutoSizeColumnsMode = GridViewAutoSizeColumnsMode.Fill;
            this.rdbGrid1.MasterGridViewTemplate.ShowRowHeaderColumn = false;
            this.rdbGrid1.MasterGridViewTemplate.ShowFilteringRow = false;
            this.rdbGrid1.EnableSorting = false;
            this.rdbGrid1.EnableGrouping = false;
           
            this.settingChartTemp();
            this.settingChartPres();
            this.settingChartLev();
            this.settingChartFlw();
            this.settingChartFlowRate();
            this.settingChartcustom1();
            this.settingChartcustom2();

            this.loadDefaultView();
            toolWindowGrid.BringToFront();
            this.stProgramm = Form1.stateProgram.CloseFile;
            this.btnShow();
        }

        private void loadDefaultView()
        {
            try
            {
                this.radDock1.LoadFromXml((System.Windows.Forms.Application.StartupPath).ToString() + "\\pdefaultView.xml");
                toolWindowGrid.BringToFront();
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(("Error in loadDefaultView()  " + ex.Message));
                int num = (int)MessageBox.Show("Not found pdefaultView.XML file\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void btnShow()
        {
            try
            {
                if (this.stProgramm == Form1.stateProgram.openFile)
                {
                    this.mOpen.Enabled = false;
                    this.mClose.Enabled = true;
                    this.mExport.Enabled = true;
                    this.btnDefaultCaption.Enabled = true;
                    this.btnChangeCaption.Enabled = true;
                    this.btnRunQuery.Enabled = true;
                    this.btnBindToChart.Enabled = true;
                    this.btnExport.Enabled = true;
                    this.btnSave.Enabled = true;
                    this.btnRefresh.Enabled = true;
                }
                else
                {
                    this.mOpen.Enabled = true;
                    this.mClose.Enabled = false;
                    this.mExport.Enabled = false;
                    this.btnDefaultCaption.Enabled = false;
                    this.btnChangeCaption.Enabled = false;
                    this.btnRunQuery.Enabled = false;
                    this.btnBindToChart.Enabled = false;
                    this.btnSave.Enabled = false;
                    this.btnExport.Enabled = false;
                    this.btnRefresh.Enabled = false;
                    if (this.con.State == ConnectionState.Open)
                    {
                        this.con.Close();
                        this.rdbGrid1.DataSource = new DataSet();
                    }
                }
            }
            catch (Exception ex)
            {
                lstReport.Items.Add("Error in map: " + ex.Message);
            }

        }

        private void mapToNumeric(int digit)
        {
            try
            {
                this.numHH.Value = (Decimal)(digit / 3600);
                this.numMM.Value = (Decimal)(digit % 3600 / 60);
                this.numSS.Value = (Decimal)(digit % 3600 % 60);
            }
            catch (Exception ex)
            {
                lstReport.Items.Add("Error in map: " + ex.Message);
            }
        }

        private bool correctRate()
        {
            try
            {
                int num1 = (int)(this.numHH.Value * new Decimal(3600) + this.numMM.Value * new Decimal(60) + this.numSS.Value);
                if (num1 == 0 || num1 < this.sampleRate)
                {
                    this.lstReport.Items.Add("Invalid data.\nPlease check and try again");
                    int num2 = (int)MessageBox.Show("Invalid data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    this.mapToNumeric(this.sampleRate);
                    return false;
                }
                else
                {
                    int num2 = this.sampleRate / this.sampleRate % this.sampleRate;
                    return true;
                }
            }
            catch (Exception ex)
            {
                lstReport.Items.Add(ex.Message);
                return false;
            }
        }

        private void settingChartTemp()
        {
            this.setChart(this.chartTemp);
            this.addSeries(this.chartTemp, "s1", this.btnS1.BackColor);
            this.addSeries(this.chartTemp, "s2", this.btnS2.BackColor);
            this.addSeries(this.chartTemp, "s3", this.btnS3.BackColor);
            this.addSeries(this.chartTemp, "s4", this.btnS4.BackColor);
            this.addSeries(this.chartTemp, "s5", this.btnS5.BackColor);
            this.addSeries(this.chartTemp, "s6", this.btnS6.BackColor);
        }

        private void settingChartPres()
        {
            this.setChart(this.chartPres);
            this.addSeries(this.chartPres, "s7", this.btnS7.BackColor);
            this.addSeries(this.chartPres, "s8", this.btnS8.BackColor);
            this.addSeries(this.chartPres, "s9", this.btnS9.BackColor);
        }

        private void settingChartLev()
        {
            this.setChart(this.chartLev);
            this.addSeries(this.chartLev, "s10", this.btnS10.BackColor);
            this.addSeries(this.chartLev, "s11", this.btnS11.BackColor);
        }

        private void settingChartFlw()
        {
            this.setChart(this.chartFlw);
            this.addSeries(this.chartFlw, "s12", this.btnS12.BackColor);
            this.addSeries(this.chartFlw, "s13", this.btnS13.BackColor);
            this.addSeries(this.chartFlw, "s14", this.btnS14.BackColor);
            this.addSeries(this.chartFlw, "s15", this.btnS15.BackColor);
            this.addSeries(this.chartFlw, "s16", this.btnS16.BackColor);
            this.addSeries(this.chartFlw, "p2", this.btnP2.BackColor);
        }

        private void settingChartFlowRate()
        {
            this.setChart(this.ChartFlowRate);
            this.addSeries(this.ChartFlowRate, "p1", this.btnP1.BackColor);
        }

        private void settingChartcustom1()
        {
            this.setChart(this.chartCustom1);
            this.addSeries(this.chartCustom1, "s17", this.btnS17.BackColor);
            this.addSeries(this.chartCustom1, "s18", this.btnS18.BackColor);
            this.addSeries(this.chartCustom1, "s19", this.btnS19.BackColor);
            this.addSeries(this.chartCustom1, "s20", this.btnS20.BackColor);
        }

        private void settingChartcustom2()
        {
            this.setChart(this.chartCustom2);
            this.addSeries(this.chartCustom2, "s21", this.btnS21.BackColor);
            this.addSeries(this.chartCustom2, "s22", this.btnS22.BackColor);
            this.addSeries(this.chartCustom2, "s23", this.btnS23.BackColor);
            this.addSeries(this.chartCustom2, "s24", this.btnS24.BackColor);
        }


        private void setChart(Dundas.Charting.WinControl.Chart myChart)
        {
            try
            {

                myChart.ChartAreas[0].AxisY.LabelStyle.Enabled = true;
                myChart.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount;
                myChart.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                myChart.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.FromArgb(110, 110, 110);
                myChart.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;
                myChart.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;

                myChart.BackColor = Color.Transparent;
                myChart.ChartAreas[0].BackColor = Color.White;
                myChart.ChartAreas[0].Position.Auto = true;
                myChart.ChartAreas[0].InnerPlotPosition.Auto = true;
                myChart.ChartAreas[0].AxisX.LabelStyle.Enabled = true;
                myChart.ChartAreas[0].AxisX.LabelStyle.ShowEndLabels = true;
                myChart.ChartAreas[0].AxisX.LabelStyle.FontColor = Color.Black;
                myChart.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
                myChart.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Seconds;
                myChart.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss\nyyyy/MM/dd";
                myChart.ChartAreas[0].AxisX.LabelsAutoFitStyle = LabelsAutoFitStyle.IncreaseFont | LabelsAutoFitStyle.DecreaseFont | LabelsAutoFitStyle.WordWrap;
                myChart.ChartAreas[0].AxisX.Minimum = this.minValue.ToOADate();
                myChart.ChartAreas[0].AxisX.Maximum = this.maxValue.ToOADate();
                myChart.ChartAreas[0].AxisX.MajorTickMark.Enabled = true;
                myChart.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;
                myChart.ChartAreas[0].AxisX.MajorTickMark.LineColor = Color.Black;
                myChart.ChartAreas[0].AxisX.MinorTickMark.LineColor = Color.Black;
                myChart.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                myChart.ChartAreas[0].AxisX.MajorGrid.IntervalType = DateTimeIntervalType.Seconds;
                myChart.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.FromArgb(110, 110, 110);
                myChart.Legends[0].Enabled = false;
                myChart.Legends[0].LegendStyle = LegendStyle.Row;
                myChart.Legends[0].Alignment = StringAlignment.Center;
                myChart.Legends[0].Docking = LegendDocking.Top;
                myChart.Legends[0].InsideChartArea = "Default";
                myChart.ChartAreas[0].CursorX.UserEnabled = true;
                myChart.ChartAreas[0].CursorX.UserSelection = true;
                myChart.ChartAreas[0].CursorX.IntervalType = DateTimeIntervalType.Seconds;
                myChart.ChartAreas[0].AxisX.View.Zoomable = true;
                myChart.ChartAreas[0].AxisY.View.Zoomable = true;
                myChart.ChartAreas[0].CursorX.AutoScroll = true;
                myChart.ChartAreas[0].CursorY.AutoScroll = true;
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(myChart.Name + "  " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void addSeries(Dundas.Charting.WinControl.Chart chartName, string seriesName, Color seriesColor)
        {
            try
            {
                chartName.Series.Add(seriesName);
                chartName.Series[seriesName].Type = SeriesChartType.FastLine;
                chartName.Series[seriesName].Color = seriesColor;
                chartName.Series[seriesName].BorderWidth = 2;
                chartName.Series[seriesName].BorderStyle = ChartDashStyle.Solid;
                chartName.Series[seriesName].XValueType = ChartValueTypes.DateTime;
            }
            catch (Exception ex)
            {
                lstReport.Items.Add(ex.Message);
            }
        }

        private void btnLeft_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.stProgramm == Form1.stateProgram.openFile)
                {
                    int num = (int)numStepChart.Value;
                    this.dtPlayBack.Value = dtPlayBack.Value.AddMinutes((double)-num);
                    DateTime dateTime = dtPlayBack.Value.AddMinutes((double)-num);
                    this.minValue = dateTime;
                    this.maxValue = dateTime.AddMinutes((double)num);
                    this.dtPlayBack.Value = dateTime.AddMinutes((double)num);
                    this.fillGird(createSelectCommand() + " where (LogDate >= #" + dateTime + "# AND LogDate <= #" + dateTime.AddMinutes((double)num) + "#) order by ID");
                    this.bindToChart();
                    this.assignCaption();
                }
                else
                {
                    int num1 = (int)MessageBox.Show("Please Open File", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
            catch (Exception ex)
            {
                lstReport.Items.Add("Error in leftClick " + ex.Message);
            }
        }

        private void btnRight_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.stProgramm == Form1.stateProgram.openFile)
                {
                    DateTime dateTime = this.dtPlayBack.Value;
                    int num = (int)this.numStepChart.Value;
                    this.minValue = dateTime;
                    this.maxValue = dateTime.AddMinutes((double)num);
                    this.dtPlayBack.Value = dateTime.AddMinutes((double)num);
                    this.fillGird(this.createSelectCommand() + " where (LogDate >= #" + dateTime + "# AND LogDate <= #" + dateTime.AddMinutes((double)num) + "#) order by ID");
                    this.bindToChart();
                    this.assignCaption();
                }
                else
                {
                    int num1 = (int)MessageBox.Show("Please Open File", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
            catch (Exception ex)
            {
                lstReport.Items.Add("Error in RightClick " + ex.Message);
            }
        }

        private void btnRunQuery_Click(object sender, EventArgs e)
        {
            this.runQuery();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                if (!this.IsfileOpen)
                    return;

                this.con.Close();
                this.con.ConnectionString = this.strSelectCommand;
                if (this.con.State == ConnectionState.Closed)
                    this.con.Open();
                int lastID = int.Parse(new OleDbCommand("SELECT MAX(id) FROM tblData", this.con).ExecuteScalar().ToString());
                if (this.con.State == ConnectionState.Closed)
                    this.con.Open();
                this.dtTo.Value = DateTime.Parse(new OleDbCommand("SELECT logDate FROM tblData where id = " + lastID, this.con).ExecuteScalar().ToString());
                this.runQuery();
            }
            catch (Exception ex)
            {
                lstReport.Items.Add("Error Refresh " + ex.Message);
            }
        }

        private void btnBindToChart_Click(object sender, EventArgs e)
        {
            this.bindToChart();
        }

        private void btnChangeCaption_Click(object sender, EventArgs e)
        {
            this.assignCaption();
        }

        private void btnDefaultCaption_Click(object sender, EventArgs e)
        {
            this.txtS1.Text = this.dtSensor.Rows[0]["captionD"].ToString();
            this.txtS2.Text = this.dtSensor.Rows[1]["captionD"].ToString();
            this.txtS3.Text = this.dtSensor.Rows[2]["captionD"].ToString();
            this.txtS4.Text = this.dtSensor.Rows[3]["captionD"].ToString();
            this.txtS5.Text = this.dtSensor.Rows[4]["captionD"].ToString();
            this.txtS6.Text = this.dtSensor.Rows[5]["captionD"].ToString();
            this.txtS7.Text = this.dtSensor.Rows[6]["captionD"].ToString();
            this.txtS8.Text = this.dtSensor.Rows[7]["captionD"].ToString();
            this.txtS9.Text = this.dtSensor.Rows[8]["captionD"].ToString();
            this.txtS10.Text = this.dtSensor.Rows[9]["captionD"].ToString();
            this.txtS11.Text = this.dtSensor.Rows[10]["captionD"].ToString();
            this.txtS12.Text = this.dtSensor.Rows[11]["captionD"].ToString();
            this.txtS13.Text = this.dtSensor.Rows[12]["captionD"].ToString();
            this.txtS14.Text = this.dtSensor.Rows[13]["captionD"].ToString();
            this.txtS15.Text = this.dtSensor.Rows[14]["captionD"].ToString();
            this.txtS16.Text = this.dtSensor.Rows[15]["captionD"].ToString();
            this.txtS17.Text = this.dtSensor.Rows[16]["captionD"].ToString();
            this.txtS18.Text = this.dtSensor.Rows[17]["captionD"].ToString();
            this.txtS19.Text = this.dtSensor.Rows[18]["captionD"].ToString();
            this.txtS20.Text = this.dtSensor.Rows[19]["captionD"].ToString();
            this.txtS21.Text = this.dtSensor.Rows[20]["captionD"].ToString();
            this.txtS22.Text = this.dtSensor.Rows[21]["captionD"].ToString();
            this.txtS23.Text = this.dtSensor.Rows[22]["captionD"].ToString();
            this.txtS24.Text = this.dtSensor.Rows[23]["captionD"].ToString();
            this.txtParam1.Text = this.dtSensor.Rows[24]["captionD"].ToString();
            this.txtParam2.Text = this.dtSensor.Rows[25]["captionD"].ToString();
            this.txtParam3.Text = this.dtSensor.Rows[26]["captionD"].ToString();
            this.txtParam4.Text = this.dtSensor.Rows[27]["captionD"].ToString();

            this.assignCaption();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                this.rdbGrid1.CurrentRow.Cells["comment"].Value = this.txtComment.Text;
                OleDbCommand oleDbCommand = new OleDbCommand();
                if (this.con.State == ConnectionState.Closed)
                    this.con.Open();
                oleDbCommand.Connection = this.con;
                oleDbCommand.CommandText = "UPDATE tblData set comment=@comment where id=@id";
                oleDbCommand.Parameters.Clear();
                oleDbCommand.Parameters.AddWithValue("@comment", this.txtComment.Text);
                oleDbCommand.Parameters.AddWithValue("@id", this.txtRow.Text);
                oleDbCommand.ExecuteNonQuery();
                int num = (int)MessageBox.Show("Update successfully ", "Comment", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add((ex.Message + " at " + DateTime.Now));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Are you sure to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.No)
                return;
            e.Cancel = true;
        }



        private void chartPres_CursorPositionChanged(object sender, CursorEventArgs e)
        {
            try
            {
                if (double.IsNaN(e.NewPosition))
                    return;
                int count = this.chartPres.Series[0].Points.Count;
                do
                {
                    --count;
                }
                while (this.chartPres.Series[0].Points[count].XValue - e.NewPosition > 0.0);
                SetCursorPositiion(count);
                this.showLableData(count);
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
            }
        }

        private void chartLev_CursorPositionChanged(object sender, CursorEventArgs e)
        {
            try
            {
                if (double.IsNaN(e.NewPosition))
                    return;
                int count = this.chartLev.Series[0].Points.Count;
                do
                {
                    --count;
                }
                while (this.chartLev.Series[0].Points[count].XValue - e.NewPosition > 0.0);
                SetCursorPositiion(count);
                this.showLableData(count);
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
            }
        }

        private void chartFlw_CursorPositionChanged(object sender, CursorEventArgs e)
        {
            try
            {
                if (double.IsNaN(e.NewPosition))
                    return;
                int count = this.chartFlw.Series[0].Points.Count;
                do
                {
                    --count;
                }
                while (this.chartFlw.Series[0].Points[count].XValue - e.NewPosition > 0.0);
                SetCursorPositiion(count);
                this.showLableData(count);
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
            }
        }

        private void chartTemp_CursorPositionChanged(object sender, CursorEventArgs e)
        {
            try
            {
                if (double.IsNaN(e.NewPosition))
                    return;
                int count = this.chartTemp.Series[0].Points.Count;
                do
                {
                    --count;
                }
                while (this.chartTemp.Series[0].Points[count].XValue - e.NewPosition > 0.0);
                SetCursorPositiion(count);
                this.showLableData(count);
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
            }
        }

        private void chartCustom1_CursorPositionChanged(object sender, CursorEventArgs e)
        {
            try
            {
                if (double.IsNaN(e.NewPosition))
                    return;
                int count = this.chartCustom1.Series[0].Points.Count;
                do
                {
                    --count;
                }
                while (this.chartCustom1.Series[0].Points[count].XValue - e.NewPosition > 0.0);
                SetCursorPositiion(count);
                this.showLableData(count);
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
            }
        }

        private void chartCustom2_CursorPositionChanged(object sender, CursorEventArgs e)
        {
            try
            {
                if (double.IsNaN(e.NewPosition))
                    return;
                int count = this.chartCustom2.Series[0].Points.Count;
                do
                {
                    --count;
                }
                while (this.chartCustom2.Series[0].Points[count].XValue - e.NewPosition > 0.0);
                SetCursorPositiion(count);
                this.showLableData(count);
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
            }
        }

        private void SetCursorPositiion(int count)
        {
            if (chartTemp.Series.Count > 0)
                if (chartTemp.Series[0].Points.Count >= count)
                    this.chartTemp.ChartAreas[0].CursorX.Position = this.chartTemp.Series[0].Points[count].XValue;

            if (chartPres.Series.Count > 0)
                if (chartPres.Series[0].Points.Count >= count)
                    this.chartPres.ChartAreas[0].CursorX.Position = this.chartPres.Series[0].Points[count].XValue;

            if (chartLev.Series.Count > 0)
                if (chartLev.Series[0].Points.Count >= count)
                    this.chartLev.ChartAreas[0].CursorX.Position = this.chartLev.Series[0].Points[count].XValue;

            if (chartFlw.Series.Count > 0)
                if (chartFlw.Series[0].Points.Count >= count)
                    this.chartFlw.ChartAreas[0].CursorX.Position = this.chartFlw.Series[0].Points[count].XValue;

            if (chartCustom1.Series.Count > 0)
                if (chartCustom1.Series[0].Points.Count >= count)
                    this.chartCustom1.ChartAreas[0].CursorX.Position = this.chartCustom1.Series[0].Points[count].XValue;

            if (chartCustom2.Series.Count > 0)
                if (chartCustom2.Series[0].Points.Count >= count)
                    this.chartCustom2.ChartAreas[0].CursorX.Position = this.chartCustom2.Series[0].Points[count].XValue;
        }


        private void showLableData(int point)
        {
            try
            {
                this.lbS1.Text = this.chartTemp.Series[0].Points[point].YValues[0].ToString("0.00");
                this.lbS2.Text = this.chartTemp.Series[1].Points[point].YValues[0].ToString("0.00");
                this.lbS3.Text = this.chartTemp.Series[2].Points[point].YValues[0].ToString("0.00");
                this.lbS4.Text = this.chartTemp.Series[3].Points[point].YValues[0].ToString("0.00");
                this.lbS5.Text = this.chartTemp.Series[4].Points[point].YValues[0].ToString("0.00");
                this.lbS6.Text = this.chartTemp.Series[5].Points[point].YValues[0].ToString("0.00");

                this.lbS7.Text = this.chartPres.Series[0].Points[point].YValues[0].ToString("0.00");
                this.lbS8.Text = this.chartPres.Series[1].Points[point].YValues[0].ToString("0.00");
                this.lbS9.Text = this.chartPres.Series[2].Points[point].YValues[0].ToString("0.00");

                this.lbS10.Text = this.chartLev.Series[0].Points[point].YValues[0].ToString("0.00");
                this.lbS11.Text = this.chartLev.Series[1].Points[point].YValues[0].ToString("0.00");

                this.lbS12.Text = this.chartFlw.Series[0].Points[point].YValues[0].ToString("0.00");
                this.lbS13.Text = this.chartFlw.Series[1].Points[point].YValues[0].ToString("0.00");
                this.lbS14.Text = this.chartFlw.Series[2].Points[point].YValues[0].ToString("0.00");
                this.lbS15.Text = this.chartFlw.Series[3].Points[point].YValues[0].ToString("0.00");
                this.lbS16.Text = this.chartFlw.Series[4].Points[point].YValues[0].ToString("0.00");
                this.lbVP2.Text = this.chartFlw.Series[5].Points[point].YValues[0].ToString("0.00"); //GOR

                this.lbVP1.Text = this.ChartFlowRate.Series[0].Points[point].YValues[0].ToString("0.00"); //flowrate


                this.lbS17.Text = this.chartCustom1.Series[0].Points[point].YValues[0].ToString("0.00");
                this.lbS18.Text = this.chartCustom1.Series[1].Points[point].YValues[0].ToString("0.00");
                this.lbS19.Text = this.chartCustom1.Series[2].Points[point].YValues[0].ToString("0.00");
                this.lbS20.Text = this.chartCustom1.Series[3].Points[point].YValues[0].ToString("0.00");

                this.lbS21.Text = this.chartCustom2.Series[0].Points[point].YValues[0].ToString("0.00");
                this.lbS22.Text = this.chartCustom2.Series[1].Points[point].YValues[0].ToString("0.00");
                this.lbS23.Text = this.chartCustom2.Series[2].Points[point].YValues[0].ToString("0.00");
                this.lbS24.Text = this.chartCustom2.Series[3].Points[point].YValues[0].ToString("0.00");


            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
            }
            finally
            {
                if (chartTemp.Series.Count > 0)
                    if (chartTemp.Series[0].Points.Count >= point)
                        this.lblDateTime.Text = DateTime.FromOADate(double.Parse(this.chartTemp.Series[0].Points[point].XValue.ToString())).ToString();
                    else
                    {
                        this.lblDateTime.Text = "unknown";
                    }
                else
                {
                    this.lblDateTime.Text = "unknown";
                }
            }
        }

        private void bindToChart()
        {
            try
            {
                this.minValue = DateTime.Parse(this.rdbGrid1.Rows[0].Cells["logDate"].Value.ToString());
                this.maxValue = DateTime.Parse(this.rdbGrid1.Rows[this.rdbGrid1.Rows.Count - 1].Cells["logDate"].Value.ToString());

                this.chartTemp.ChartAreas[0].AxisX.Minimum = this.minValue.ToOADate();
                this.chartTemp.ChartAreas[0].AxisX.Maximum = this.maxValue.ToOADate();

                this.chartPres.ChartAreas[0].AxisX.Minimum = this.minValue.ToOADate();
                this.chartPres.ChartAreas[0].AxisX.Maximum = this.maxValue.ToOADate();

                this.chartLev.ChartAreas[0].AxisX.Minimum = this.minValue.ToOADate();
                this.chartLev.ChartAreas[0].AxisX.Maximum = this.maxValue.ToOADate();

                this.chartFlw.ChartAreas[0].AxisX.Minimum = this.minValue.ToOADate();
                this.chartFlw.ChartAreas[0].AxisX.Maximum = this.maxValue.ToOADate();

                this.ChartFlowRate.ChartAreas[0].AxisX.Minimum = this.minValue.ToOADate();
                this.ChartFlowRate.ChartAreas[0].AxisX.Maximum = this.maxValue.ToOADate();

                this.chartCustom1.ChartAreas[0].AxisX.Minimum = this.minValue.ToOADate();
                this.chartCustom1.ChartAreas[0].AxisX.Maximum = this.maxValue.ToOADate();

                this.chartCustom2.ChartAreas[0].AxisX.Minimum = this.minValue.ToOADate();
                this.chartCustom2.ChartAreas[0].AxisX.Maximum = this.maxValue.ToOADate();

                this.chartTemp.DataSource = this.tblData;
                this.chartPres.DataSource = this.tblData;
                this.chartLev.DataSource = this.tblData;
                this.chartFlw.DataSource = this.tblData;
                this.ChartFlowRate.DataSource = this.tblData;
                this.chartCustom1.DataSource = this.tblData;
                this.chartCustom2.DataSource = this.tblData;

                if (this.chkS1.Checked)
                {
                    this.chartTemp.Series[0].ValueMemberX = "LogDate";
                    this.chartTemp.Series[0].ValueMembersY = "s1";
                }
                if (this.chkS2.Checked)
                {
                    this.chartTemp.Series[1].ValueMemberX = "LogDate";
                    this.chartTemp.Series[1].ValueMembersY = "s2";
                }
                if (this.chkS3.Checked)
                {
                    this.chartTemp.Series[2].ValueMemberX = "LogDate";
                    this.chartTemp.Series[2].ValueMembersY = "s3";
                }
                if (this.chkS4.Checked)
                {
                    this.chartTemp.Series[3].ValueMemberX = "LogDate";
                    this.chartTemp.Series[3].ValueMembersY = "s4";
                }
                if (this.chkS5.Checked)
                {
                    this.chartTemp.Series[4].ValueMemberX = "LogDate";
                    this.chartTemp.Series[4].ValueMembersY = "s5";
                }
                if (this.chkS6.Checked)
                {
                    this.chartTemp.Series[5].ValueMemberX = "LogDate";
                    this.chartTemp.Series[5].ValueMembersY = "s6";
                }
                if (this.chkS7.Checked)
                {
                    this.chartPres.Series[0].ValueMemberX = "LogDate";
                    this.chartPres.Series[0].ValueMembersY = "s7";
                }
                if (this.chkS8.Checked)
                {
                    this.chartPres.Series[1].ValueMemberX = "LogDate";
                    this.chartPres.Series[1].ValueMembersY = "s8";
                }
                if (this.chkS9.Checked)
                {
                    this.chartPres.Series[2].ValueMemberX = "LogDate";
                    this.chartPres.Series[2].ValueMembersY = "s9";
                }
                if (this.chkS10.Checked)
                {
                    this.chartLev.Series[0].ValueMemberX = "LogDate";
                    this.chartLev.Series[0].ValueMembersY = "s10";
                }
                if (this.chkS11.Checked)
                {
                    this.chartLev.Series[1].ValueMemberX = "LogDate";
                    this.chartLev.Series[1].ValueMembersY = "s11";
                }
                if (this.chkS12.Checked)
                {
                    this.chartFlw.Series[0].ValueMemberX = "LogDate";
                    this.chartFlw.Series[0].ValueMembersY = "s12";
                }
                if (this.chkS13.Checked)
                {
                    this.chartFlw.Series[1].ValueMemberX = "LogDate";
                    this.chartFlw.Series[1].ValueMembersY = "s13";
                }
                if (this.chkS14.Checked)
                {
                    this.chartFlw.Series[2].ValueMemberX = "LogDate";
                    this.chartFlw.Series[2].ValueMembersY = "s14";
                }
                if (this.chkS15.Checked)
                {
                    this.chartFlw.Series[3].ValueMemberX = "LogDate";
                    this.chartFlw.Series[3].ValueMembersY = "s15";
                }
                if (this.chkS16.Checked)
                {
                    this.chartFlw.Series[4].ValueMemberX = "LogDate";
                    this.chartFlw.Series[4].ValueMembersY = "s16";
                }

                if (this.chkPrm2.Checked)
                {
                    this.chartFlw.Series[5].ValueMemberX = "LogDate";
                    this.chartFlw.Series[5].ValueMembersY = "p2";
                }

                if (this.chkPrm1.Checked)
                {
                    this.ChartFlowRate.Series[0].ValueMemberX = "LogDate";
                    this.ChartFlowRate.Series[0].ValueMembersY = "p1";
                }

                this.chartTemp.DataBind();
                this.chartPres.DataBind();
                this.chartLev.DataBind();
                this.chartFlw.DataBind();
                this.ChartFlowRate.DataBind();
                this.chartCustom1.DataBind();
                this.chartCustom2.DataBind();
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
                int num = (int)MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void readSensor()
        {
            try
            {
                if (this.con.State == ConnectionState.Closed)
                    this.con.Open();
                ((DbDataAdapter)new OleDbDataAdapter("select * from tblSensor order by idSensor", this.con)).Fill(this.dtSensor);
                this.cpuID = this.dtSensor.Rows[1]["sampleRate"].ToString();
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(("Error in readSensor(). " + ex.Message));
                int num = (int)MessageBox.Show("Error in readSensor().\n " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void readQueryCounter()
        {
            try
            {
                if (this.con.State == ConnectionState.Closed)
                    this.con.Open();
                this.queryCounter = int.Parse(new OleDbCommand("SELECT Max(flag) FROM tblData", this.con).ExecuteScalar().ToString());
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(("Error in readQueryCounter(). " + ex.Message));
                int num = (int)MessageBox.Show("Error in readQueryCounter().\n " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void bindFiled()
        {
            try
            {
                this.sampleRate = int.Parse(this.dtSensor.Rows[0]["sampleRate"].ToString());
                this.lblSampleRate.Text = sampleRate.ToString() + " Second";

                this.txtS1.Text = this.dtSensor.Rows[0]["captionU"].ToString();
                this.txtS2.Text = this.dtSensor.Rows[1]["captionU"].ToString();
                this.txtS3.Text = this.dtSensor.Rows[2]["captionU"].ToString();
                this.txtS4.Text = this.dtSensor.Rows[3]["captionU"].ToString();
                this.txtS5.Text = this.dtSensor.Rows[4]["captionU"].ToString();
                this.txtS6.Text = this.dtSensor.Rows[5]["captionU"].ToString();
                this.txtS7.Text = this.dtSensor.Rows[6]["captionU"].ToString();
                this.txtS8.Text = this.dtSensor.Rows[7]["captionU"].ToString();
                this.txtS9.Text = this.dtSensor.Rows[8]["captionU"].ToString();
                this.txtS10.Text = this.dtSensor.Rows[9]["captionU"].ToString();
                this.txtS11.Text = this.dtSensor.Rows[10]["captionU"].ToString();
                this.txtS12.Text = this.dtSensor.Rows[11]["captionU"].ToString();
                this.txtS13.Text = this.dtSensor.Rows[12]["captionU"].ToString();
                this.txtS14.Text = this.dtSensor.Rows[13]["captionU"].ToString();
                this.txtS15.Text = this.dtSensor.Rows[14]["captionU"].ToString();
                this.txtS16.Text = this.dtSensor.Rows[15]["captionU"].ToString();
                this.txtS17.Text = this.dtSensor.Rows[16]["captionU"].ToString();
                this.txtS18.Text = this.dtSensor.Rows[17]["captionU"].ToString();
                this.txtS19.Text = this.dtSensor.Rows[18]["captionU"].ToString();
                this.txtS20.Text = this.dtSensor.Rows[19]["captionU"].ToString();
                this.txtS21.Text = this.dtSensor.Rows[20]["captionU"].ToString();
                this.txtS22.Text = this.dtSensor.Rows[21]["captionU"].ToString();
                this.txtS23.Text = this.dtSensor.Rows[22]["captionU"].ToString();
                this.txtS24.Text = this.dtSensor.Rows[23]["captionU"].ToString();
                this.txtParam1.Text = this.dtSensor.Rows[24]["captionU"].ToString();
                this.txtParam2.Text = this.dtSensor.Rows[25]["captionU"].ToString();
                this.txtParam3.Text = this.dtSensor.Rows[26]["captionU"].ToString();
                this.txtParam4.Text = this.dtSensor.Rows[27]["captionU"].ToString();

                this.chkS1.Checked = bool.Parse(this.dtSensor.Rows[0]["enable"].ToString());
                this.chkS2.Checked = bool.Parse(this.dtSensor.Rows[1]["enable"].ToString());
                this.chkS3.Checked = bool.Parse(this.dtSensor.Rows[2]["enable"].ToString());
                this.chkS4.Checked = bool.Parse(this.dtSensor.Rows[3]["enable"].ToString());
                this.chkS5.Checked = bool.Parse(this.dtSensor.Rows[4]["enable"].ToString());
                this.chkS6.Checked = bool.Parse(this.dtSensor.Rows[5]["enable"].ToString());
                this.chkS7.Checked = bool.Parse(this.dtSensor.Rows[6]["enable"].ToString());
                this.chkS8.Checked = bool.Parse(this.dtSensor.Rows[7]["enable"].ToString());
                this.chkS9.Checked = bool.Parse(this.dtSensor.Rows[8]["enable"].ToString());
                this.chkS10.Checked = bool.Parse(this.dtSensor.Rows[9]["enable"].ToString());
                this.chkS11.Checked = bool.Parse(this.dtSensor.Rows[10]["enable"].ToString());
                this.chkS12.Checked = bool.Parse(this.dtSensor.Rows[11]["enable"].ToString());
                this.chkS13.Checked = bool.Parse(this.dtSensor.Rows[12]["enable"].ToString());
                this.chkS14.Checked = bool.Parse(this.dtSensor.Rows[13]["enable"].ToString());
                this.chkS15.Checked = bool.Parse(this.dtSensor.Rows[14]["enable"].ToString());
                this.chkS16.Checked = bool.Parse(this.dtSensor.Rows[15]["enable"].ToString());
                this.chkS17.Checked = bool.Parse(this.dtSensor.Rows[16]["enable"].ToString());
                this.chkS18.Checked = bool.Parse(this.dtSensor.Rows[17]["enable"].ToString());
                this.chkS19.Checked = bool.Parse(this.dtSensor.Rows[18]["enable"].ToString());
                this.chkS20.Checked = bool.Parse(this.dtSensor.Rows[19]["enable"].ToString());
                this.chkS21.Checked = bool.Parse(this.dtSensor.Rows[20]["enable"].ToString());
                this.chkS22.Checked = bool.Parse(this.dtSensor.Rows[21]["enable"].ToString());
                this.chkS23.Checked = bool.Parse(this.dtSensor.Rows[22]["enable"].ToString());
                this.chkS24.Checked = bool.Parse(this.dtSensor.Rows[23]["enable"].ToString());
                this.chkPrm1.Checked = bool.Parse(this.dtSensor.Rows[24]["enable"].ToString());
                this.chkPrm2.Checked = bool.Parse(this.dtSensor.Rows[25]["enable"].ToString());
                this.chkPrm3.Checked = bool.Parse(this.dtSensor.Rows[26]["enable"].ToString());
                this.chkPrm4.Checked = bool.Parse(this.dtSensor.Rows[27]["enable"].ToString());

                this.lbSU1.Text = this.dtSensor.Rows[0]["unitU"].ToString();
                this.lbSU2.Text = this.dtSensor.Rows[1]["unitU"].ToString();
                this.lbSU3.Text = this.dtSensor.Rows[2]["unitU"].ToString();
                this.lbSU4.Text = this.dtSensor.Rows[3]["unitU"].ToString();
                this.lbSU5.Text = this.dtSensor.Rows[4]["unitU"].ToString();
                this.lbSU6.Text = this.dtSensor.Rows[5]["unitU"].ToString();
                this.lbSU7.Text = this.dtSensor.Rows[6]["unitU"].ToString();
                this.lbSU8.Text = this.dtSensor.Rows[7]["unitU"].ToString();
                this.lbSU9.Text = this.dtSensor.Rows[8]["unitU"].ToString();
                this.lbSU10.Text = this.dtSensor.Rows[9]["unitU"].ToString();
                this.lbSU11.Text = this.dtSensor.Rows[10]["unitU"].ToString();
                this.lbSU12.Text = this.dtSensor.Rows[11]["unitU"].ToString();
                this.lbSU13.Text = this.dtSensor.Rows[12]["unitU"].ToString();
                this.lbSU14.Text = this.dtSensor.Rows[13]["unitU"].ToString();
                this.lbSU15.Text = this.dtSensor.Rows[14]["unitU"].ToString();
                this.lbSU16.Text = this.dtSensor.Rows[15]["unitU"].ToString();
                this.lbSU17.Text = this.dtSensor.Rows[16]["unitU"].ToString();
                this.lbSU18.Text = this.dtSensor.Rows[17]["unitU"].ToString();
                this.lbSU19.Text = this.dtSensor.Rows[18]["unitU"].ToString();
                this.lbSU20.Text = this.dtSensor.Rows[19]["unitU"].ToString();
                this.lbSU21.Text = this.dtSensor.Rows[20]["unitU"].ToString();
                this.lbSU22.Text = this.dtSensor.Rows[21]["unitU"].ToString();
                this.lbSU23.Text = this.dtSensor.Rows[22]["unitU"].ToString();
                this.lbSU24.Text = this.dtSensor.Rows[23]["unitU"].ToString();
                this.lbUP1.Text = this.dtSensor.Rows[24]["unitU"].ToString();
                this.lbUP2.Text = this.dtSensor.Rows[25]["unitU"].ToString();
                //this.lbUP3.Text = this.dtSensor.Rows[26]["unitU"].ToString();
                // this.lbUP4.Text = this.dtSensor.Rows[27]["unitU"].ToString();


                this.lbNS1.Text = this.dtSensor.Rows[0]["captionU"].ToString();
                this.lbNS2.Text = this.dtSensor.Rows[1]["captionU"].ToString();
                this.lbNS3.Text = this.dtSensor.Rows[2]["captionU"].ToString();
                this.lbNS4.Text = this.dtSensor.Rows[3]["captionU"].ToString();
                this.lbNS5.Text = this.dtSensor.Rows[4]["captionU"].ToString();
                this.lbNS6.Text = this.dtSensor.Rows[5]["captionU"].ToString();
                this.lbNS7.Text = this.dtSensor.Rows[6]["captionU"].ToString();
                this.lbNS8.Text = this.dtSensor.Rows[7]["captionU"].ToString();
                this.lbNS9.Text = this.dtSensor.Rows[8]["captionU"].ToString();
                this.lbNS10.Text = this.dtSensor.Rows[9]["captionU"].ToString();
                this.lbNS11.Text = this.dtSensor.Rows[10]["captionU"].ToString();
                this.lbNS12.Text = this.dtSensor.Rows[11]["captionU"].ToString();
                this.lbNS13.Text = this.dtSensor.Rows[12]["captionU"].ToString();
                this.lbNS14.Text = this.dtSensor.Rows[13]["captionU"].ToString();
                this.lbNS15.Text = this.dtSensor.Rows[14]["captionU"].ToString();
                this.lbNS16.Text = this.dtSensor.Rows[15]["captionU"].ToString();
                this.lbNS17.Text = this.dtSensor.Rows[16]["captionU"].ToString();
                this.lbNS18.Text = this.dtSensor.Rows[17]["captionU"].ToString();
                this.lbNS19.Text = this.dtSensor.Rows[18]["captionU"].ToString();
                this.lbNS20.Text = this.dtSensor.Rows[19]["captionU"].ToString();
                this.lbNS21.Text = this.dtSensor.Rows[20]["captionU"].ToString();
                this.lbNS22.Text = this.dtSensor.Rows[21]["captionU"].ToString();
                this.lbNS23.Text = this.dtSensor.Rows[22]["captionU"].ToString();
                this.lbNS24.Text = this.dtSensor.Rows[23]["captionU"].ToString();
                this.lbNP1.Text = this.dtSensor.Rows[24]["captionU"].ToString();
                this.lbNP2.Text = this.dtSensor.Rows[25]["captionU"].ToString();
                //this.lbNP3.Text = this.dtSensor.Rows[26]["captionU"].ToString();
                // this.lbNP4.Text = this.dtSensor.Rows[27]["captionU"].ToString();

                this.lbS1.Text = "";
                this.lbS2.Text = "";
                this.lbS3.Text = "";
                this.lbS4.Text = "";
                this.lbS5.Text = "";
                this.lbS6.Text = "";
                this.lbS7.Text = "";
                this.lbS8.Text = "";
                this.lbS9.Text = "";
                this.lbS10.Text = "";
                this.lbS11.Text = "";
                this.lbS12.Text = "";
                this.lbS13.Text = "";
                this.lbS14.Text = "";
                this.lbS15.Text = "";
                this.lbS16.Text = "";
                this.lbS17.Text = "";
                this.lbS18.Text = "";
                this.lbS19.Text = "";
                this.lbS20.Text = "";
                this.lbS21.Text = "";
                this.lbS22.Text = "";
                this.lbS23.Text = "";
                this.lbS24.Text = "";
                this.lbVP1.Text = "";
                this.lbVP2.Text = "";
                //this.lbVP3.Text = "";
                //this.lbVP4.Text = "";
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(("Error in BindField(). " + ex.Message));
                int num = (int)MessageBox.Show("Error in BindField().\n " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private string createSelectCommand()
        {
            try
            {
                this.st1 = "";
                this.st2 = "";
                this.st3 = "";
                this.st4 = "";
                this.st5 = "";
                this.st6 = "";
                this.st7 = "";
                this.st8 = "";
                this.st9 = "";
                this.st10 = "";
                this.st11 = "";
                this.st12 = "";
                this.st13 = "";
                this.st14 = "";
                this.st15 = "";
                this.st16 = "";
                this.st17 = "";
                this.st18 = "";
                this.st19 = "";
                this.st20 = "";
                this.st21 = "";
                this.st22 = "";
                this.st23 = "";
                this.st24 = "";
                this.param1 = "";
                this.param2 = "";
                this.param3 = "";
                this.param4 = "";



                if (this.chkS1.Checked)
                    this.st1 = "s1,";
                if (this.chkS2.Checked)
                    this.st2 = "s2,";
                if (this.chkS3.Checked)
                    this.st3 = "s3,";
                if (this.chkS4.Checked)
                    this.st4 = "s4,";
                if (this.chkS5.Checked)
                    this.st5 = "s5,";
                if (this.chkS6.Checked)
                    this.st6 = "s6,";
                if (this.chkS7.Checked)
                    this.st7 = "s7,";
                if (this.chkS8.Checked)
                    this.st8 = "s8,";
                if (this.chkS9.Checked)
                    this.st9 = "s9,";
                if (this.chkS10.Checked)
                    this.st10 = "s10,";
                if (this.chkS11.Checked)
                    this.st11 = "s11,";
                if (this.chkS12.Checked)
                    this.st12 = "s12,";
                if (this.chkS13.Checked)
                    this.st13 = "s13,";
                if (this.chkS14.Checked)
                    this.st14 = "s14,";
                if (this.chkS15.Checked)
                    this.st15 = "s15,";
                if (this.chkS16.Checked)
                    this.st16 = "s16,";
                if (this.chkS17.Checked)
                    this.st17 = "s17,";
                if (this.chkS18.Checked)
                    this.st18 = "s18,";
                if (this.chkS19.Checked)
                    this.st19 = "s19,";
                if (this.chkS20.Checked)
                    this.st20 = "s20,";
                if (this.chkS21.Checked)
                    this.st21 = "s21,";
                if (this.chkS22.Checked)
                    this.st22 = "s22,";
                if (this.chkS23.Checked)
                    this.st23 = "s23,";
                if (this.chkS24.Checked)
                    this.st24 = "s24,";
                if (this.chkPrm1.Checked)
                    this.param1 = "p1,";
                if (this.chkPrm2.Checked)
                    this.param2 = "p2,";
                if (this.chkPrm3.Checked)
                    this.param3 = "p3,";
                if (this.chkPrm4.Checked)
                    this.param4 = "p4,";
                string str = "select id,logDate,time,comment, " + this.st1 + this.st2 + this.st3 + this.st4 + this.st5 + this.st6 + this.st7 + this.st8 + this.st9 + this.st10 + this.st11 + this.st12 + this.st13 + this.st14 + this.st15 + this.st16 + this.st17 + this.st18 + this.st19 + this.st20 + this.st21 + this.st22 + this.st23 + this.st24 + this.param1 + this.param2 + this.param3 + this.param4;
                return str.Remove(str.Length - 1, 1) + " from tblData";
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(("Error in createSelectCommand(). " + ex.Message));
                int num = (int)MessageBox.Show("Error in createSelectCommand().\n " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return null;
            }
        }

        private bool fillGird(string strSql)
        {
            try
            {
                if (this.con.State == ConnectionState.Closed)
                    this.con.Open();
                OleDbDataReader oleDbDataReader = new OleDbCommand(strSql, this.con).ExecuteReader();
                this.tblData.Reset();
                this.tblData.Dispose();
                this.tblData = new System.Data.DataTable();
                this.tblData.Load((IDataReader)oleDbDataReader);
                this.rdbGrid1.DataSource = this.tblData;
                foreach (GridViewDataColumn gridViewDataColumn in (Collection<GridViewDataColumn>)this.rdbGrid1.Columns)
                {
                    gridViewDataColumn.ReadOnly = true;
                    gridViewDataColumn.TextAlignment = ContentAlignment.MiddleCenter;
                    gridViewDataColumn.HeaderTextAlignment = ContentAlignment.MiddleCenter;
                    gridViewDataColumn.StretchVertically = true;
                }
                //this.rdbGrid1.Columns["comment"].ReadOnly = false;
                this.rdbGrid1.Columns["logDate"].FormatString = "{0:d}";
                this.rdbGrid1.Columns["time"].FormatString = "{0:HH:mm:ss}";
                this.dv = new DataView(this.tblData);
                this.cm = (CurrencyManager)this.BindingContext[this.dv];
                this.showPosition();
                return true;
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(("Error in FillGrid(). " + ex.Message));
                int num = (int)MessageBox.Show("Error in FillGrid().\n " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return false;
            }
        }

        private void assignCaption()
        {
            try
            {
                if (this.chkS1.Checked)
                {
                    this.tblData.Columns["s1"].Caption = this.txtS1.Text;
                    this.rdbGrid1.Columns["s1"].HeaderText = this.txtS1.Text;
                }
                if (this.chkS2.Checked)
                {
                    this.tblData.Columns["s2"].Caption = this.txtS2.Text;
                    this.rdbGrid1.Columns["s2"].HeaderText = this.txtS2.Text;
                }
                if (this.chkS3.Checked)
                {
                    this.tblData.Columns["s3"].Caption = this.txtS3.Text;
                    this.rdbGrid1.Columns["s3"].HeaderText = this.txtS3.Text;
                }
                if (this.chkS4.Checked)
                {
                    this.tblData.Columns["s4"].Caption = this.txtS4.Text;
                    this.rdbGrid1.Columns["s4"].HeaderText = this.txtS4.Text;
                }
                if (this.chkS5.Checked)
                {
                    this.tblData.Columns["s5"].Caption = this.txtS5.Text;
                    this.rdbGrid1.Columns["s5"].HeaderText = this.txtS5.Text;
                }
                if (this.chkS6.Checked)
                {
                    this.tblData.Columns["s6"].Caption = this.txtS6.Text;
                    this.rdbGrid1.Columns["s6"].HeaderText = this.txtS6.Text;
                }
                if (this.chkS7.Checked)
                {
                    this.tblData.Columns["s7"].Caption = this.txtS7.Text;
                    this.rdbGrid1.Columns["s7"].HeaderText = this.txtS7.Text;
                }
                if (this.chkS8.Checked)
                {
                    this.tblData.Columns["s8"].Caption = this.txtS8.Text;
                    this.rdbGrid1.Columns["s8"].HeaderText = this.txtS8.Text;
                }
                if (this.chkS9.Checked)
                {
                    this.tblData.Columns["s9"].Caption = this.txtS9.Text;
                    this.rdbGrid1.Columns["s9"].HeaderText = this.txtS9.Text;
                }
                if (this.chkS10.Checked)
                {
                    this.tblData.Columns["s10"].Caption = this.txtS10.Text;
                    this.rdbGrid1.Columns["s10"].HeaderText = this.txtS10.Text;
                }
                if (this.chkS11.Checked)
                {
                    this.tblData.Columns["s11"].Caption = this.txtS11.Text;
                    this.rdbGrid1.Columns["s11"].HeaderText = this.txtS11.Text;
                }
                if (this.chkS12.Checked)
                {
                    this.tblData.Columns["s12"].Caption = this.txtS12.Text;
                    this.rdbGrid1.Columns["s12"].HeaderText = this.txtS12.Text;
                }
                if (this.chkS13.Checked)
                {
                    this.tblData.Columns["s13"].Caption = this.txtS13.Text;
                    this.rdbGrid1.Columns["s13"].HeaderText = this.txtS13.Text;
                }
                if (this.chkS14.Checked)
                {
                    this.tblData.Columns["s14"].Caption = this.txtS14.Text;
                    this.rdbGrid1.Columns["s14"].HeaderText = this.txtS14.Text;
                }
                if (this.chkS15.Checked)
                {
                    this.tblData.Columns["s15"].Caption = this.txtS15.Text;
                    this.rdbGrid1.Columns["s15"].HeaderText = this.txtS15.Text;
                }
                if (this.chkS16.Checked)
                {
                    this.tblData.Columns["s16"].Caption = this.txtS16.Text;
                    this.rdbGrid1.Columns["s16"].HeaderText = this.txtS16.Text;
                }
                if (this.chkS17.Checked)
                {
                    this.tblData.Columns["s17"].Caption = this.txtS17.Text;
                    this.rdbGrid1.Columns["s17"].HeaderText = this.txtS17.Text;
                }
                if (this.chkS18.Checked)
                {
                    this.tblData.Columns["s18"].Caption = this.txtS18.Text;
                    this.rdbGrid1.Columns["s18"].HeaderText = this.txtS18.Text;
                }
                if (this.chkS19.Checked)
                {
                    this.tblData.Columns["s19"].Caption = this.txtS19.Text;
                    this.rdbGrid1.Columns["s19"].HeaderText = this.txtS19.Text;
                }
                if (this.chkS20.Checked)
                {
                    this.tblData.Columns["s20"].Caption = this.txtS20.Text;
                    this.rdbGrid1.Columns["s20"].HeaderText = this.txtS20.Text;
                }
                if (this.chkS21.Checked)
                {
                    this.tblData.Columns["s21"].Caption = this.txtS21.Text;
                    this.rdbGrid1.Columns["s21"].HeaderText = this.txtS21.Text;
                }
                if (this.chkS22.Checked)
                {
                    this.tblData.Columns["s22"].Caption = this.txtS22.Text;
                    this.rdbGrid1.Columns["s22"].HeaderText = this.txtS22.Text;
                }
                if (this.chkS23.Checked)
                {
                    this.tblData.Columns["s23"].Caption = this.txtS23.Text;
                    this.rdbGrid1.Columns["s23"].HeaderText = this.txtS23.Text;
                }
                if (this.chkS24.Checked)
                {
                    this.tblData.Columns["s24"].Caption = this.txtS24.Text;
                    this.rdbGrid1.Columns["s24"].HeaderText = this.txtS24.Text;
                }

                if (this.chkPrm1.Checked)
                {
                    this.tblData.Columns["p1"].Caption = this.txtParam1.Text;
                    this.rdbGrid1.Columns["p1"].HeaderText = this.txtParam1.Text;
                }
                if (this.chkPrm2.Checked)
                {
                    this.tblData.Columns["p2"].Caption = this.txtParam2.Text;
                    this.rdbGrid1.Columns["p2"].HeaderText = this.txtParam2.Text;
                }
                if (this.chkPrm3.Checked)
                {
                    this.tblData.Columns["p3"].Caption = this.txtParam3.Text;
                    this.rdbGrid1.Columns["p3"].HeaderText = this.txtParam3.Text;
                }
                if (this.chkPrm4.Checked)
                {
                    this.tblData.Columns["p4"].Caption = this.txtParam4.Text;
                    this.rdbGrid1.Columns["p4"].HeaderText = this.txtParam4.Text;
                }




                this.lbNS1.Text = this.txtS1.Text;
                this.lbNS2.Text = this.txtS2.Text;
                this.lbNS3.Text = this.txtS3.Text;
                this.lbNS4.Text = this.txtS4.Text;
                this.lbNS5.Text = this.txtS5.Text;
                this.lbNS6.Text = this.txtS6.Text;
                this.lbNS7.Text = this.txtS7.Text;
                this.lbNS8.Text = this.txtS8.Text;
                this.lbNS9.Text = this.txtS9.Text;
                this.lbNS10.Text = this.txtS10.Text;
                this.lbNS11.Text = this.txtS11.Text;
                this.lbNS12.Text = this.txtS12.Text;
                this.lbNS13.Text = this.txtS13.Text;
                this.lbNS14.Text = this.txtS14.Text;
                this.lbNS15.Text = this.txtS15.Text;
                this.lbNS16.Text = this.txtS16.Text;
                this.lbNS17.Text = this.txtS17.Text;
                this.lbNS18.Text = this.txtS18.Text;
                this.lbNS19.Text = this.txtS19.Text;
                this.lbNS20.Text = this.txtS20.Text;
                this.lbNS21.Text = this.txtS21.Text;
                this.lbNS22.Text = this.txtS22.Text;
                this.lbNS23.Text = this.txtS23.Text;
                this.lbNS24.Text = this.txtS24.Text;
                this.lbNP1.Text = this.txtParam1.Text;
                this.lbNP2.Text = this.txtParam2.Text;
                // this.lbNP3.Text = this.txtParam3.Text;
                //this.lbNP4.Text = this.txtParam4.Text;

                this.chartTemp.Series[0].Name = this.txtS1.Text;
                this.chartTemp.Series[1].Name = this.txtS2.Text;
                this.chartTemp.Series[2].Name = this.txtS3.Text;
                this.chartTemp.Series[3].Name = this.txtS4.Text;
                this.chartTemp.Series[4].Name = this.txtS5.Text;
                this.chartTemp.Series[5].Name = this.txtS6.Text;

                this.chartPres.Series[0].Name = this.txtS7.Text;
                this.chartPres.Series[1].Name = this.txtS8.Text;
                this.chartPres.Series[2].Name = this.txtS9.Text;

                this.chartLev.Series[0].Name = this.txtS10.Text;
                this.chartLev.Series[1].Name = this.txtS11.Text;

                this.chartFlw.Series[0].Name = this.txtS12.Text;
                this.chartFlw.Series[1].Name = this.txtS13.Text;
                this.chartFlw.Series[2].Name = this.txtS14.Text;
                this.chartFlw.Series[3].Name = this.txtS15.Text;
                this.chartFlw.Series[4].Name = this.txtS16.Text;

                this.chartFlw.Series[5].Name = this.txtParam2.Text;

                this.ChartFlowRate.Series[0].Name = this.txtParam1.Text;

                this.chartCustom1.Series[0].Name = this.txtS17.Text;
                this.chartCustom1.Series[1].Name = this.txtS18.Text;
                this.chartCustom1.Series[2].Name = this.txtS19.Text;
                this.chartCustom1.Series[3].Name = this.txtS20.Text;

                this.chartCustom2.Series[0].Name = this.txtS21.Text;
                this.chartCustom2.Series[1].Name = this.txtS22.Text;
                this.chartCustom2.Series[2].Name = this.txtS23.Text;
                this.chartCustom2.Series[3].Name = this.txtS24.Text;
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(("Error in AssignCaption(). " + ex.Message));
                int num = (int)MessageBox.Show("Error in AssignCaption().\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void showPosition()
        {
            this.txtPosition.Text = this.cm.Count.ToString();
        }

        public static void ExportDataTableToExcel(System.Data.DataTable dataTable, Worksheet sheetToAddTo)
        {

            object[,] objArray1 = new object[1, dataTable.Columns.Count];
            for (int index = 0; index < dataTable.Columns.Count; ++index)
                objArray1[0, index] = dataTable.Columns[index].Caption;
            Microsoft.Office.Interop.Excel.Range range = sheetToAddTo.get_Range(sheetToAddTo.Cells[1, 1], sheetToAddTo.Cells[1, dataTable.Columns.Count]);
            range.Value2 = objArray1;
            range.EntireRow.Font.Bold = true;
            Marshal.ReleaseComObject(range);
            object[,] objArray2 = new object[dataTable.Rows.Count, dataTable.Columns.Count];
            Form1.m_dataRowCount = dataTable.Rows.Count;
            for (int index1 = 0; index1 < dataTable.Rows.Count; ++index1)
            {
                for (int index2 = 0; index2 < dataTable.Columns.Count; ++index2)
                    objArray2[index1, index2] = dataTable.Rows[index1][index2];
            }
            Microsoft.Office.Interop.Excel.Range excelRange = sheetToAddTo.get_Range(sheetToAddTo.Cells[2, 1], sheetToAddTo.Cells[(dataTable.Rows.Count + 1), dataTable.Columns.Count]);
            short num = (short)1;
            string str = string.Empty;
            foreach (DataColumn dataColumn in (InternalDataCollectionBase)dataTable.Columns)
            {
                string format = string.Empty;
                if (dataColumn.DataType.Equals(typeof(string)))
                    format = "@";
                else if (dataColumn.DataType.Equals(typeof(DateTime)))
                    format = "dd-MM-yyyy";
                if (!string.IsNullOrEmpty(format))
                    Form1.FormatColumn(excelRange, (int)num, format);
                ++num;
            }
            excelRange.Value2 = objArray2;
            Marshal.ReleaseComObject(excelRange);

        }

        public static void FormatColumn(Microsoft.Office.Interop.Excel.Range excelRange, int col, string format)
        {
            ((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[1, col]).EntireColumn.NumberFormat = format;
        }

        private bool ExportData(string fileNameExcell)
        {
            string Filename = fileNameExcell;
            Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)null;
            Workbook workbook = (Workbook)null;
            Worksheet sheetToAddTo = (Worksheet)null;
            object obj = Missing.Value;
            try
            {
                application = (Microsoft.Office.Interop.Excel.Application)new ApplicationClass();
                workbook = application.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                sheetToAddTo = (Worksheet)workbook.Worksheets[1];
                application.DisplayAlerts = false;
                application.AlertBeforeOverwriting = false;
                workbook.SaveAs(Filename, obj, obj, obj, obj, obj, XlSaveAsAccessMode.xlShared, obj, obj, obj, obj, obj);
                Form1.ExportDataTableToExcel(this.tblData, sheetToAddTo);
                sheetToAddTo.Cells.EntireColumn.AutoFit();
                sheetToAddTo.SaveAs(Filename, obj, obj, obj, obj, obj, obj, obj, obj, obj);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message, System.Windows.Forms.Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return false;
            }
            finally
            {
                application.Quit();
                Marshal.ReleaseComObject(sheetToAddTo);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(application);
                GC.Collect();
            }
            return true;
        }

        private void runQuery()
        {
            try
            {
                int Step = (int)(this.numHH.Value * new Decimal(3600) + this.numMM.Value * new Decimal(60) + this.numSS.Value);
                bool flag = Step == this.sampleRate;
                if (Step == 0 || Step < this.sampleRate)
                {
                    this.lstReport.Items.Add("Please confirm Step");
                    MessageBox.Show("Please confirm Step", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.mapToNumeric(this.sampleRate);
                }
                else
                {
                    int remain = Step % this.sampleRate;
                    if (remain != 0)
                    {
                        this.mapToNumeric(Step - remain);
                        this.lstReport.Items.Add("Please confirm Step");
                        MessageBox.Show("Please confirm Step", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else
                    {
                        int UserStep = Step / this.sampleRate;
                        ++this.queryCounter;
                        if (this.con.State == ConnectionState.Closed)
                            this.con.Open();

                        int Id_Min_Date = int.Parse(new OleDbCommand("SELECT MIN(id) FROM tblData WHERE ( LogDate>=#" + this.dtFrom.Value.AddMilliseconds(-30.0) + "#)", this.con).ExecuteScalar().ToString());
                        int Id_Max_Date = int.Parse(new OleDbCommand("SELECT MAX(id) FROM tblData WHERE ( LogDate<=#" + this.dtTo.Value.AddMilliseconds(30.0) + " #)", this.con).ExecuteScalar().ToString());
                        int k = UserStep;

                        if (!flag)   //step != sampleRate
                        {
                            OleDbCommand oleDbCommand = new OleDbCommand();
                            if (this.con.State == ConnectionState.Closed)
                                this.con.Open();
                            oleDbCommand.CommandText = "UPDATE tblData SET flag = @queryCounter where ID =@id ";
                            oleDbCommand.Connection = this.con;
                            for (int i = Id_Min_Date; i <= Id_Max_Date; ++i)
                            {
                                if (k == UserStep)
                                {
                                    oleDbCommand.Parameters.Clear();
                                    oleDbCommand.Parameters.AddWithValue("@queryCounter", this.queryCounter);
                                    oleDbCommand.Parameters.AddWithValue("@id", i);
                                    oleDbCommand.ExecuteNonQuery();
                                    k = 0;
                                }
                                ++k;
                            }
                            this.fillGird(this.createSelectCommand() + " WHERE flag = " + this.queryCounter + this.strOrder);
                            this.assignCaption();
                        }
                        else
                        {
                            this.fillGird(this.createSelectCommand() + " where ( LogDate>=#" + this.dtFrom.Value.AddMilliseconds(-10.0) + " # And (LogDate)<=#" + this.dtTo.Value.AddMilliseconds(10.0) + " #)  order by id");
                            this.assignCaption();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add("please check data for query.");
                int num = (int)MessageBox.Show("please check data for query.\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void mOpen_Click(object sender, EventArgs e)
        {
            try
            {
                this.stProgramm = Form1.stateProgram.CloseFile;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Access 2007|*.accdb";
                openFileDialog.Title = "Open Access Database";
                int num1 = (int)openFileDialog.ShowDialog();
                if (openFileDialog.FileName != "")
                {
                    this.dbName = (openFileDialog.FileName).ToString();
                    this.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + this.dbName;
                    this.con.ConnectionString = this.connectionString;
                    this.strSelectCommand = this.connectionString;
                    this.readSensor();
                    if (this.cpuID != this.GetCPUId())
                    {
                        int num2 = (int)MessageBox.Show("please check hardware . . .", "Check", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        if (this.con.State == ConnectionState.Open)
                            this.con.Close();
                        this.btnShow();
                    }
                    else
                    {
                        if (!CheckDellID())
                        {
                            MessageBox.Show("SE-Error\nPlease contact to administrator. . .", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            if (this.con.State == ConnectionState.Open)
                                this.con.Close();
                            this.btnShow();
                        }
                        else
                        {
                            this.readQueryCounter();
                            this.bindFiled();
                            this.fillGird(this.createSelectCommand() + this.strOrder);

                            this.assignCaption();
                            this.dtPlayBack.Value = DateTime.Parse(this.rdbGrid1.Rows[0].Cells[2].Value.ToString());
                            this.dtFrom.Value = this.dtPlayBack.Value;
                            this.dtTo.Value = DateTime.Parse(this.rdbGrid1.Rows[this.rdbGrid1.RowCount - 1].Cells[2].Value.ToString());
                            this.stProgramm = Form1.stateProgram.openFile;
                            this.lblName.Text = (this.dbName).ToString();
                            this.IsfileOpen = true;
                            this.btnShow();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.stProgramm = Form1.stateProgram.CloseFile;
                this.lstReport.Items.Add(ex.Message);
                int num = (int)MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                this.btnShow();
            }
        }

        private void mClose_Click(object sender, EventArgs e)
        {
            try
            {
                this.stProgramm = Form1.stateProgram.CloseFile;
                DataSet dataSet = new DataSet();
                this.con.Close();
                this.rdbGrid1.DataSource = dataSet;
                this.btnShow();
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
                int num = (int)MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void mSave_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XML files|*.xml|All files|*.*";
                saveFileDialog.FileName = "";
                if (saveFileDialog.ShowDialog() != DialogResult.OK)
                    return;
                this.radDock1.SaveToXml(saveFileDialog.FileName);
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
                int num = (int)MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void mLoadSchema_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "XML files|*.xml|All files|*.*";
                if (openFileDialog.ShowDialog() != DialogResult.OK)
                    return;
                this.radDock1.LoadFromXml(openFileDialog.FileName);
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
                int num = (int)MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void mExport_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel 2007|*.xlsx";
                saveFileDialog.Title = "Save As . . .";
                int num = (int)saveFileDialog.ShowDialog();
                if (!(saveFileDialog.FileName != ""))
                    return;
                this.excellName = (saveFileDialog.FileName).ToString();
                this.ExportData(this.excellName);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void mExit_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void mChart_Click(object sender, EventArgs e)
        {
            try
            {
                this.radDock1.LoadFromXml((System.Windows.Forms.Application.StartupPath).ToString() + "\\pchartView.xml");
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
                int num = (int)MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void mDefault_Click(object sender, EventArgs e)
        {
            try
            {
                this.radDock1.LoadFromXml((System.Windows.Forms.Application.StartupPath).ToString() + "\\pdefaultView.xml");
            }
            catch (Exception ex)
            {
                this.lstReport.Items.Add(ex.Message);
                int num = (int)MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void mAbout_Click(object sender, EventArgs e)
        {
            //int num = (int)new AboutBox1().ShowDialog();
        }

        private void rdbGrid1_CurrentRowChanged(object sender, CurrentRowChangedEventArgs e)
        {
            if (this.rdbGrid1.CurrentRow == null || !(this.rdbGrid1.CurrentRow is GridViewDataRowInfo))
                return;
            this.txtComment.Text = this.rdbGrid1.CurrentRow.Cells["comment"].Value.ToString();
            this.txtRow.Text = this.rdbGrid1.CurrentRow.Cells["id"].Value.ToString();
        }







        //Chk Temp
        private void chkCS1_CheckedChanged(object sender, EventArgs e)
        {
            this.chartTemp.Series[0].Enabled = this.chkCS1.Checked;
        }

        private void chkCS2_CheckedChanged(object sender, EventArgs e)
        {
            this.chartTemp.Series[1].Enabled = this.chkCS2.Checked;
        }

        private void chkCS3_CheckedChanged(object sender, EventArgs e)
        {
            this.chartTemp.Series[2].Enabled = this.chkCS3.Checked;
        }

        private void chkCS4_CheckedChanged(object sender, EventArgs e)
        {
            this.chartTemp.Series[3].Enabled = this.chkCS4.Checked;
        }

        private void chkCS5_CheckedChanged(object sender, EventArgs e)
        {
            this.chartTemp.Series[4].Enabled = this.chkCS5.Checked;
        }

        private void chkCS6_CheckedChanged(object sender, EventArgs e)
        {
            this.chartTemp.Series[5].Enabled = this.chkCS6.Checked;
        }


        // Chk Pressure
        private void chkCS7_CheckedChanged(object sender, EventArgs e)
        {
            this.chartPres.Series[0].Enabled = this.chkCS7.Checked;
        }

        private void chkCS8_CheckedChanged(object sender, EventArgs e)
        {
            this.chartPres.Series[1].Enabled = this.chkCS8.Checked;
        }

        private void chkCS9_CheckedChanged(object sender, EventArgs e)
        {
            this.chartPres.Series[2].Enabled = this.chkCS9.Checked;
        }


        //chk Level
        private void chkCS10_CheckedChanged(object sender, EventArgs e)
        {
            this.chartLev.Series[0].Enabled = this.chkCS10.Checked;
        }

        private void chkCS11_CheckedChanged(object sender, EventArgs e)
        {
            this.chartLev.Series[1].Enabled = this.chkCS11.Checked;
        }


        //Chk Flow
        private void chkCS12_CheckedChanged(object sender, EventArgs e)
        {
            this.chartFlw.Series[0].Enabled = this.chkCS12.Checked;
        }

        private void chkCS13_CheckedChanged(object sender, EventArgs e)
        {
            this.chartFlw.Series[1].Enabled = this.chkCS13.Checked;
        }

        private void chkCS14_CheckedChanged(object sender, EventArgs e)
        {
            this.chartFlw.Series[2].Enabled = this.chkCS14.Checked;
        }

        private void chkCS15_CheckedChanged(object sender, EventArgs e)
        {
            this.chartFlw.Series[3].Enabled = this.chkCS15.Checked;
        }

        private void chkCS16_CheckedChanged(object sender, EventArgs e)
        {
            this.chartFlw.Series[4].Enabled = this.chkCS16.Checked;
        }

        private void chkP2_CheckedChanged(object sender, EventArgs e)
        {
            this.chartFlw.Series[5].Enabled = this.chkP2.Checked;
        }


        //chk Custom 1
        private void chkCS17_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.chartCustom1.Series[0].Enabled = this.chkCS17.Checked;
            }
            catch { }
        }

        private void chkCS18_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.chartCustom1.Series[1].Enabled = this.chkCS18.Checked;
            }
            catch { }
        }

        private void chkCS19_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.chartCustom1.Series[2].Enabled = this.chkCS19.Checked;
            }
            catch { }
        }

        private void chkCS20_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.chartCustom1.Series[3].Enabled = this.chkCS20.Checked;
            }
            catch { }
        }


        //chk Custom 2
        private void chkCS21_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.chartCustom2.Series[0].Enabled = this.chkCS21.Checked;
            }
            catch { }
        }

        private void chkCS22_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.chartCustom2.Series[1].Enabled = this.chkCS22.Checked;
            }
            catch { }
        }

        private void chkCS23_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.chartCustom2.Series[2].Enabled = this.chkCS23.Checked;
            }
            catch { }
        }

        private void chkCS24_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.chartCustom2.Series[3].Enabled = this.chkCS24.Checked;
            }
            catch { }
        }


        //chk Flowrate
        private void chkP1_CheckedChanged(object sender, EventArgs e)
        {
            this.ChartFlowRate.Series[0].Enabled = this.chkP1.Checked;
        }







        //btn Temp
        private void btnS1_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS1.BackColor = this.radColorDialog1.SelectedColor;
            this.chartTemp.Series[0].Color = this.btnS1.BackColor;
        }

        private void btnS2_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS2.BackColor = this.radColorDialog1.SelectedColor;
            this.chartTemp.Series[1].Color = this.btnS2.BackColor;
        }

        private void btnS3_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS3.BackColor = this.radColorDialog1.SelectedColor;
            this.chartTemp.Series[2].Color = this.btnS3.BackColor;
        }

        private void btnS4_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS4.BackColor = this.radColorDialog1.SelectedColor;
            this.chartTemp.Series[3].Color = this.btnS4.BackColor;
        }

        private void btnS5_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS5.BackColor = this.radColorDialog1.SelectedColor;
            this.chartTemp.Series[4].Color = this.btnS5.BackColor;
        }

        private void btnS6_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS6.BackColor = this.radColorDialog1.SelectedColor;
            this.chartTemp.Series[5].Color = this.btnS6.BackColor;
        }

        //btn Pressure
        private void btnS7_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS7.BackColor = this.radColorDialog1.SelectedColor;
            this.chartPres.Series[0].Color = this.btnS7.BackColor;
        }

        private void btnS8_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS8.BackColor = this.radColorDialog1.SelectedColor;
            this.chartPres.Series[1].Color = this.btnS8.BackColor;
        }

        private void btnS9_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS9.BackColor = this.radColorDialog1.SelectedColor;
            this.chartPres.Series[2].Color = this.btnS9.BackColor;
        }

        //btn Level
        private void btnS10_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS10.BackColor = this.radColorDialog1.SelectedColor;
            this.chartLev.Series[0].Color = this.btnS10.BackColor;
        }

        private void btnS11_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS11.BackColor = this.radColorDialog1.SelectedColor;
            this.chartLev.Series[1].Color = this.btnS11.BackColor;
        }

        //btn Flow
        private void btnS12_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS12.BackColor = this.radColorDialog1.SelectedColor;
            this.chartFlw.Series[0].Color = this.btnS12.BackColor;
        }

        private void btnS13_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS13.BackColor = this.radColorDialog1.SelectedColor;
            this.chartFlw.Series[1].Color = this.btnS13.BackColor;
        }

        private void btnS14_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS14.BackColor = this.radColorDialog1.SelectedColor;
            this.chartFlw.Series[2].Color = this.btnS14.BackColor;
        }

        private void btnS15_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS15.BackColor = this.radColorDialog1.SelectedColor;
            this.chartFlw.Series[3].Color = this.btnS15.BackColor;
        }

        private void btnS16_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnS16.BackColor = this.radColorDialog1.SelectedColor;
            this.chartFlw.Series[4].Color = this.btnS16.BackColor;
        }

        private void btnP2_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnP2.BackColor = this.radColorDialog1.SelectedColor;
            this.chartFlw.Series[5].Color = this.btnP2.BackColor;
        }

        private void btnP1_Click(object sender, EventArgs e)
        {
            if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                return;
            this.btnP1.BackColor = this.radColorDialog1.SelectedColor;
            this.ChartFlowRate.Series[0].Color = this.btnP1.BackColor;
        }

        //btn custom 1
        private void btnS17_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                    return;
                this.btnS17.BackColor = this.radColorDialog1.SelectedColor;
                this.chartCustom1.Series[0].Color = this.btnS17.BackColor;
            }
            catch { }
        }
        private void btnS18_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                    return;
                this.btnS18.BackColor = this.radColorDialog1.SelectedColor;
                this.chartCustom1.Series[1].Color = this.btnS18.BackColor;
            }
            catch { }
        }
        private void btnS19_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                    return;
                this.btnS19.BackColor = this.radColorDialog1.SelectedColor;
                this.chartCustom1.Series[2].Color = this.btnS19.BackColor;
            }
            catch { }
        }
        private void btnS20_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                    return;
                this.btnS20.BackColor = this.radColorDialog1.SelectedColor;
                this.chartCustom1.Series[3].Color = this.btnS20.BackColor;
            }
            catch { }
        }


        //btn custom 2
        private void btnS21_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                    return;
                this.btnS21.BackColor = this.radColorDialog1.SelectedColor;
                this.chartCustom2.Series[0].Color = this.btnS21.BackColor;
            }
            catch { }
        }
        private void btnS22_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                    return;
                this.btnS22.BackColor = this.radColorDialog1.SelectedColor;
                this.chartCustom2.Series[1].Color = this.btnS22.BackColor;
            }
            catch { }
        }
        private void btnS23_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                    return;
                this.btnS23.BackColor = this.radColorDialog1.SelectedColor;
                this.chartCustom2.Series[2].Color = this.btnS23.BackColor;
            }
            catch { }
        }

        private void btnS24_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.radColorDialog1.ShowDialog() != DialogResult.OK)
                    return;
                this.btnS24.BackColor = this.radColorDialog1.SelectedColor;
                this.chartCustom2.Series[3].Color = this.btnS24.BackColor;
            }
            catch
            {
            }
        }



        private void chkTmpLegend_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            this.chartTemp.Legends[0].Enabled = this.chkTmpLegend.Checked;
        }

        private void chkPresLegend_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            this.chartPres.Legends[0].Enabled = this.chkPresLegend.Checked;
        }

        private void chkLevLegend_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            this.chartLev.Legends[0].Enabled = this.chkLevLegend.Checked;
        }

        private void chkFlwLegend_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            this.chartFlw.Legends[0].Enabled = this.chkFlwLegend.Checked;
        }

        private void ChkFlowRateLegend_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            this.ChartFlowRate.Legends[0].Enabled = this.ChkFlowRateLegend.Checked;
        }

        private void chkCustom1Legend_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            this.chartCustom1.Legends[0].Enabled = this.chkCustom1Legend.Checked;
        }

        private void chkCustom2Legend_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            this.chartCustom2.Legends[0].Enabled = this.chkCustom2Legend.Checked;
        }


        private void button1_Click(object sender, EventArgs e)
        {

            chartPres.Series[0].MarkerSize = 5;
            // chartPres.Series[0].Points[50].Label = "jkjhwerhwejklrjhkwer";

            //  chartPres.Series[0].Points[50].AxisLabel = "aaaaaaaaaaa";
            chartPres.Series[0].ShowLabelAsValue = true;

            // Set axis label 
            chartPres.Series[0].Points[2].AxisLabel = "My Axis Label\nLabel Line #2";

            // Set data point label
            chartPres.Series[0].Points[2].Label = "My Point Label\nLabel Line #2";
            chartPres.Series[0].Points[2].ShowLabelAsValue = true;
            chartPres.Invalidate();

            // chartPres.Invalidate();
        }














    }
}
