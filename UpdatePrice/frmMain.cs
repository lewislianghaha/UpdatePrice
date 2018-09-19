using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using UpdatePrice.Db;
using UpdatePrice.Interface;

namespace UpdatePrice
{
    public partial class FrmMain : Form
    {
        DbCon conn = new DbCon();
        public int LastInterid = 0;
        public int Lastcheck = 0;   //获取最后一次点击单选按钮

        public FrmMain()
        {
            InitializeComponent();
            OnRegisterEvents();
        }

        private void OnRegisterEvents()
        {
            comList.Click += comList_Click;
            comCust.Click += comCust_Click;
            btnOpen.Click += btnOpen_Click;
            btnImport.Click += btnImport_Click;
            btnExit.Click += btnExit_Click;
            rdjin.CheckedChanged += rdjin_CheckedChanged;
        }

        /// <summary>
        /// 单选按钮改变时执行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void rdjin_CheckedChanged(object sender, EventArgs e)
        {
            var check = rdjin.Checked ? 0 : 1;
            if (check != Lastcheck)
            {
                comList.Text = "";
                comCust.Text = "";
                LastInterid = 0;
            }
            Lastcheck = check;
        }

        /// <summary>
        ///获取连接信息 
        /// </summary>
        /// <returns></returns>
        public SqlConnection GetConnection()
        {
            //选择了晶创或江苏帐套的设置(0:晶创;1:江苏)
            var checkid = rdjin.Checked ? 0 : 1;
            var constr = checkid == 0 ? "AIS20180317163111" : "AIS20180507150305";
            var sqlconn = conn.GetConnect(constr);
            return sqlconn;
        }

        /// <summary>
        /// 设置客户列表信息
        /// </summary>
        /// <param name="sqlconn"></param>
        /// <param name="finterid"></param>
        /// <returns></returns>
        int GetCustToList(SqlConnection sqlconn, int finterid)
        {
            var ds = conn.GetCustList(sqlconn, finterid);
            comCust.DataSource = ds.Tables[0];
            comCust.DisplayMember = "FName";           //设置显示值
            comCust.ValueMember = "FItemID";          //设置默认值内码
            //comCust.SelectedIndex = 0;             //表示默认选中是第一项
            return finterid;
        }

        /// <summary>
        /// 价格方案列表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void comList_Click(object sender, EventArgs e)
        {
            try
            {
                var sqlconn = GetConnection();
                //判断下拉列表是否有值(若没有才与数据库连接获取数据)
                if (comList.Items.Count != 0)
                {
                    comCust.Text = "";
                }
                else
                {
                    var ds = conn.GetPriceList(sqlconn);
                    comList.DataSource = ds.Tables[0];
                    comList.DisplayMember = "FName";     //设置显示值
                    comList.ValueMember = "FInterID";   //设置默认值内码
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 客户列表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void comCust_Click(object sender, EventArgs e)
        {
            try
            {
                var sqlconn = GetConnection();
                if (comList.Items.Count == 0) throw (new Exception("价格方案必须选择"));

                var dv = (DataRowView)comList.Items[comList.SelectedIndex];
                var finterid = Convert.ToInt32(dv["FInterID"]);

                //记录次此所选择的Finterid(当下一次再选择价格方案时若与上一回的finterid一致时,即不用与数据库连接读取数据)
                if (finterid != 1 || finterid != 3)
                {
                    //若为第一次获取客户列表就直接读取记录
                    if (comCust.Items.Count == 0)
                    {
                        LastInterid = GetCustToList(sqlconn, finterid);
                    }
                    //若客户列表内已有值就先判断再读取(作用:避免重复与数据库交互)
                    else
                    {
                        if (LastInterid != finterid)
                        {
                            LastInterid = GetCustToList(sqlconn, finterid);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 打开EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                var sqlcon = GetConnection();
                var openFileDialog = new OpenFileDialog { Filter = "Xlsx文件|*.xlsx" };
                if (openFileDialog.ShowDialog() != DialogResult.OK) return;
                var strPath = openFileDialog.FileName;
                var exc = new Excel();
                var dt = exc.OpenExcel(strPath, sqlcon);
                gvdtl.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导入EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btnImport_Click(object sender, EventArgs e)
        {
            int fCustId;
            string fcustName;
            var sqlcon = GetConnection();
            var exec = new Excel();
            var dt = new DataTable();
            var result = new FrmErrorResult();

            try
            {
                if (gvdtl.Rows.Count == 0) throw new Exception("没有EXCEL内容,请导入后再继续");

                //获取价格方案所选择的值
                var listdv = (DataRowView)comList.Items[comList.SelectedIndex];
                var fInterId = Convert.ToInt32(listdv["FInterID"]);
                var fPriceName = listdv["FName"].ToString();

                //获取客户所选择的值

                if (fInterId == 1 || fInterId == 3)
                {
                    fCustId = 0;
                    fcustName = "";
                }
                else
                {
                    var custdv = (DataRowView)comCust.Items[comCust.SelectedIndex];
                    fCustId = Convert.ToInt32(custdv["FItemID"]);
                    fcustName = custdv["FName"].ToString();
                }

                var dbName = rdjin.Checked ? "晶创" : "江苏";

                var clickMessage = string.Format("您所选择的信息为:\n 帐套:{0}\n 价格方案:{1}\n 客户:{2}\n 是否继续?", dbName, fPriceName, fcustName);


                if (MessageBox.Show(clickMessage, "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    dt = (DataTable)gvdtl.DataSource;

                    //获取验证从DataGridView里的DataTable能更新的值
                    var canImportdt = exec.VaildCanImportTable(sqlcon, fInterId, fCustId, dt);
                    //获取验证从DataGridView里的DataTable能更新的值
                    var cannotImportdt = exec.ErrorTable;

                    //若验证的结果与DataGridView一致,即表示全部能更新
                    if (dt.Rows.Count == canImportdt.Rows.Count)
                    {
                        exec.ImportExcel(sqlcon, fInterId, fCustId, canImportdt);
                        MessageBox.Show("已成功更新,请到K3系统对应的单据进行查阅", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    //若不能更新的记录与原DataTable一致的话
                    else if (cannotImportdt.Rows.Count == dt.Rows.Count)
                    {
                        var errormessage = "抱歉地通知您,Excel里所有记录都不能成功导入\n 请整理数据后再进行更新";
                        MessageBox.Show(errormessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    //若不是全部能更新的话就执行此(将可以更新的放到ImportExcel方法执行,将不能更新传到frmErrorResult窗体中显示)
                    else if (canImportdt.Rows.Count > cannotImportdt.Rows.Count || canImportdt.Rows.Count == cannotImportdt.Rows.Count || canImportdt.Rows.Count < cannotImportdt.Rows.Count)
                    {

                        exec.ImportExcel(sqlcon, fInterId, fCustId, canImportdt);
                        var message = "已成为更新一部份信息,不能更新的信息请按确定进行查阅";
                        MessageBox.Show(message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        result.LoadErrorRecord(cannotImportdt);
                        result.ShowDialog();
                    }
                }
                comList.Text = "";
                comCust.Text = "";
                //清空原来DataGridView内的内容(无论成功与否都会执行)
                var dtclear = (DataTable)gvdtl.DataSource;
                dtclear.Rows.Clear();
                gvdtl.DataSource = dtclear;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// 退出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btnExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否退出?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                Close();
        }
    }
}
