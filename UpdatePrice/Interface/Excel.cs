using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace UpdatePrice.Interface
{
    public class Excel
    {
        #region  查询物料是否存在

        private string _SearItem = @"
                                        
                                       SELECT COUNT(*) as tcount
                                       from dbo.t_ICItem 
                                       WHERE FItemID ={0}

                                     ";

        #endregion

        #region   临时表

        private string _SearPriceList = @"
                                            
                                            SELECT TOP 1 A.FInterID,A.FRelatedID,A.FItemID,A.FBegDate,A.FEndDate,A.FPrice,A.FCheckerID,A.FChecked,A.FCheckDate,A.FMainterID,A.FMaintDate
                                            FROM dbo.ICPrcPlyEntry A                                            

                                         ";

        #endregion

        #region   更新语句

        private string _UpdatePrice = @"
                                          UPDATE ICprc SET Icprc.FBegDate=@FBegDate,Icprc.FEndDate=@FEndDate,Icprc.FPrice=@FPrice,Icprc.FCheckerID=16394,Icprc.FChecked=1,Icprc.FCheckDate=@FCheckDate,Icprc.FMainterID=16394,Icprc.FMaintDate=@FCheckDate
                                          FROM ICPrcPlyEntry ICprc
                                          WHERE ICprc.FInterID=@FInterID
                                          AND ICprc.FRelatedID=@FCustID
                                          AND ICprc.FItemID=@FItemID  
                                        ";

        #endregion

        #region   更新前先检查是否在价目表中存在

        private string _SearErrorRecord = @"
                                                
                                               SELECT  count(*) as tcount
                                               FROM ICPrcPlyEntry icprc 
                                               WHERE icprc.FInterID={0}
                                               AND icprc.FRelatedID={1}
                                               AND icprc.FItemID={2};

                                            ";
        #endregion

        private DataTable _errorTable;

        public DataTable ErrorTable
        {
            get { return _errorTable; }
        }

        /// <summary>
        /// 打开EXCEL
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="sqlcon"></param>
        /// <returns></returns>
        public DataTable OpenExcel(string filename, SqlConnection sqlcon)
        {
            var importExcelDt = new DataTable();
            var dt = new DataTable();

            try
            {
                //var pubs = ConfigurationManager.ConnectionStrings["OleDbConStr"];  //读取配置文件
                //var conSplit = pubs.ConnectionString.Split(';');
                //var strcon = conSplit[0] + ";" + string.Format(conSplit[1], filename) + ";" + conSplit[2] + ";" + conSplit[3];

                //var con = new OleDbConnection(strcon);               //建立连接
                //const string strSql = "select * from [Sheet1$]";     //表名的写法也应注意不同，对应的excel表为sheet1，在这里要在其后加美元符号$，并用中括号
                //// var cmd = new OleDbCommand(strSql, con);          //建立要执行的命令
                //var da = new OleDbDataAdapter(strSql, con);          //建立数据适配器
                //da.Fill(ds);

                //change date:2018-07-03 使用NPOI技术进行导入EXCEL至DATATABLE
                importExcelDt = OpenExcelToDataTable(filename);

                //将从EXCEL过来的记录集为空的行清除
                // dt = RemoveEmptyRows(ds.Tables[0]);
                dt = RemoveEmptyRows(importExcelDt);

                //判断EXCEL里的物料是否存在
                var vaildResult = VaildExcelItem(dt, sqlcon);
                if (vaildResult["Code"].ToString() != "0")
                {
                    MessageBox.Show(vaildResult["Code"].ToString(), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //dt = ds.Tables[0];
                    dt.Rows.Clear();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return dt;
        }

        /// <summary>
        /// 读取EXCEL内容到DATATABLE内
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private DataTable OpenExcelToDataTable(string filename)
        {
            IWorkbook wk;
            var dt = new DataTable();
            using (var fsRead = File.OpenRead(filename))
            {
                wk = new XSSFWorkbook(fsRead);
                //获取第一个sheet
                var sheet = wk.GetSheetAt(0);
                //获取第一行
                var hearRow = sheet.GetRow(0);
                //创建列标题
                for (int i = hearRow.FirstCellNum; i < hearRow.Cells.Count; i++)
                {
                    var dataColumn = new DataColumn();
                    switch (i)
                    {
                        case 0:
                            dataColumn.ColumnName = "物料代码";
                            break;
                        case 1:
                            dataColumn.ColumnName = "物料名称";
                            break;
                        case 2:
                            dataColumn.ColumnName = "规格";
                            break;
                        case 3:
                            dataColumn.ColumnName = "单价";
                            break;
                        case 4:
                            dataColumn.ColumnName = "生效日期";
                            break;
                        case 5:
                            dataColumn.ColumnName = "失效日期";
                            break;
                    }
                    dt.Columns.Add(dataColumn);
                }

                //创建完标题后,开始从第二行起读取对应列的值
                for (int r = 1; r <= sheet.LastRowNum; r++)
                {
                    bool result = false;
                    var dr = dt.NewRow();
                    //获取当前行
                    var row = sheet.GetRow(r);
                    //读取每列
                    for (int j = 0; j < row.Cells.Count; j++)
                    {
                        //循环获取行中的单元格
                        var cell = row.GetCell(j);
                        //循环获取行中的单元格的值
                        //dr[j] = j == 4 || j == 5 ? cell.DateCellValue.ToString() : cell.ToString();
                        dr[j] = GetCellValue(cell);
                        //全为空就不取
                        if (dr[j].ToString() != "")
                        {
                            result = true;
                        }
                    }
                    if (result == true)
                    {
                        //把每行增加到DataTable
                        dt.Rows.Add(dr);
                    }

                }
            }
            return dt;
        }

        //检查单元格的值
        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank: //空数据类型 这里类型注意一下，不同版本NPOI大小写可能不一样,有的版本是Blank（首字母大写)
                    return string.Empty;
                case CellType.Boolean: //bool类型
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric: //数字类型
                    if (HSSFDateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        return cell.DateCellValue.ToString();
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString();
                    }
                case CellType.Unknown: //无法识别类型
                default: //默认类型                    
                    return cell.ToString();//
                case CellType.String: //string 类型
                    return cell.StringCellValue;
                case CellType.Formula: //带公式类型
                    try
                    {
                        var e = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }

        /// <summary>
        /// 验证导入的物料是否在数据库中存在
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sqlcon"></param>
        /// <returns></returns>
        protected Hashtable VaildExcelItem(DataTable dt, SqlConnection sqlcon)
        {
            var result = new Hashtable();
            result["Code"] = "0";

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                sqlcon.Open();
                var sqlcommand = new SqlCommand(string.Format(_SearItem, dt.Rows[i][0]), sqlcon);
                var sqlcom = sqlcommand.ExecuteReader();
                sqlcom.Read();
                var tcount = sqlcom.GetInt32(0);
                sqlcon.Close();

                if (tcount != 0) continue;
                result["Code"] = string.Format("检验到物料代码:{0} 在数据库中不存在,\n请修改后再继续", dt.Rows[i][0]);
                break;
            }
            return result;
        }

        /// <summary>
        /// 将从EXCEL导入的DATATABLE的空白行清空
        /// </summary>
        /// <param name="dt"></param>
        protected DataTable RemoveEmptyRows(DataTable dt)
        {
            var removeList = new List<DataRow>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool isNull = true;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    //将不为空的行标记为False
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                    {
                        isNull = false;
                    }
                }
                //将整行都为空白的记录
                if (isNull)
                {
                    removeList.Add(dt.Rows[i]);
                }
            }

            //将整理出来的所有空白行通过循环进行删除
            for (int i = 0; i < removeList.Count; i++)
            {
                dt.Rows.Remove(removeList[i]);
            }
            return dt;
        }

        /// <summary>
        /// 验证能成功更新的记录并生成一个新的DataTable
        /// </summary>
        /// <param name="sqlconn"></param>
        /// <param name="finterid"></param>
        /// <param name="fCustId"></param>
        /// <param name="resourceTable"></param>
        /// <returns></returns>
        public DataTable VaildCanImportTable(SqlConnection sqlconn, int finterid, int fCustId, DataTable resourceTable)
        {
            var dt = new DataTable();
            var failddt = new DataTable();

            //对DT创建自定义列(包括列名及数据类型) 接收能进行更新
            dt.Columns.Add("物料代码", Type.GetType("System.Int32"));        //物料ID FItemID
            dt.Columns.Add("物料名称", Type.GetType("System.String"));      //物料名称 FName
            dt.Columns.Add("规格", Type.GetType("System.String"));         //规格     FModel
            dt.Columns.Add("单价", Type.GetType("System.Decimal"));       //单价 FPrice 
            dt.Columns.Add("生效日期", Type.GetType("System.DateTime")); //生效日期 FBegDate 
            dt.Columns.Add("失效日期", Type.GetType("System.DateTime"));//失效日期 FEndDate 

            //接收不能更新的信息
            failddt.Columns.Add("物料代码", Type.GetType("System.Int32"));        //物料ID FItemID
            failddt.Columns.Add("物料名称", Type.GetType("System.String"));      //物料名称 FName
            failddt.Columns.Add("规格", Type.GetType("System.String"));         //规格     FModel
            failddt.Columns.Add("单价", Type.GetType("System.Decimal"));       //单价 FPrice 
            failddt.Columns.Add("生效日期", Type.GetType("System.DateTime")); //生效日期 FBegDate 
            failddt.Columns.Add("失效日期", Type.GetType("System.DateTime"));//失效日期 FEndDate 


            try
            {
                for (int i = 0; i <= resourceTable.Rows.Count - 1; i++)
                {
                    sqlconn.Open();
                    var sqlcommand = new SqlCommand(string.Format(_SearErrorRecord, finterid, fCustId, resourceTable.Rows[i][0]), sqlconn);
                    var sqlcom = sqlcommand.ExecuteReader();
                    sqlcom.Read();
                    var tcount = sqlcom.GetInt32(0);
                    if (tcount != 0)
                    {
                        var newrow = dt.NewRow();
                        newrow["物料代码"] = resourceTable.Rows[i][0];
                        newrow["物料名称"] = resourceTable.Rows[i][1];
                        newrow["规格"] = resourceTable.Rows[i][2];
                        newrow["单价"] = resourceTable.Rows[i][3];
                        newrow["生效日期"] = resourceTable.Rows[i][4];
                        newrow["失效日期"] = resourceTable.Rows[i][5];
                        dt.Rows.Add(newrow);
                    }
                    else
                    {
                        var newrow = failddt.NewRow();
                        newrow["物料代码"] = resourceTable.Rows[i][0];
                        newrow["物料名称"] = resourceTable.Rows[i][1];
                        newrow["规格"] = resourceTable.Rows[i][2];
                        newrow["单价"] = resourceTable.Rows[i][3];
                        newrow["生效日期"] = resourceTable.Rows[i][4];
                        newrow["失效日期"] = resourceTable.Rows[i][5];
                        failddt.Rows.Add(newrow);
                    }
                    sqlconn.Close();
                }
                _errorTable = failddt;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return dt;
        }

        /// <summary>
        /// 导入EXCEL至DATATABLE
        /// </summary>
        /// <param name="sqlconn"></param>
        /// <param name="fCustId"></param>
        /// <param name="dtTable"></param>
        /// <param name="finterid"></param>
        /// <returns></returns>
        public void ImportExcel(SqlConnection sqlconn, int finterid, int fCustId, DataTable dtTable)
        {
            var sqladpter = new SqlDataAdapter();
            var ds = new DataSet();

            try
            {
                using (sqladpter.SelectCommand = new SqlCommand(_SearPriceList, sqlconn))
                {
                    //将临时表插入Dataset
                    sqladpter.Fill(ds);

                    //建立更新相关设置
                    sqladpter.UpdateCommand = new SqlCommand(_UpdatePrice, sqlconn);

                    sqladpter.UpdateCommand.Parameters.Add("@FInterID", SqlDbType.Int, 8, "FInterID");                 //价目方案ID
                    sqladpter.UpdateCommand.Parameters.Add("@FCustID", SqlDbType.Int, 8, "FRelatedID");                //客户ID
                    sqladpter.UpdateCommand.Parameters.Add("@FItemID", SqlDbType.Int, 8, "FItemID");                   //物料ID
                    sqladpter.UpdateCommand.Parameters.Add("@FBegDate", SqlDbType.DateTime, 8, "FBegDate");            //生效日期
                    sqladpter.UpdateCommand.Parameters.Add("@FEndDate", SqlDbType.DateTime, 8, "FEndDate");            //失效日期
                    sqladpter.UpdateCommand.Parameters.Add("@FPrice", SqlDbType.Decimal, 28, "FPrice");                //单价
                    sqladpter.UpdateCommand.Parameters.Add("@FCheckDate", SqlDbType.DateTime, 8, "FCheckDate");        //审核日期
                    sqladpter.UpdateCommand.Parameters.Add("@FMaintDate", SqlDbType.DateTime, 8, "FMaintDate");        //维护日期

                    //开始更新
                    for (int i = 0; i <= dtTable.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < 1; j++)
                        {
                            ds.Tables[0].Rows[j].BeginEdit();
                            ds.Tables[0].Rows[j]["FInterID"] = finterid;
                            ds.Tables[0].Rows[j]["FRelatedID"] = fCustId;
                            ds.Tables[0].Rows[j]["FItemID"] = dtTable.Rows[i][0];   //物料ID
                            ds.Tables[0].Rows[j]["FBegDate"] = dtTable.Rows[i][4];  //生效日期
                            ds.Tables[0].Rows[j]["FEndDate"] = dtTable.Rows[i][5];  //失效日期
                            ds.Tables[0].Rows[j]["FPrice"] = dtTable.Rows[i][3];    //单价
                            ds.Tables[0].Rows[j]["FCheckDate"] = DateTime.Now;      //审核日期
                            ds.Tables[0].Rows[j]["FMaintDate"] = DateTime.Now;      //维护日期
                            ds.Tables[0].Rows[j].EndEdit();
                        }
                        //在保存前将ITEMID记录(用于当出现错误时的提示)
                        //itemid = Convert.ToInt32(ds.Tables[0].Rows[0]["FItemID"]);
                        sqladpter.Update(ds.Tables[0]);
                    }

                    //完成更新后将相关内容清空
                    ds.Tables[0].Clear();
                    sqladpter.Dispose();
                    ds.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// 导出EXCEL
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public Hashtable ExportExcel(string filename, DataTable dt)
        {
            var result = new Hashtable();

            try
            {
                //声明一个WorkBook(XSSFWorkbook类是用于导入导出07版本EXCEL的 而HSSFWorkbook类是用于导入导出03版本的EXCEL)
                var xssfWorkbook = new XSSFWorkbook();

                //为WorkBook创建work(创建工作表)
                var sheet = xssfWorkbook.CreateSheet("Sheet1");
                //创建第一行并对第一行内各列值进行设置
                var row = sheet.CreateRow(0);

                for (int l = 0; l < dt.Columns.Count; l++)
                {
                    //设置列宽
                    sheet.SetColumnWidth(l, (int)((20 + 0.72) * 256));

                    switch (l)
                    {
                        case 0:
                            row.CreateCell(l).SetCellValue("物料代码");
                            break;
                        case 1:
                            row.CreateCell(l).SetCellValue("物料名称");
                            break;
                        case 2:
                            row.CreateCell(l).SetCellValue("规格");
                            break;
                        case 3:
                            row.CreateCell(l).SetCellValue("单价");
                            break;
                        case 4:
                            row.CreateCell(l).SetCellValue("生效日期");
                            break;
                        case 5:
                            row.CreateCell(l).SetCellValue("失效日期");
                            break;
                    }
                }

                //创建每一行及对每行内的列进行赋值
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    //从第二行开始,因为第一行已作标题使用
                    row = sheet.CreateRow(j + 1);
                    //row.CreateCell(0).SetCellValue(j + 1);

                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        row.CreateCell(k).SetCellValue(dt.Rows[j][k].ToString());
                        // row.CreateCell(k + 1).SetCellValue(dt.Rows[i * 10000 + j][k].ToString());
                    }
                }
                //写入数据
                var file = new FileStream(filename, FileMode.Create);
                xssfWorkbook.Write(file);
                file.Close();
                result["Code"] = 0;
            }
            catch (Exception ex)
            {
                result["Code"] = ex.Message;
            }

            return result;
        }
    }
}
