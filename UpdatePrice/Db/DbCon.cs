using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace UpdatePrice.Db
{
    public class DbCon
    {
        #region   查询价格方案(价格方案列表使用)

        private string _SearPriceList = @"

                                             SELECT A.FInterID,A.FName
                                             from dbo.ICPrcPly A          
                                                
                                         ";

        #endregion

        #region 查询客户(客户列表使用)

        private string _SearCustomer = @"
                                           
                                            SELECT DISTINCT B.FItemID,B.FName
                                            from dbo.ICPrcPlyEntry A
                                            INNER JOIN dbo.t_Organization B ON A.FRelatedID=B.FItemID
                                            WHERE A.FInterID={0};                                            
 
                                        ";

        #endregion


        /// <summary>
        /// 获取需要连接的帐套并返回连接信息
        /// </summary>
        /// <param name="dbString"></param>
        /// <returns></returns>
        public SqlConnection GetConnect(string dbString)
        {
            var pubs = ConfigurationManager.ConnectionStrings["Connstring"];  //读取配置文件

            var consplit = pubs.ConnectionString.Split(';');
            var strcon = consplit[0] + ";" + string.Format(consplit[1], dbString) + ";" + consplit[2] + ";" + consplit[3] + ";" + consplit[4] + ";" + consplit[5] + ";" + consplit[6] + ";" + consplit[7];

            var conn = new SqlConnection(strcon);
            return conn;
        }

        /// <summary>
        /// 获取列表信息(价格方案)
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataSet GetPriceList(SqlConnection sql)
        {
            var sqlDataAdapter = new SqlDataAdapter();
            var ds = new DataSet();

            try
            {
                sqlDataAdapter.SelectCommand = new SqlCommand(_SearPriceList, sql);
                sqlDataAdapter.Fill(ds);
            }
            catch (Exception ex)
            {
                throw (new Exception(ex.Message));
            }


            return ds;
        }

        /// <summary>
        /// 获取客户列表信息
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="fInterId"></param>
        /// <returns></returns>
        public DataSet GetCustList(SqlConnection sql,int fInterId)
        {
            var sqlDataAdapter = new SqlDataAdapter();
            var ds = new DataSet();

            try
            {
                sqlDataAdapter.SelectCommand=new SqlCommand(string.Format(_SearCustomer,fInterId),sql);
                sqlDataAdapter.Fill(ds);
            }
            catch (Exception ex)
            {
                throw (new Exception(ex.Message));
            }
            return ds;
        }
    }
}
