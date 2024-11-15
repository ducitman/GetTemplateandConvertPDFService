using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetTemplateandConvertPDFService.Models
{
    public class DBHelperOracle
    {
        public static bool Error = false;
        public static string ErrorMessage = "";
        static OracleConnection _oracleConnection = new OracleConnection();

        /*---------------------------- Các phương thức kết nối tới cơ sở dữ liệu ----------------------------*/

        #region Các phương thức kết nối tới cơ sở dữ liệu 
        ///<summary>
        /// Xâu kết nối
        /// </summary>
        /// 

        public static string ConnectionString
        {
            get
            {
                return _oracleConnection.ConnectionString;
            }
            set
            {
                Error = false;
                try
                {
                    _oracleConnection = new OracleConnection(value);
                    _oracleConnection.Open();
                    _oracleConnection.Close();
                }
                catch (Exception ex)
                {
                    Error = true;
                    ErrorMessage = ex.Message;
                }
            }
        }

        /// <summary>
        /// Thiết lập kết nối sử dụng Window Authentication Mode
        /// </summary>
        /// <param name="DataSource">Server Name</param>
        /// <param name="InitialCatalog">Database Name</param>
        /// <param name="ConnectTimeout">Timeout</param>
        ///User Id = BOSS; Password = BOSS; Data Source = (DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = 10.118.11.26)(PORT = 1521))(CONNECT_DATA =(SERVER = DEDICATED)(SERVICE_NAME = BOSSDB)))
        public static void SetConnection(string DataSource, string UserId, string Password)
        {
            OracleConnectionStringBuilder builder = new OracleConnectionStringBuilder();
            builder.DataSource = DataSource;
            builder.UserID = UserId;
            builder.Password = Password;
            ConnectionString = builder.ConnectionString;
        }
        #endregion
        /*---------------------------------------------------------------------------------------------------*/


        /*------------------ Các phương thức trả về một DataTable từ một câu lệnh truy vấn ------------------*/
        #region Các phương thức trả về một dataTable từ 1 lệnh truy vấn
        /// <summary>
        /// Phương thức thực thi câu lệnh truy vấn trả về kết quả là một DataTable
        /// </summary>
        /// <param name="TruyVan">Câu lệnh truy vấn</param>
        /// <returns>DataTable chứa kết quả truy vấn</returns>
        /// 
        public static DataTable GetData(string Query)
        {
            Error = false;
            DataTable tbl = new DataTable();
            try
            {
                if (_oracleConnection.State == ConnectionState.Closed) _oracleConnection.Open();
                OracleDataAdapter adt = new OracleDataAdapter(Query, _oracleConnection);
                adt.Fill(tbl);
                adt.Dispose();
            }
            catch (Exception ex)
            {
                Error = true;
                ErrorMessage = ex.Message;
            }
            finally
            {
                if (_oracleConnection.State != ConnectionState.Closed) _oracleConnection.Close();
            }
            return tbl;
        }
        #endregion
        /*---------------------------------------------------------------------------------------------------*/


        /*---------------------------------- Các phương thức xử lý dữ liệu ----------------------------------*/
        #region Các phương thức xử lý câu lệnh
        ///<summary>
        ///Phương thức thực thi sửa đổi dữ liệu
        /// </summary>
        /// <param name="Query"> Insert, Update, Delete Query </param>
        /// 
        public static void ExecuteQuery(string Query)
        {
            Error = false;
            try
            {
                if (_oracleConnection.State == ConnectionState.Closed) _oracleConnection.Open();
                OracleCommand cmd = new OracleCommand(Query);
                cmd.Connection = _oracleConnection;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Error = true;
                ErrorMessage = ex.Message;
            }
            finally
            {
                if (_oracleConnection.State != ConnectionState.Closed) _oracleConnection.Close();
            }
        }

        public static void ExecuteQueryUsingTran(string[] Query)
        {
            Error = false;
            try
            {
                if (_oracleConnection.State == ConnectionState.Closed) _oracleConnection.Open();
                OracleCommand cmd = _oracleConnection.CreateCommand();
                OracleTransaction transaction;
                transaction = _oracleConnection.BeginTransaction(IsolationLevel.ReadCommitted);
                cmd.Transaction = transaction;
                try
                {
                    foreach (string item in Query)
                    {
                        if (!String.IsNullOrEmpty(item))
                        {
                            cmd.CommandText = item;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    Error = true;
                    ErrorMessage = ex.Message;
                }
            }
            catch (Exception ex)
            {
                Error = true;
                ErrorMessage = ex.Message;
            }
            finally
            {
                if (_oracleConnection.State != ConnectionState.Closed) _oracleConnection.Close();
            }
        }
        /// <summary>
        /// Phương thức thực thi một câu lệnh truy vấn sử dụng TRANSACTION
        /// </summary>
        /// <param name="Query">Câu lệnh Insert, Update hoặc Delete</param>
        /// 
        public static void ExecuteQueryUsingTran(string Query)
        {
            ExecuteQuery(Query);
        }

        /// <summary>
        /// Phương thức thực thi một câu lệnh truy lấy dữ liệu
        /// </summary>
        /// <param name="Query">Câu lệnh truy vấn</param>
        /// <returns>Số bản ghi của bảng</returns>
        /// 
        public static bool CheckExist(string Query)
        {
            DataTable tbl = new DataTable();
            tbl = GetData(Query);
            return (tbl.Rows.Count > 0);
        }
        #endregion
    }
}
