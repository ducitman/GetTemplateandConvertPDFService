using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetTemplateandConvertPDFService.Models
{
    public class DBLOGHelper
    {
        public static bool Error = false;
        public static string ErrorMessage = "";
        static SqlConnection _SqlConnection = new SqlConnection();

        /*---------------------------- Các phương thức kết nối tới cơ sở dữ liệu ----------------------------*/

        #region Các phương thức kết nối tới cơ sở dữ liệu
        /// <summary>
        /// Xâu kết nối
        /// </summary>
        public static string ConnectionString
        {
            get { return _SqlConnection.ConnectionString; }
            set
            {
                Error = false;
                try
                {
                    _SqlConnection = new SqlConnection(value);
                    _SqlConnection.Open();
                    _SqlConnection.Close();
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
        public static void SetConnection(string DataSource, string InitialCatalog, string UserId, string Password)
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = DataSource;
            builder.InitialCatalog = InitialCatalog;
            builder.UserID = UserId;
            builder.Password = Password;
            ConnectionString = builder.ConnectionString;
        }
        #endregion
        /*---------------------------------------------------------------------------------------------------*/


        /*------------------ Các phương thức trả về một DataTable từ một câu lệnh truy vấn ------------------*/

        #region Các phương thức trả về một DataTable từ một câu lệnh truy vấn
        /// <summary>
        /// Phương thức thực thi câu lệnh truy vấn trả về kết quả là một DataTable
        /// </summary>
        /// <param name="TruyVan">Câu lệnh truy vấn</param>
        /// <returns>DataTable chứa kết quả truy vấn</returns>
        public static DataTable GetData(string Query)
        {
            Error = false;
            DataTable tbl = new DataTable();
            try
            {
                if (_SqlConnection.State == ConnectionState.Closed) _SqlConnection.Open();
                SqlDataAdapter adp = new SqlDataAdapter(Query, _SqlConnection);
                adp.Fill(tbl);
                adp.Dispose();
            }
            catch (Exception ex)
            {
                Error = true;
                ErrorMessage = ex.Message;
            }
            finally
            {
                if (_SqlConnection.State != ConnectionState.Closed)
                    _SqlConnection.Close();
            }
            return tbl;
        }
        /// <summary>
        /// Phương thức thực thi câu lệnh truy vấn trả về kết quả là true/false
        /// </summary>
        /// <param name="TruyVan">Câu lệnh truy vấn</param>
        /// <returns>Trả kết quả truy vấn true/false</returns>
        public static bool CheckData(string Query)
        {
            Error = false;
            bool result = false;
            DataTable tbl = new DataTable();
            try
            {
                if (_SqlConnection.State == ConnectionState.Closed) _SqlConnection.Open();
                SqlDataAdapter adp = new SqlDataAdapter(Query, _SqlConnection);
                adp.Fill(tbl);
                adp.Dispose();
                result = tbl.Rows.Count > 0;
            }
            catch (Exception ex)
            {
                Error = true;
                ErrorMessage = ex.Message;
            }
            finally
            {
                if (_SqlConnection.State != ConnectionState.Closed)
                    _SqlConnection.Close();
            }
            return result;
        }
        #endregion
        /*---------------------------------------------------------------------------------------------------*/


        /*---------------------------------- Các phương thức xử lý dữ liệu ----------------------------------*/

        #region Các phương thức xử lý dữ liệu
        /// <summary>
        /// Phương thức cho phép thực thi một câu lệnh sửa đổi dữ liệu
        /// </summary>
        /// <param name="Query">Câu lệnh Insert, Update hoặc Delete</param>
        public static void ExecuteQuery(string Query)
        {
            Error = false;
            try
            {
                if (_SqlConnection.State == ConnectionState.Closed) _SqlConnection.Open();
                SqlCommand command = new SqlCommand(Query);
                command.Connection = _SqlConnection;
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Error = true;
                ErrorMessage = ex.Message;
            }
            finally
            {
                if (_SqlConnection.State != ConnectionState.Closed)
                    _SqlConnection.Close();
            }
        }

        /// <summary>
        /// Phương thức thực thi một câu lệnh truy vấn sử dụng TRANSACTION
        /// </summary>
        /// <param name="Query">Câu lệnh Insert, Update hoặc Delete</param>
        public static void ExecuteQueryUsingTran(string Query)
        {
            ExecuteQuery("BEGIN TRANSACTION;" + Query + "; COMMIT;");
        }

        /// <summary>
        /// Phương thức thực thi một câu lệnh truy lấy dữ liệu
        /// </summary>
        /// <param name="Query">Câu lệnh truy vấn</param>
        /// <returns>Số bản ghi của bảng</returns>
        public static bool CheckExist(string Query)
        {
            DataTable tbl = new DataTable();
            tbl = GetData(Query);
            return (tbl.Rows.Count > 0);
        }
        #endregion
    }
}
