using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetTemplateandConvertPDFService.Models
{
    public class DBBusiness
    {
        public void SetConnectionString()
        {
            string oraSPECConnectionString = ConfigurationManager.ConnectionStrings["SPECDB"].ConnectionString;
            string oraBOSSConnectionString = ConfigurationManager.ConnectionStrings["BOSSDB"].ConnectionString;
            string oraLOGConnectionString = ConfigurationManager.ConnectionStrings["LOGDB"].ConnectionString;
            
            DBHelperOracle.ConnectionString = oraSPECConnectionString;
            
            DBLOGHelper.ConnectionString = oraLOGConnectionString;
        }
        public DataTable GetOrclData(string query)
        {
            DataTable tbl = new DataTable();
            tbl = DBHelperOracle.GetData(query);
            return tbl;
        }
        public DataTable GetSQLData(string query)
        {
            DataTable tbl = new DataTable();
            tbl = DBLOGHelper.GetData(query);
            return tbl;
        }
        public bool ExcuteSQLQuery(string query)
        {
            try
            {
                DBLOGHelper.ExecuteQuery(query);
            }
            catch(Exception ex)
            {
                throw ex;
            }
            return DBLOGHelper.Error;
        }
    }
}
