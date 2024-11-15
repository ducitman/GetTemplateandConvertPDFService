﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace GetTemplateandConvertPDFService.Models
{
    public class DB
    {
        HttpClient _httpClient = new HttpClient();
        public async Task<DataTable> DB_table(string Application_Name, string Service_name, string DB_instance, string query)
        {
            HttpClient _httpClient = new HttpClient();
            string endpoint = "http://10.118.11.112/Common/API/API/DB/query?DB=Oracle";
            DataTable dt = new DataTable();
            try
            {
                var data = new { Application_Name = Application_Name, Service_name = Service_name, DB_instance = DB_instance, DB_query = query };
                var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                HttpResponseMessage response = await _httpClient.PostAsync(endpoint, content);
                string responseBody = await response.Content.ReadAsStringAsync();
                dt = JsonConvert.DeserializeObject<DataTable>(responseBody);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }
        public async Task<DataTable> DB_table_SQL(string Application_Name, string Service_name, string DB_instance, string query)
        {
            HttpClient _httpClient = new HttpClient();
            string endpoint = "http://10.118.11.112/Common/API/API/DB/query?DB=SQLServer";
            DataTable dt = new DataTable();
            try
            {
                var data = new { Application_Name = Application_Name, Service_name = Service_name, DB_instance = DB_instance, DB_query = query };
                var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                HttpResponseMessage response = await _httpClient.PostAsync(endpoint, content);
                string responseBody = await response.Content.ReadAsStringAsync();
                dt = JsonConvert.DeserializeObject<DataTable>(responseBody);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }
        public async Task<string> SQLServer_NonQuery(string Application_Name, string Service_name, string DB_instance, string query)
        {
            HttpClient _httpClient = new HttpClient();
            string endpoint = "http://10.118.11.112/Common/API/API/DB/NonQuery?DB=SQLServer";

            try
            {
                var data = new { Application_Name = Application_Name, Service_name = Service_name, DB_instance = DB_instance, DB_query = query };
                var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                HttpResponseMessage response = await _httpClient.PostAsync(endpoint, content);
                var contents = await response.Content.ReadAsStringAsync();
                return contents;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public async Task<string> ORCL_NonQuery(string Application_Name, string Service_name, string DB_instance, string query)
        {
            HttpClient _httpClient = new HttpClient();
            string endpoint = "http://10.118.11.112/Common/API/API/DB/NonQuery?DB=Oracle";
            // string endpoint = "http://10.118.11.112/Common/API/API/DB/NonQuery?DB=SQLServer";

            try
            {
                var data = new { Application_Name = Application_Name, Service_name = Service_name, DB_instance = DB_instance, DB_query = query };
                var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                HttpResponseMessage response = await _httpClient.PostAsync(endpoint, content);
                var contents = await response.Content.ReadAsStringAsync();
                return contents;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
