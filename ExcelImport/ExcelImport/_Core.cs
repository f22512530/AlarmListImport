using Dapper;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelImport
{
    public class _Core
    {
        string strConn = string.Empty;
        public _Core()
        {
            strConn = "server=127.0.0.1;uid=hartford;pwd=51321357;database=ds;SslMode=None;";
            m_connection = new MySql.Data.MySqlClient.MySqlConnection(strConn);

            if (m_connection.State != ConnectionState.Open)
                m_connection.Open();
        }

        private bool IsConnected
        {
            get { return m_connection_p.State == ConnectionState.Open; }
        }

        private MySql.Data.MySqlClient.MySqlConnection m_connection_p;
        private MySql.Data.MySqlClient.MySqlConnection m_connection
        {
            get
            {
                if (IsConnected && m_connection_p.Ping())
                    return m_connection_p;
                else
                {
                    m_connection_p = new MySql.Data.MySqlClient.MySqlConnection(strConn);
                    m_connection_p.Open();
                    return m_connection_p;
                }
            }
            set
            {
                m_connection_p = value;
            }
        }

        public bool CreateTable(string tableName)
        {
            string command = @"CREATE TABLE `hartnet`.`alarmlist` (
                               `No` INT NOT NULL AUTO_INCREMENT,
                               `alarmNum` VARCHAR(45) NULL,
                               `alarmMsg_cht` VARCHAR(200) NULL,
                               `alarmMsg_eng` VARCHAR(200) NULL,
                               `IsSend` VARCHAR(10) NOT NULL,
                               `version` VARCHAR(45) NOT NULL,
                               PRIMARY KEY (`No`),
                               UNIQUE INDEX `No_UNIQUE` (`No` ASC));";
            return SetData(command, new { });
        }

        public bool IsExist(String tableName)
        {
            String sqlcommand = @"show tables from `hartnet` where `tables_in_hartnet` = '" + tableName + @"';";
            bool IsCorrect = false;
            IsCorrect = GetData(sqlcommand, new { }, out List<string> dt);

            //資料正確擷取，而且資料列的數列為1筆
            return IsCorrect && dt.Count == 1;
        }

        public bool InsertRow(string alarmNum, string alarmMsg_cht, string alarmMsg_eng, bool IsSend, string version)
        {
            bool IsCorrect = false;
            try
            {
                string command = @"INSERT INTO `hartnet`.`alarmlist` (`alarmNum`, `alarmMsg_cht`, `alarmMsg_eng`,
                                                                      `IsSend`, `version`) VALUES (
                                                                      @alarmNum, @alarmMsg_cht, @alarmMsg_eng,
                                                                      @IsSend, @version);";
                IsCorrect = SetData(command, new
                            {
                                alarmNum,
                                alarmMsg_cht,
                                alarmMsg_eng,
                                IsSend,
                                version
                            });
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return IsCorrect;
        }

        public bool GetData<T>(string command, object data, out List<T> result)
        {
            bool IsCorrect = false;
            result = null;
            try
            {
                if (IsConnected)
                {
                    result = m_connection.Query<T>(command, data).ToList();
                    if (result.Count > 0)
                        IsCorrect = true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return IsCorrect;
        }

        public bool SetData(string command, object data)
        {
            bool IsCorrect = false;
            try
            {
                if (IsConnected)
                {
                    int s = m_connection.Execute(command, data);
                    IsCorrect = s > 0 ? true : false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return IsCorrect;
        }
    }
}
