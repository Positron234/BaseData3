using System;
using System.Data.SqlClient;
namespace BaseData1
{
    class BaseData
    {
        //Строка подключения БД
        SqlConnection SqlConnection = new SqlConnection(@"Data Source=DESKTOP-C8L1TCD;Initial Catalog=Filarmoni;Integrated Security=True");
        public void OpenConnection()//открыть соединение
        {
            if (SqlConnection.State == System.Data.ConnectionState.Closed)
            {
                SqlConnection.Open();
            }

        }
        public void ClosedConnection()// закрыть соединение
        {
            if (SqlConnection.State == System.Data.ConnectionState.Open)
            {
                SqlConnection.Close();
            }
        }
        public SqlConnection GetConnection()//вернуть соединение
        {
            return SqlConnection;
        }
    }
}
