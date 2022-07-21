using System;
using System.Data;
using System.Data.OleDb;

namespace SamuraiTools.Database
{
    public abstract class SqlDbConnection : IDisposable
    {
        protected OleDbConnection dbConnection;

        protected string correctWritePassword;

        private string connectionString;
        public string ConnectionString
        {
            get
            {
                return connectionString;
            }
            set
            {
                CloseConnection();
                connectionString = value;
            }
        }

        public SqlDbConnection()
        { }

        public SqlDbConnection(string connectionString) : this()
        {
            this.ConnectionString = connectionString;
            OpenConnection();
        }

        public virtual void OpenConnection()
        {
            dbConnection = new OleDbConnection();

            dbConnection.ConnectionString = ConnectionString;
            dbConnection.Open();
        }

        public virtual void CloseConnection()
        {
            if (dbConnection?.State == ConnectionState.Open)
            {
                dbConnection.Close();
            }

            dbConnection?.Dispose();
        }

        protected virtual bool IterateRecords(System.Action<DataRow> DoAction, string query)
        {
            bool returnValue;

            DataTable table = new DataTable();
            OleDbDataAdapter adapter = new OleDbDataAdapter(query, dbConnection);

            try
            {
                adapter.Fill(table);

                if (table.Rows.Count > 0)
                {
                    foreach (DataRow row in table.Rows)
                    {
                        DoAction(row);
                    }

                    returnValue = true;
                }
                else
                {
                    returnValue = false;
                }
                return returnValue;
            }
            finally
            {
                adapter.Dispose();
                table.Dispose();
            }
        }

        private static Random rand = new Random();
        protected virtual string GenerateRandomID(ushort length)
        {
            char[] allowableCharacters = "abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray();
            System.Text.StringBuilder returnValue = new System.Text.StringBuilder();
            ushort index;

            for (index = 0; index < length; index++)
            {
                returnValue = returnValue.Append(allowableCharacters[rand.Next(0, allowableCharacters.Length)]);
            }

            return returnValue.ToString();
        }

        protected string GenerateRandomID(ushort length, System.Predicate<string> IsUnique)
        {
            string returnValue;

            do
            {
                returnValue = GenerateRandomID(length);
            }
            while (!IsUnique(returnValue));

            return returnValue;
        }

        #region IDisposable Support
        protected bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                    CloseConnection();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~SqlDbConnection()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            //GC.SuppressFinalize(this);
        }
        #endregion
    }
}
