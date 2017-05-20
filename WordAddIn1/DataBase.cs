using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WordPresent
{
    class Format
    {
        public float width;
        public float height;
        public object style;
    }

    class Data
    {
        public string type;
        public string format;
        public string data;
    }

    class DataBase
    {
        #region static member
        public static DataBase instance;

        static string connectString =   @"Provider=Microsoft.ACE.OLEDB.12.0;" + 
                                        @"Extended Properties='Excel 12.0; HDR=NO; IMEX=1;';" +
                                        @"Data source= {0}";
        static string dataString = "SELECT * FROM [{0}]";

        static DataBase()
        {
            DataBase.instance = new DataBase();
        }

        #endregion

        public string filePath;
        public string path;
        public string fileName;

        public string selectTableName;

        public string[] tableNames;

        private DataTable optionTable = new DataTable();
        private DataTable formatTable = new DataTable();
        private DataTable dataTable = new DataTable();

        public Dictionary<string, string> optionDictionary = new Dictionary<string, string>();
        public Dictionary<string, Format> formatDictionary = new Dictionary<string, Format>();
        public List<Data> dataList = new List<Data>();

        public void ConnectToAccess()
        {
            System.Data.OleDb.OleDbConnection con = new
                System.Data.OleDb.OleDbConnection();
            // TODO: Modify the connection string and include any
            // additional required properties for your database.
            con.ConnectionString = string.Format(connectString,filePath);
            try
            {
                con.Open();
                // Insert code to process data.
                DataTable schema = con.GetSchema("Tables");
            }
            catch (Exception)
            {
                MessageBox.Show("Failed to connect to data source");
            }
            finally
            {
                con.Close();
            }
        }

        public string[] GetTableName()
        {

            System.Data.OleDb.OleDbConnection con = new
                System.Data.OleDb.OleDbConnection();
            // TODO: Modify the connection string and include any
            // additional required properties for your database.
            con.ConnectionString = string.Format(connectString, filePath);
            try
            {
                con.Open();

                // Insert code to process data.
                DataTable schema = con.GetSchema("Tables");
                List<string> ss = new List<string>();
                foreach (DataRow row in schema.Rows)
                {
                    ss.Add(row["Table_Name"] as string);
                }
                tableNames = ss.ToArray();
                return tableNames;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }

            return null;
        }

        // try get column index in a data row
        int TryGetColumn(DataRow row,string key)
        {

            for (int i = 0; row[i] != null; i++)
            {
                if (row[i].ToString() == key)
                    return i;
            }

            return -1;
        }

        public bool GetOption()
        {
            bool ok = GetTable("Option$", optionTable);

            // get column for name,data
            int name = TryGetColumn(optionTable.Rows[0], "name");
            if(name == -1)
            {
                return false;
            }
            int data = TryGetColumn(optionTable.Rows[0], "data");
            if(data == -1)
            {
                return false;
            }

            optionDictionary.Clear();
            try
            {
                for (int i = 1; i < optionTable.Rows.Count;i++ )
                {
                    DataRow row = optionTable.Rows[i];
                    optionDictionary.Add(row[name] as string, row[data] as string);
                }
            }
            catch (Exception)
            {

            }

            return ok;
        }

        public bool GetFormat()
        {
            bool ok =  GetTable("Format$", formatTable);
            int format = TryGetColumn(formatTable.Rows[0], "format");
            int width = TryGetColumn(formatTable.Rows[0], "width");
            int height = TryGetColumn(formatTable.Rows[0], "height");
            int style = TryGetColumn(formatTable.Rows[0], "style");
            
            formatDictionary.Clear();
            try
            {
                for (int i = 1; i < formatTable.Rows.Count; i++)
                {
                    DataRow row = formatTable.Rows[i];
                    formatDictionary.Add(row[format] as string, new Format
                    {
                        width = (row[width] as string) == null ? 0 : float.Parse(row[width] as string),
                        height = (row[height] as string) == null ? 0 : float.Parse(row[height] as string),
                        style = row[style] as string,
                    });
                }
            }
            catch (Exception)
            {
            }

            return ok;
        }

        public bool GetDataTable()
        {
            bool ok = GetTable(selectTableName, dataTable);

            int type = TryGetColumn(dataTable.Rows[0], "type");
            int format = TryGetColumn(dataTable.Rows[0], "format");
            int data = TryGetColumn(dataTable.Rows[0], "data");

            dataList.Clear();
            try
            {
                for (int i = 1; i < dataTable.Rows.Count; i++)
                {
                    DataRow row = dataTable.Rows[i];
                    dataList.Add(new Data
                        {
                            type = row[type] as string,
                            format = row[format] as string,
                            data = row[data] as string,
                        });
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return ok;
        }
        internal bool GetTable(string name, DataTable table)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter(
                    string.Format(dataString, name),
                    string.Format(connectString, filePath));

            try
            {
                // Insert code to process data.
                table.Clear();
                adapter.Fill(table);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
            finally
            {
                if (adapter != null)
                    adapter.Dispose();
            }

            return true;
        }
    }
}
