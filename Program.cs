using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    static class Program
    {

        static HashSet<string> hsh = new HashSet<string>();

        /// <summary>
        /// Creates a SQL script that creates a table where the columns matches that of the specified DataTable.
        /// </summary>
        public static string BuildCreateTableScript(DataTable Table)
        {
            StringBuilder result = new StringBuilder();
            if (!hsh.Contains(Table.TableName))
            {
                result.AppendFormat("IF OBJECT_ID('dbo.[{0}]', 'U') IS NOT NULL  DROP TABLE dbo.[{0}];\r\n", Table.TableName);

                hsh.Add(Table.TableName);
            }
            result.AppendFormat("CREATE TABLE [{1}] ({0}   ", Environment.NewLine, Table.TableName);

            bool FirstTime = true;
            foreach (DataColumn column in Table.Columns.OfType<DataColumn>())
            {
                if (FirstTime) FirstTime = false;
                else
                    result.Append("   ,");

                result.AppendFormat("[{0}] {1} {2} {3}",
                    column.ColumnName, // 0
                    GetSQLTypeAsString(column.DataType), // 1
                    column.AllowDBNull ? "NULL" : "NOT NULL", // 2
                    Environment.NewLine // 3
                );
            }
            result.AppendFormat(") ON [PRIMARY]{0}{0}{0}", Environment.NewLine);

            // Build an ALTER TABLE script that adds keys to a table that already exists.
            if (Table.PrimaryKey.Length > 0)
                result.Append(BuildKeysScript(Table));

            return result.ToString();
        }

        /// <summary>
        /// Builds an ALTER TABLE script that adds a primary or composite key to a table that already exists.
        /// </summary>
        private static string BuildKeysScript(DataTable Table)
        {
            // Already checked by public method CreateTable. Un-comment if making the method public
            // if (Helper.IsValidDatatable(Table, IgnoreZeroRows: true)) return string.Empty;
            if (Table.PrimaryKey.Length < 1) return string.Empty;

            StringBuilder result = new StringBuilder();

            if (Table.PrimaryKey.Length == 1)
                result.AppendFormat("ALTER TABLE {1}{0}   ADD PRIMARY KEY ({2}){0}GO{0}{0}", Environment.NewLine, Table.TableName, Table.PrimaryKey[0].ColumnName);
            else
            {
                List<string> compositeKeys = Table.PrimaryKey.OfType<DataColumn>().Select(dc => dc.ColumnName).ToList();
                string keyName = compositeKeys.Aggregate((a, b) => a + b);
                string keys = compositeKeys.Aggregate((a, b) => string.Format("{0}, {1}", a, b));
                result.AppendFormat("ALTER TABLE {1}{0}ADD CONSTRAINT pk_{3} PRIMARY KEY ({2}){0}GO{0}{0}", Environment.NewLine, Table.TableName, keys, keyName);
            }

            return result.ToString();
        }

        static TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

        /// <summary>
        /// Returns the SQL data type equivalent, as a string for use in SQL script generation methods.
        /// </summary>
        private static string GetSQLTypeAsString(Type DataType)
        {
            string pcase = textInfo.ToTitleCase(DataType.Name);

            switch (pcase)
            {
                case "Boolean": return "[bit]";
                case "Char": return "[char]";
                case "SByte": return "[tinyint]";
                case "Int16": return "[smallint]";
                case "Int32": return "[int]";
                case "Int64": return "[bigint]";
                case "Byte": return "[tinyint] UNSIGNED";
                case "UInt16": return "[smallint] UNSIGNED";
                case "UInt32": return "[int] UNSIGNED";
                case "UInt64": return "[bigint] UNSIGNED";
                case "Single": return "[float]";
                case "Double": return "[float]";
                case "Decimal": return "[decimal]";
                case "DateTime": return "[datetime]";
                case "Guid": return "[uniqueidentifier]";
                case "Object": return "[variant]";
                case "String": return "[nvarchar](4000)";
                default: return "[nvarchar](MAX)";
            }
        }

        public static List<DataTable> access(string filePath)
        {
            // string baglantiCumlesi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Database1.accdb;Persist Security Info=False;";

            string strConn;
            if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".accdb")
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";;Persist Security Info=False;";
            else
                strConn = "??";
            OleDbConnection conn = new OleDbConnection(strConn);
            List<DataTable> dtt = new List<DataTable>();
            try
            {
                conn.Open();
                DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                string fn = System.IO.Path.GetFileNameWithoutExtension(filePath);


                //Looping Total Sheet of Xl File
                foreach (DataRow schemaRow in schemaTable.Rows)
                {
                    DataTable dtexcel = new DataTable();
                    string sheet = schemaRow["TABLE_NAME"].ToString();
                    if (!sheet.EndsWith("_"))
                    {
                        string query = "SELECT * FROM [" + sheet + "]";
                        OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);
                        dtexcel.Locale = CultureInfo.CurrentCulture;
                        if (dtexcel.Rows.Count > 0)
                        {
                            daexcel.Fill(0, 1000, dtexcel);
                            dtexcel.TableName = fn + '_' + sheet;
                            dtt.Add(dtexcel);
                        }

                    }
                }
            }
            finally
            {
                conn.Close();
            }
            return dtt;
        }


        private static OleDbConnection _conn = null;

        public static OleDbConnection GetOleConn(string filePath)
        {
            if (_conn != null) return _conn;

            bool hasHeaders = true;
            string HDR = hasHeaders ? "Yes" : "No";
            string strConn;
            if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
            else
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
            _conn = new OleDbConnection(strConn);
            _conn.Open();
            return _conn;
        }
        public static List<DataTable> exceldata(string filePath, bool assingle = false)
        {
            OleDbConnection conn = GetOleConn(filePath);
            List<DataTable> dtt = new List<DataTable>();
            try
            {
                DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                string fn = System.IO.Path.GetFileNameWithoutExtension(filePath);

                DataTable dtexcelSingle = new DataTable(fn);
                dtexcelSingle.Locale = CultureInfo.CurrentCulture;
                //Looping Total Sheet of Xl File
                foreach (DataRow schemaRow in schemaTable.Rows)
                {
                    //Looping a first Sheet of Xl File
                    //  schemaRow = schemaTable.Rows[0];
                    string sheet = schemaRow["TABLE_NAME"].ToString();
                    if (!sheet.EndsWith("_"))
                    {
                        string query = "SELECT * FROM [" + sheet + "]";
                        OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);

                        if (assingle)
                        {
                            daexcel.Fill(dtexcelSingle);
                        }
                        else
                        {
                            DataTable dtexcel = new DataTable();
                            dtexcel.Locale = CultureInfo.CurrentCulture;
                            daexcel.Fill(0, 1000, dtexcel);

                            if (dtexcel.Rows.Count > 0)
                            {
                                //    dtexcel.Locale = CultureInfo.CurrentCulture;
                                //  dtexcel.TableName = fn + '_' + sheet;
                                dtexcel.TableName = sheet;
                                dtt.Add(dtexcel);
                            }
                        }
                    }
                }

                if (assingle) dtt.Add(dtexcelSingle);
            }
            finally
            {
                //     conn.Close();
            }

            return dtt;
        }

        public enum Ftype
        {
            xls, csv, mdb
        }

        public static char[] sepchars = { '\t', ',', '|', ';' };

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            bool singletable = false;
            char csvSeperatorChar = '\t';
            Console.WriteLine("HolyOne csv 2 mssql importer, by Aytek Ustundag, www.tahribat.com");

            if (args.Length < 2)
            {
                Console.WriteLine("USAGE:");
                Console.WriteLine("csv2mssql.exe <ConnectionString> <Filename>");
                Console.WriteLine("EXAMPLE:");
                Console.WriteLine(@"csv2mssql.exe ""Data Source=(local);Initial Catalog=dbname;Integrated Security=SSPI"" ""data.csv""");
                Console.WriteLine(@"csv2mssql.exe ""Data Source=(local);Initial Catalog=dbname;Integrated Security=SSPI"" ""excelfile.xlsx""");

                Console.WriteLine("");
                return;
            }

            string filename = args[1];
            string connstr = args[0];

            if (args.Length >= 3)
                singletable = args[2].Equals("/single", StringComparison.InvariantCultureIgnoreCase);


            if (!System.IO.File.Exists(filename))
            {
                Console.WriteLine(@"Input file ""{0}"" not found", filename);
                return;
            }

            //string baglantiCumlesi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\Database1.accdb;Persist Security Info=False;";

            //   foreach (string filename in args)
            {


                //string filename = "x.csv";
                /*    string bulk_data_filename = "x.csv";
                    StreamReader file = new StreamReader(bulk_data_filename);
                    CsvReader csv = new CsvReader(file, true, ',');
                    SqlBulkCopy copy = new SqlBulkCopy(conn);
                    copy.DestinationTableName = System.IO.Path.GetFileNameWithoutExtension(bulk_data_filename);
                    copy.WriteToServer(csv);
                    */

                string tablename = System.IO.Path.GetFileNameWithoutExtension(filename);



                string ext = System.IO.Path.GetExtension(filename);
                List<DataTable> dts = new List<DataTable>();

                Ftype mode = Ftype.csv;
                if (ext.Equals(".xls", StringComparison.InvariantCultureIgnoreCase) || ext.Equals(".xlsx", StringComparison.InvariantCultureIgnoreCase))
                {
                    mode = Ftype.xls;
                    dts = exceldata(filename, singletable);

                }
                else
                if (ext.Equals(".accdb", StringComparison.InvariantCultureIgnoreCase) || ext.Equals(".mdb", StringComparison.InvariantCultureIgnoreCase))
                {
                    mode = Ftype.mdb;
                    dts = access(filename);

                }
                else

                {
                    //csv mode 

                    using (var csvStreamReader = new StreamReader(filename))
                    using (LumenWorks.Framework.IO.Csv.CsvReader csvReader = new LumenWorks.Framework.IO.Csv.CsvReader(csvStreamReader, true))
                    {
                        DataTable dt = new DataTable(tablename);
                        int tmpcnt = 0;
                        string myline = "";

                        while (tmpcnt < 100)
                        {
                            myline = csvStreamReader.ReadLine();
                            if (myline == null) break;
                            if (tmpcnt == 0)
                            {//first line

                                Dictionary<char, int> charcounter = new Dictionary<char, int>();
                                foreach (char cx in sepchars)
                                {
                                    int chcount = myline.Count(f => f == cx);
                                    charcounter.Add(cx, chcount);
                                }
                                charcounter = charcounter.OrderByDescending(x => x.Value).ToDictionary(x => x.Key, x => x.Value);

                                csvSeperatorChar = charcounter.First().Key;
                                Console.WriteLine("Resolving seperator char:" + ((csvSeperatorChar == '\t') ? "<TAB>" : (csvSeperatorChar.ToString())));

                            }
                            tmpcnt++;

                            string[] cells = myline.Replace(@"""", "").Split(new char[] { csvSeperatorChar });
                            while (dt.Columns.Count < cells.Length)
                            {
                                dt.Columns.Add(cells[dt.Columns.Count]);
                            }
                            dt.Rows.Add(cells);
                        }

                        //  dt.Load(csvReader);
                        dts.Add(dt);
                    }

                }

                if (dts.Count == 0)
                {
                    Console.WriteLine("No data table found to import");
                    return;
                }
                foreach (DataTable dt in dts)
                {
                    string str = BuildCreateTableScript(dt);
                    SqlConnection conn = new SqlConnection(connstr);

                    conn.Open();

                    try
                    {
                        using (SqlCommand cmd = new SqlCommand(str, conn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception exx)
                    {
                        Console.WriteLine("\tWarning:" + exx.Message + " ,Appending...");
                    }

                    DataTable dtexcelSingle = new DataTable(tablename);

                    SqlTransaction transaction = conn.BeginTransaction();
                    try
                    {
                        int batchsize = 0;
                        Console.WriteLine("Importing table {0}", dt.TableName);

                        if (mode == Ftype.csv)
                            using (StreamReader file = new StreamReader(filename))
                            {
                                using (LumenWorks.Framework.IO.Csv.CsvReader csv = new LumenWorks.Framework.IO.Csv.CsvReader(file, true, csvSeperatorChar))
                                // using (CsvReader csv = new CsvReader(file, true, csvSeperatorChar,'\0','\0','#', ValueTrimmingOptions.None))
                                {
                                    csv.SkipEmptyLines = true;
                                    csv.SupportsMultiline = true;
                                    csv.MissingFieldAction = LumenWorks.Framework.IO.Csv.MissingFieldAction.ReplaceByNull;
                                    //    csv.DefaultParseErrorAction = ParseErrorAction.AdvanceToNextLine;

                                    SqlBulkCopy copy = new SqlBulkCopy(conn, SqlBulkCopyOptions.KeepIdentity, transaction);
                                    //  SqlBulkCopy copy = new SqlBulkCopy(connstr, SqlBulkCopyOptions.KeepIdentity );
                                    copy.BulkCopyTimeout = 9999999;
                                    copy.DestinationTableName = tablename;
                                    copy.WriteToServer(csv);
                                    batchsize = copy.RowsCopiedCount();
                                    transaction.Commit();
                                }
                            }
                        else
                        {
                            string sheet = dt.TableName;
                            if (sheet.EndsWith("_"))
                            {
                                continue;
                            }
                            OleDbConnection oconn = GetOleConn(filename);
                            try
                            {
                                {
                                    {
                                        string query = "SELECT * FROM [" + sheet + "]";
                                        OleDbDataAdapter daexcel = new OleDbDataAdapter(query, oconn);
                                        using (OleDbCommand cmd = new OleDbCommand(query, oconn))
                                        {
                                            using (OleDbDataReader rdr = cmd.ExecuteReader())
                                            {
                                                SqlBulkCopy copy = new SqlBulkCopy(conn, SqlBulkCopyOptions.KeepIdentity, transaction);
                                                //  SqlBulkCopy copy = new SqlBulkCopy(connstr, SqlBulkCopyOptions.KeepIdentity );
                                                copy.BulkCopyTimeout = 9999999;
                                                copy.DestinationTableName = dt.TableName;
                                                copy.WriteToServer(rdr);
                                                batchsize = copy.RowsCopiedCount();
                                                transaction.Commit();

                                            }
                                        }
                                    }
                                }
                            }
                            finally
                            {
                            }

                        }
                        Console.WriteLine("Finished inserting {0} records.", batchsize);
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        Console.WriteLine("Err:" + ex.Message);
                    }
                    finally
                    {
                        conn.Close();
                    }

                }
            }
            //    Console.ReadKey();
        }
    }
}
