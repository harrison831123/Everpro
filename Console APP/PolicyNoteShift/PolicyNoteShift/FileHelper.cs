using NLog;
using SevenZip;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PolicyNoteShift
{
    public class FileHelper
    {
        private static Logger logger = NLog.LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Map NetWork Drive建立網路磁碟機
        /// </summary>
        /// <param name="ls_diskcode">Disk Code</param>
        /// <param name="ls_message">Error Message</param>
        /// <returns>True:Success False:Failure</returns>
        public Boolean MapNetworkDrive(string ls_path, string ls_account, string ls_password, ref string ls_diskcode)
        {
            try
            {
                ls_diskcode = "A";
                string[] lsa_array = System.Environment.GetLogicalDrives();
                int li_count = 1;

                while (!(Array.IndexOf(lsa_array, ls_diskcode + ":\\") == -1 | li_count > 26))
                {
                    ls_diskcode = Convert.ToChar(Convert.ToInt32(Convert.ToChar(ls_diskcode)) + 1).ToString();
                    li_count += 1;
                }
                ls_diskcode += ":";

                using (System.Diagnostics.Process lp_process = new System.Diagnostics.Process())
                {
                    lp_process.StartInfo.UseShellExecute = false;
                    lp_process.StartInfo.FileName = "net.exe";
                    lp_process.StartInfo.CreateNoWindow = true;
                    lp_process.StartInfo.UseShellExecute = false;
                    lp_process.StartInfo.RedirectStandardError = true;
                    lp_process.StartInfo.Arguments = "use " + ls_diskcode + " " + ls_path + " /user:" + ls_account + " " + ls_password;
                    lp_process.Start();
                    lp_process.WaitForExit();
                    // This code assumes the process you are starting will terminate itself.
                    // Given that it is started without a window so you cannot terminate it
                    // on the desktop, it must terminate itself or you can do it programmatically
                    // from this application using the Kill method.
                }
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                return false;
            }
        }
        /// <summary>
        /// UnMap NetWork Drive
        /// </summary>
        /// <param name="ls_diskcode">Disk Code</param>
        /// <param name="ls_message">Error Message</param>
        /// <returns>True:Success False:Failure</returns>
        public Boolean UnmapNetworkDrive(string ls_diskcode)
        {
            try
            {

                using (System.Diagnostics.Process lp_process = new System.Diagnostics.Process())
                {
                    lp_process.StartInfo.FileName = "net.exe";
                    lp_process.StartInfo.CreateNoWindow = true;
                    lp_process.StartInfo.UseShellExecute = false;
                    lp_process.StartInfo.RedirectStandardError = true;

                    lp_process.StartInfo.Arguments = "use " + ls_diskcode + " /delete /Y";
                    lp_process.Start();
                }
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                return false;
            }
        }

        #region User Define DataTable
        /// <summary>
        /// User Import DataTable Layout Information Return Define DatatTable
        /// </summary>
        /// <param name="ldt_source">Define Layout</param>
        /// <returns>DatatTable</returns>
        public DataTable ChangeUserDefineData2Dt<T>(List<T> items)
        {
            DataTable ldt_dt = new DataTable();
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            try
            {
                foreach (T item in items)
                {
                    DataColumn ldc_dc = new DataColumn();
                    for (int i = 0; i < Props.Length; i++)
                    {
                        switch (Props[i].Name)
                        {
                            case "tablename":
                                if (ldt_dt.TableName == "")
                                {
                                    ldt_dt.TableName = Props[i].GetValue(item, null).ToString();
                                }
                                break;
                            case "colname":
                                ldc_dc.ColumnName = Props[i].GetValue(item, null).ToString();
                                break;
                            case "coltype":
                                switch (Props[i].GetValue(item, null).ToString())
                                {
                                    case "string":
                                        ldc_dc.DataType = System.Type.GetType("System.String");
                                        break;
                                    case "integer":
                                        ldc_dc.DataType = System.Type.GetType("System.Int32");
                                        break;
                                    case "boolean":
                                        ldc_dc.DataType = System.Type.GetType("System.Boolean");
                                        break;
                                    case "timespan":
                                        ldc_dc.DataType = System.Type.GetType("System.TimeSpan");
                                        break;
                                    case "datetime":
                                        ldc_dc.DataType = System.Type.GetType("System.DateTime");
                                        break;
                                    case "decimal":
                                        ldc_dc.DataType = System.Type.GetType("System.Decimal");
                                        break;
                                    default:
                                        ldc_dc.DataType = System.Type.GetType("System.String");
                                        break;
                                }
                                break;
                        }
                    }
                    ldt_dt.Columns.Add(ldc_dc);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
            return ldt_dt;
        }

        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }
        #endregion

        #region ListToDataTable || DataSet
        /// <summary>
        /// ModelListData chage to DatatTable
        /// </summary>
        /// <param name="Name">Model List</param>
        /// <returns>DatatTable</returns>
        public DataTable ListToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }

        public DataSet ToDataSet<T>(List<T> list)
        {
            Type elementType = typeof(T);
            DataSet ds = new DataSet();
            DataTable t = new DataTable();
            ds.Tables.Add(t);

			//add a column to table for each public property on T
			foreach (var propInfo in elementType.GetProperties())
			{
				Type ColType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

				t.Columns.Add(propInfo.Name, ColType);
			}

			//go through each property on T and add each value to the table
			foreach (T item in list)
            {
                DataRow row = t.NewRow();

                foreach (var propInfo in elementType.GetProperties())
                {
                    row[propInfo.Name] = propInfo.GetValue(item, null) ?? DBNull.Value;
                }

                t.Rows.Add(row);
            }

            return ds;
        }
        #endregion

        /// <summary>
        ///7z解壓檔案
        /// </summary>
        /// <param name="SourceFilePath">來源檔位置</param>
        /// /// <param name="OutputDirectory">檔案解壓位置</param>
        /// /// <param name="password">zipfilePw</param>
        public void ExtractFiles(string SourceFilePath, string OutputDirectory, string password)
        {
            //SevenZipExtractor要先指定7z.dll位置
            SevenZipExtractor.SetLibraryPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "7z.dll"));

            using (var extractor = new SevenZipExtractor(SourceFilePath, password))
            {
                for (var i = 0; i < extractor.ArchiveFileData.Count; i++)
                {
                    extractor.ExtractFiles(OutputDirectory, extractor.ArchiveFileData[i].Index);
                }
            }
        }

        /// <summary>
        ///檔案資料寫入DataTable
        /// </summary>
        /// <param name="SourceFilePath">來源檔位置</param>
        /// /// <param name="OutputDirectory">檔案解壓位置</param>
        /// /// <param name="password">zipfilePw</param>
        public DataTable ReadFileToDataTable(string ls_readfile, string fileencoding, string CompanyCode, DataTable ldt_import, string zipfilename = "")
        {
            //DataTable ldt_import = new DataTable();
            try
            {              
                if (ls_readfile.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) || ls_readfile.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    int li_colcount = ldt_import.Columns.Count;

                    using (StreamReader sr = new StreamReader(ls_readfile, System.Text.Encoding.GetEncoding(fileencoding)))
                    {
                        String ls_linedata;
                        DataRow lr_new;
                        while ((ls_linedata = sr.ReadLine()) != null)
                        {
                            string[] lsa_data = ls_linedata.Split(',');

                            lr_new = ldt_import.NewRow();
                            int i = 0;
                            for (i = 0; i < lsa_data.Length; i++)
                            {
                                if (i < li_colcount)
                                {
                                    lr_new[i] = lsa_data[i].Trim();
                                }
                            }
                            if (zipfilename != "")
                                lr_new[i] = zipfilename;
                            else
                                lr_new[i] = ls_readfile;

                            lr_new[i + 1] = CompanyCode;

                            ldt_import.Rows.Add(lr_new);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error("ReadFileToDataTable-Error：" + ls_readfile + "；" + ex.Message);
            }
            return ldt_import;
        }

    }
}
