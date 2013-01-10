using System;
using System.Data;
using System.ServiceProcess;
using System.Text;
using Microsoft.Win32;

namespace SQLInstanceQuery
{
	class SqlServerServiceQuery
	{
		private DataTable dataTable;
		private readonly String servername = Environment.MachineName;

		private const String INSTANCESBASEKEY = @"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL";
		private const String INSTANCEBASEKEY = @"SOFTWARE\Microsoft\Microsoft SQL Server\";
		private const String EXPRESS = "Express";

		public const String MSDTC = "MSDTC";
		public const String SQL_SERVER = "SQL Server";
		public const String SQL_AGENT = "SQL Agent";
		public const String FTSEARCH = "Full Text Search";
		public const String REPORT = "Reporting Services";
		public const String ANALYSIS = "Analysis Services";
		public const String SSIS = "SSIS(2005) Service";
		public const String SSIS2008 = "SSIS(2008) Service";
		public const String BROWSER = "SQL Browser";

		public SqlServerServiceQuery()
        {
            BuildAndLoadTable();
        }

		public string InstanceList()
        {
            var stringBuilder = new StringBuilder();
			var dataView = new DataView(SelectDistinct("InstanceList", dataTable, "InstanceName")) {Sort = "InstanceName"};

			foreach (DataRowView drv in dataView)
            {
                stringBuilder.Append(drv[0] + ",");
            }
            return stringBuilder.ToString().TrimEnd(",".ToCharArray());
        }

        public DataView Instances()
        {
        	return new DataView(SelectDistinct("InstanceList", dataTable, "InstanceName")) {Sort = "InstanceName"};
        }

        public DataView Services(String instancename)
        {
        	return new DataView(dataTable)
        	         	{RowFilter = string.Format("InstanceName = '{0}'", instancename), Sort = "DisplayName ASC"};
        }

        public String ServiceToDisplay(String servicename)
        {
            var svc = servicename.ToUpper();

            if (svc == "MSDTC")
                return MSDTC;
            if (svc == "MSSQLSERVER" || svc.StartsWith("MSSQL$"))
                return SQL_SERVER;
            if (svc == "MSFTESQL" || svc.StartsWith("MSFTESQL$") || svc == "MSSEARCH" || svc == "MSSQLFDLAUNCHER" || svc.StartsWith("MSSQLFDLAUNCHER$"))
                return FTSEARCH;
            if (svc == "SQLSERVERAGENT" || svc.StartsWith("SQLAGENT$"))
                return SQL_AGENT;
            if (svc == "REPORTSERVER" || svc.StartsWith("REPORTSERVER$"))
                return REPORT;
            if (svc == "MSSQLSERVEROLAPSERVICE" || svc.StartsWith("MSOLAP$"))
                return ANALYSIS;
            if (svc.StartsWith("MSDTSSERVER"))
            {
                switch (svc)
                {
                    case "MSDTSSERVER100":
                        return SSIS2008;
                    default:
                        return SSIS;
                }
            }
            return svc == "SQLBROWSER" ? BROWSER : SQL_SERVER;
            // default
        }

		
        private static bool IsRelaventService(String servicename)
        {
        	var svc = servicename.ToUpper();

        	return svc == "MSSQLSERVER" || svc.StartsWith("MSSQL$") ||
        	       (svc == "SQLSERVERAGENT" || svc.StartsWith("SQLAGENT$") ||
        	        (svc == "MSDTC" || (svc.StartsWith("MSDTSSERVER") ||
        	                            (svc.StartsWith("MSFTESQL") || svc == "MSSEARCH" ||
        	                             (svc.StartsWith("REPORTSERVER") ||
        	                              (svc == "SQLBROWSER" ||
        	                               (svc.StartsWith("MSOLAP") ||
        	                                svc == "MSSQLSERVEROLAPSERVICE" ||
        	                                (svc.StartsWith("MSSQLFDLAUNCHER") ||
        	                                 svc == "MSSQLFDLAUNCHER"))))))));
        }

		private void GetServices()
        {
			try
            {
            	var services = ServiceController.GetServices();

            	foreach (var t in services)
            	{
            		AddServiceToList(t.ServiceName);
            	}
            }
            catch (Exception err)
            {
            	throw new Exception("Error on get services list, " + err.Message);
            }
        }

        private void AddServiceToList(string servicename)
        {
            if (IsRelaventService(servicename))
            {
                dataTable.Rows.Add(new Object[] { servicename, ServiceToInstance(servicename), ServiceToDisplay(servicename) });
            }
        }

        private void BuildAndLoadTable()
        {
            dataTable = new DataTable("ServiceList");
        	var dataType = Type.GetType("System.String");
        	if (dataType == null) throw new ArgumentNullException(string.Format("Error on system type"));
        	var dcServiceName = new DataColumn("ServiceName", dataType);
            var dcInstanceName = new DataColumn("InstanceName", dataType);
            var dcDisplayName = new DataColumn("DisplayName", dataType);
            dataTable.Columns.Add(dcServiceName);
            dataTable.Columns.Add(dcInstanceName);
            dataTable.Columns.Add(dcDisplayName);

            var key = new DataColumn[1];
            key[0] = dcServiceName;
            dataTable.PrimaryKey = key;

            GetServices();
            Filter2008ExpressAgent();
        }

        private void Filter2008ExpressAgent()
        {
        	try
            {
                var registryKey = Registry.LocalMachine.OpenSubKey(INSTANCESBASEKEY);
            	if (registryKey != null)
            	{
            		var instances = registryKey.GetValueNames();

            		foreach (var ins in instances)
            		{
            			var openSubKey = Registry.LocalMachine.OpenSubKey(INSTANCEBASEKEY + registryKey.GetValue(ins) + @"\Setup");
            			if (openSubKey == null) continue;
            			var edition = openSubKey.GetValue("Edition").ToString();
            			var version = openSubKey.GetValue("Version").ToString();
            			openSubKey.Close();

            			if (String.IsNullOrEmpty(edition) || String.IsNullOrEmpty(version)) continue;
            			if (version.StartsWith("10.") && edition.StartsWith(EXPRESS, StringComparison.InvariantCultureIgnoreCase))
            				RemoveSqlAgent(ins);
            		}
            	}
            	if (registryKey != null) registryKey.Close();
            }
            catch (Exception err)
            {
				throw new Exception("Filter 2008 Express Agent, " + err.Message);
            }
        }

        private void RemoveSqlAgent(String instance)
        {
            var service = "SQLAgent$" + instance;

            var row = dataTable.Rows.Find(service);

            if (null != row)
            {
                dataTable.Rows.Remove(row);
            }
        }

        private String ServiceToInstance(String svc)
        {
        	if (svc.Contains("$"))
            {
                return servername + "\\" + svc.Split(new[] { '$' })[1];
            }
        	return servername;
        }

		private bool ColumnEqual(object a, object b)
        {
            if (a == DBNull.Value && b == DBNull.Value) //  both are DBNull.Value
                return true;
            if (a == DBNull.Value || b == DBNull.Value) //  only one is DBNull.Value
                return false;
            return (a.Equals(b));  // value type standard comparison
        }

        private DataTable SelectDistinct(string tableName, DataTable sourceTable, string fieldName)
        {
            var dt = new DataTable(tableName);
            dt.Columns.Add(fieldName, sourceTable.Columns[fieldName].DataType);

            object lastValue = null;
            foreach (var dr in sourceTable.Select("", fieldName))
            {
            	if (lastValue != null && (ColumnEqual(lastValue, dr[fieldName]))) continue;
            	lastValue = dr[fieldName];
            	dt.Rows.Add(new[] { lastValue });
            }
            return dt;
        }

	}
}
