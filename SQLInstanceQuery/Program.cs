using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace SQLInstanceQuery
{
	class Program
	{
		static void Main(string[] args)
		{
			var sqlServerServiceQuery = new SqlServerServiceQuery();
			var services = sqlServerServiceQuery.Instances();
			foreach(DataRow service in services.Table.Rows)
			{
				Console.WriteLine(service["InstanceName"]);
			}
			Console.ReadLine();
		}
	}
}
