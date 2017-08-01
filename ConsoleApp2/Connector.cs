using System;
using System.ComponentModel;
using System.Configuration;
using MySql.Data.MySqlClient;

namespace DataProvider.MySQL
{
	internal static class Connector
	{
		#region Private Variables

		private const string PRIMARY_DB_CONNECTION_NAME = "MasterDatabase";
		private const string LOG_DB_CONNECTION_NAME = "LogConnection";

		private static string primaryConnectionString = string.Empty;
		private static string loggingConnectionString = string.Empty;

		#endregion

		#region Constructor

		static Connector()
		{
			Initialize();
		}

		#endregion

		private static void Initialize()
		{
			try
			{
				primaryConnectionString = ConfigurationManager.ConnectionStrings[PRIMARY_DB_CONNECTION_NAME].ConnectionString;
			}
			catch (ConfigurationErrorsException e)
			{
				throw new Exception("Unable to read primary database configuration file", e);
			}

			try
			{
				loggingConnectionString = ConfigurationManager.ConnectionStrings[LOG_DB_CONNECTION_NAME].ConnectionString;
			}
			catch (ConfigurationErrorsException e)
			{
				throw new Exception("Unable to read logging database configuration file", e);
			}
		}

		public static MySqlConnection GetConnection(DataAccessType accessType = DataAccessType.Data)
		{
			return accessType == DataAccessType.Data
				? new MySqlConnection(primaryConnectionString)
				: new MySqlConnection(loggingConnectionString);
		}
	}
}