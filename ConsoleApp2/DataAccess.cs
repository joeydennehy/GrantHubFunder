using System;
using System.Data;
using System.Linq;
using MySql.Data.MySqlClient;

namespace DataProvider.MySQL
{

	public enum DataAccessType
	{
		Data,
		Logging
	}

	public class DataAccess
	{
		public MySqlDataReader GetReader(Command command)
		{
			MySqlConnection connection = Connector.GetConnection();

			MySqlCommand cmd = new MySqlCommand(command.SqlStatement, connection);
			cmd.CommandTimeout = 65535;
			if (command.ParameterCollection != null && command.ParameterCollection.Parameters.Any())
			{
				cmd.Parameters.AddRange(command.ParameterCollection.Parameters);
			}

			connection.Open();

			MySqlDataReader dataReader;
			try
			{
				dataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection|CommandBehavior.SequentialAccess);
			}
			catch (Exception eError)
			{
				throw new Exception(string.Format("{0}:{1}", eError.Message, eError.StackTrace));
			}
			return dataReader;
		}

		private object ExecuteScalarQuery(Command command)
		{
			MySqlConnection connection = Connector.GetConnection();

			MySqlCommand cmd = new MySqlCommand(command.SqlStatement, connection);
			cmd.CommandTimeout = 65535;
			if (command.ParameterCollection != null && command.ParameterCollection.Parameters.Any())
			{
				cmd.Parameters.AddRange(command.ParameterCollection.Parameters);
			}

			connection.Open();

			object selectedValue;

			try
			{
				selectedValue = cmd.ExecuteScalar();
			}
			catch (Exception eError)
			{
				throw new Exception(string.Format("{0}:{1}", eError.Message, eError.StackTrace));
			}
			finally
			{
				connection.Close();
			}

			return selectedValue;
		}

		public string GetStringValue(Command command)
		{
			object selectedValue = ExecuteScalarQuery(command);

			return selectedValue == null ? string.Empty : (string)selectedValue;
		}

		public int GetIntValue(Command command)
		{
			object selectedValue = ExecuteScalarQuery(command);

			return selectedValue == null ? 0 : int.Parse(selectedValue.ToString());
		}
	}
	
}

