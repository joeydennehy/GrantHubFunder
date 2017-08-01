using System;
using System.Collections.Generic;
using System.Linq;
using MySql.Data.MySqlClient;
using System.Data;

namespace DataProvider.MySQL
{
	public class ParameterSet
	{
		private readonly Dictionary<string, MySqlParameter> parameters;

		public ParameterSet()
		{
			parameters = new Dictionary<string, MySqlParameter>();
		}

		public MySqlParameter[] Parameters
		{
			get
			{
				return parameters.Values.ToArray();
			}
		}

		public void Add(DbType dataType, string paramId, object paramValue)
		{
			string innerParam = paramId.Replace("@", "");
			string formattedParam = string.Format("@{0}", innerParam);

			MySqlParameter param = new MySqlParameter
			{
				DbType = dataType,
				Direction = ParameterDirection.Input,
				ParameterName = formattedParam,
				Value = paramValue ?? DBNull.Value,
			};

			parameters.Add(innerParam, param);
		}
	}
}
