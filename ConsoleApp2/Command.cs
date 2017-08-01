using System;
using System.Resources;
using ResourceReader = Foundations.ResourceReader;

namespace DataProvider.MySQL
{

	public class Command
	{
		private string sqlStatementId;

		public string SqlStatementId
		{
			get
			{
				return sqlStatementId;
			}
			set
			{
				if (string.IsNullOrEmpty(value))
					throw new ArgumentException("Command Statement ID cannot be null or empty.");

				sqlStatementId = value;

				SqlStatement = ResourceReader.GetSql(value);
				
				if (string.IsNullOrEmpty(SqlStatement))
				{
					throw new ArgumentException(string.Format("Command Statement ID: '{0}' could not be found.", value));
				}
			}
		}

		public string SqlStatement { get; private set; }
		public ParameterSet ParameterCollection { get; set; }
	}
}
