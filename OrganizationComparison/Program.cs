using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;

namespace OrganizationComparison
{
	internal class OrganizationComparison
	{
		private static void Main(string[] args)
		{
			IEnumerable<string> organization1 = File.ReadLines(args[0]);
			IEnumerable<string> organization2 = File.ReadLines(args[1]);

			List<Organizations> duplicates = new List<Organizations>();


	
				foreach (var line in organization2)
				{
					string duplicate = IsDuplicateOrganization(organization1.ToList(), line);
					if (duplicate != null)
					{
						duplicates.Add(new Organizations
						{
							Name = duplicate
						});
					}
				}

			StringBuilder sb = new StringBuilder();


			foreach (Organizations dedupe in duplicates)
			{
				sb.AppendLine(dedupe.Name);
			}

			File.WriteAllText("duplicates.csv", sb.ToString());

			
		}


		private static string IsDuplicateOrganization(List<string> org1, string compare)
		{
			foreach (string dedupe in org1)
			{
				if (string.IsNullOrWhiteSpace(compare)) continue;
				if (dedupe.ToLower() == compare.ToLower())
					return string.Format("Name: {0} - {1}:{2}", dedupe, compare, "100");
				var s = dedupe.ToLower();
				var t = compare.ToLower();

				if (string.IsNullOrEmpty(s))
				{
					if (string.IsNullOrEmpty(t))
						return null;
					return null;
				}

				if (string.IsNullOrEmpty(t))
					return null;

				var n = s.Length;
				var m = t.Length;
				int[,] d = new int[n + 1, m + 1];

				// initialize the top and right of the table to 0, 1, 2, ...
				for (var i = 0; i <= n; d[i, 0] = i++) ;
				for (var j = 1; j <= m; d[0, j] = j++) ;

				for (var i = 1; i <= n; i++)
					for (var j = 1; j <= m; j++)
					{
						var cost = t[j - 1] == s[i - 1] ? 0 : 1;
						var min1 = d[i - 1, j] + 1;
						var min2 = d[i, j - 1] + 1;
						var min3 = d[i - 1, j - 1] + cost;
						d[i, j] = Math.Min(Math.Min(min1, min2), min3);
					}

				var distance = d[n, m];
				var bigger = Math.Max(s.Length, t.Length);
				var percent = (int)((bigger - distance) / (double)bigger * 100);

				if (percent >= 85)
					return string.Format("Name: {0} - {1}:{2}", dedupe, compare, percent);
			}
			return null;
		}

		private struct Organizations
		{
			public string Name;
		}
	}
}