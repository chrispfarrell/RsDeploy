using System.IO;
using RsDeploy;

namespace ConsoleRunner
{
	class Program
	{
		/// <summary>
		/// Local file path where reports are located
		/// </summary>
		private const string FileSystemPath = @"C:\Projects\SomeProject\SomeProject.Reports\";

		/// <summary>
		/// Remote SSRS Folder to deploy to
		/// </summary>
		private const string SSRSFolderName = "MyReports";

		/// <summary>
		/// Remote folder to deploy to.  Use / for root
		/// </summary>
		private const string SSRSFolderParentPath = "/";

		/// <summary>
		/// Data source to use for reports
		/// </summary>
		private const string DataSourceName = "My Datasource";

		static void Main(string[] args)
		{
			var reportDeployment = new ReportDeployment();

			reportDeployment.DeleteAndRecreateFolder(SSRSFolderParentPath, SSRSFolderName);

			var di = new DirectoryInfo(FileSystemPath);
			var reports = di.GetFiles("*.rdl");


			foreach(var report in reports)
			{
				var reportName = Path.GetFileNameWithoutExtension(report.Name);

				reportDeployment.PublishReport(FileSystemPath, SSRSFolderName, reportName);
				reportDeployment.SetDataSourceForReport(SSRSFolderName, reportName, DataSourceName);
			}
		}
	}
}
