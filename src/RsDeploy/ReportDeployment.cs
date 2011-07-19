using System;
using System.IO;
using System.Linq;
using Microsoft.SqlServer.ReportingServices2010;

namespace RsDeploy
{
	/// <summary>
	/// 
	/// </summary>
	public class ReportDeployment 
	{
		private readonly ReportingService2010 _rs;

		public ReportDeployment()
		{
			_rs = new ReportingService2010 {Credentials = Credentials, Url = Url};
		}
		public ReportDeployment(ReportingService2010 rs)
		{
			_rs = rs;
		}

		private string _url = "http://localhost:80/ReportServer/ReportService2010.asmx";
		
		///<example>http://[YourServer]/ReportServer/ReportService2010.asmx</example>
		public string Url
		{
			get { return _url; }
			set { _url = value; }
		}

		public string FileSystemPath { get; set; }

		private System.Net.ICredentials _credentials = System.Net.CredentialCache.DefaultCredentials;
		public System.Net.ICredentials Credentials
		{
			get{ return _credentials; }
			set { _credentials = value; }
		}

		/// <summary>
		/// This routine deletes and recreates a folder on a remote SSRS server
		/// </summary>
		/// <param name="reportingServerFolderName">Typically this is a / for the root folder</param>
		/// <param name="folderName">Name of the folder to create</param>
		public void DeleteAndRecreateFolder(string reportingServerFolderName, string folderName)
		{
			reportingServerFolderName = NormalizeReportingServerFolderName(reportingServerFolderName);
			
			var remoteFolderExists = _rs.ListChildren(reportingServerFolderName, false).ToList<CatalogItem>().Where(a => a.Name == folderName).Any();

			if(remoteFolderExists)
			{
				_rs.DeleteItem(reportingServerFolderName + folderName);
			}
			var folder = _rs.CreateFolder(folderName, reportingServerFolderName, null);
		}	

		/// <summary>
		/// Routine deploys a report from local file system to remote SSRS server using WebServices
		/// </summary>
		/// <param name="fileSystemPath">fully qualified file system path</param>
		/// <param name="reportingServerFolderName">reporting server folder to deploy to relative to root</param>
		/// <param name="reportName">Text name of report</param>
		public void PublishReport(string fileSystemPath, string reportingServerFolderName,string reportName)
		{
			fileSystemPath = NormalizeFileSystemPath(fileSystemPath);
			reportingServerFolderName = NormalizeReportingServerFolderName(reportingServerFolderName);

			var reportDefinition = ReadReportDefinitionFromFileSystem(fileSystemPath, reportName);
			DeployReportToReportingServer(reportingServerFolderName, reportName, reportDefinition);
		}

		private DataSource [] GetServerDataSources(string dataSourceReferenceName,string dataSourceName)
		{
			var dataSources = new DataSource[1];
			var ds = new DataSource();
			ds.Item = _rs.GetDataSourceContents("/Data Sources/"+dataSourceReferenceName);
			ds.Name = dataSourceName;
			dataSources[0] = ds;

			return dataSources;
		}
		public void SetDataSourceForReport(string reportingServerFolderName, string reportName, string dataSourceReferenceName)
		{
			reportingServerFolderName = NormalizeReportingServerFolderName(reportingServerFolderName);

			try
			{
				var dataSources = _rs.GetItemDataSources(string.Format("{0}/{1}", reportingServerFolderName, reportName));
			
				if(dataSources.Length != 1)
				{
					throw new Exception(String.Format("Only reports with a single datasource are supported. This report has {0} datasources",dataSources.Length));
				}
			
				if(dataSources[0].Item.GetType() == typeof(InvalidDataSourceReference))
				{
					//TODO if report has multiple datasources this isn't gonna work
					_rs.SetItemDataSources(String.Format("{0}/{1}", reportingServerFolderName, reportName), GetServerDataSources(dataSourceReferenceName,dataSources[0].Name));
				}
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex);
			}
		}

		private void DeployReportToReportingServer(string reportingServerFolderName, string reportName, byte[] definition)
		{
			try
			{
				Warning[] warnings = null;
				
				var report = _rs.CreateCatalogItem("Report", reportName, reportingServerFolderName, true, definition, null, out warnings);
				
				if ((warnings != null))
				{
					foreach (var warning in warnings)
					{
						Console.WriteLine(warning.Message);
					}
				}
				else
				{
					Console.WriteLine("Report: {0} published successfully with no warnings", reportName);
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
			}
		}

		private static byte[] ReadReportDefinitionFromFileSystem(string fileSystemPath, string reportName)
		{
			Byte[] definition = null;

			try
			{
				var stream = File.OpenRead(String.Format("{0}{1}.rdl", fileSystemPath, reportName));
				definition = new Byte[stream.Length + 1];
				stream.Read(definition, 0, Convert.ToInt32(stream.Length));
				stream.Close();
			}
			catch (IOException e)
			{
				System.Console.WriteLine(e.Message);
			}
			return definition;
		}

		/// <summary>
		/// Checks to see if FileSystemPath ends with \.  If not present, it is added
		/// </summary>
		/// <param name="fileSystemPath"></param>
		/// <returns></returns>
		private static string NormalizeFileSystemPath(string fileSystemPath)
		{
			if (!fileSystemPath.EndsWith(@"\"))
			{
				fileSystemPath += "\\";
			}
			return fileSystemPath;
		}

		/// <summary>
		/// Checks to see if the reporting server folder name begins with a /.  If not present, it is added
		/// </summary>
		private static string NormalizeReportingServerFolderName(string reportingServerFolderName)
		{
			if(reportingServerFolderName[0] != '/')
			{
				reportingServerFolderName = String.Format("/{0}", reportingServerFolderName);
			}
			return reportingServerFolderName;
		}
	}
}
