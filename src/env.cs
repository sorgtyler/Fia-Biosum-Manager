using System;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for env.
	/// </summary>
	public class env
	{
		public env()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		public string strAppDir
		{
			get
			{
				string strAppDir="";
				strAppDir = 
					System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase.ToString());
				int intAt = strAppDir.IndexOf("file:\\",0, strAppDir.Length);

				if (intAt >=0) 
				{
					int intLen = strAppDir.Length - 5;
					strAppDir = strAppDir.Substring(6);
				}
				return strAppDir;
			}
			
		}
		public string strWinDir
		{
			get 
			{
				return System.Environment.SystemDirectory.ToString();
			}
		}
		public string strTempDir
		{
			get
			{
				string strTempDir ="";
				strTempDir = System.Environment.GetEnvironmentVariable("tmp");
				if (strTempDir.Length == 0) 
				{
					strTempDir = System.Environment.GetEnvironmentVariable("temp");
				}
				return strTempDir;
			}
		}
		public string strOperatingSystem
		{
			get
			{
				return System.Environment.OSVersion.ToString();
			}
		}
		public string strHomeDrive
		{
			get
			{
				return System.Environment.GetEnvironmentVariable("homedrive");
			}
		}
		public string strHomePath
		{
			get
			{
				return System.Environment.GetEnvironmentVariable("homepath");
			}
		}
		public string strApplicationDataDirectory
		{
			get
			{
				return System.Environment.GetEnvironmentVariable("appdata");
					
			}
		}
	}
}
