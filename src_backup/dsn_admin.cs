using System;
using System.Runtime.InteropServices;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Admin-Class for DSN connections.
	/// </summary>
	public class DSNAdmin
	{
		public DSNAdmin()
		{
		}

		/// <summary>
		/// Creates a User DSN.
		/// </summary>
		/// <param name="DSName">Name of the DSN</param>
		/// <param name="DBPath">Path to database</param>
		/// <returns>True if successfull</returns>
		public bool AddUserDSN(string DSName, string DBPath)
		{
			return SQLConfigDataSource((IntPtr)0, 1, "Microsoft Access Driver (*.MDB)\0", 
				"DSN=" + DSName + "\0Uid=\0pwd=\0DBQ=" + DBPath + "\0");

			//return SQLConfigDataSource((IntPtr)0, 1, "Microsoft Access Driver (*.MDB)\0", 
			//	"DSN=" + DSName + "\0Uid=Admin\0pwd=\0DBQ=" + DBPath + "\0");
		}
		
		/// <summary>
		/// Modify a User DSN.
		/// </summary>
		/// <param name="DSName">Name of the DSN</param>
		/// <param name="DBPath">Path to database</param>
		/// <returns>True if successfull</returns>
		public bool EditUserDSN(string DSName, string DBPath)
		{
			return SQLConfigDataSource((IntPtr)0, 2, "Microsoft Access Driver (*.MDB)\0", 
				"DSN=" + DSName + "\0Uid=Admin\0pwd=\0DBQ=" + DBPath + "\0");
		}

		/// <summary>
		/// Remove a User DSN.
		/// </summary>
		/// <param name="DSName">Name of the DSN</param>
		/// <returns>True if successfull</returns>
		public bool RemoveUserDSN(string DSName)
		{
			return SQLConfigDataSource((IntPtr)0, 3, "Microsoft Access Driver (*.MDB)\0", 
				"DSN=" + DSName + "\0Uid=\0pwd=\0DBQ=\0");

		}

		/// <summary>
		/// Add a System DSN.
		/// </summary>
		/// <param name="DSName">Name of the DSN</param>
		/// <param name="DBPath">Path to database</param>
		/// <returns>True if successfull</returns>
		public bool AddSystemDSN(string DSName, string DBPath)
		{
			return SQLConfigDataSource((IntPtr)0, 4, "Microsoft Access Driver (*.MDB)\0", 
				"DSN=" + DSName + "\0Uid=Admin\0pwd=\0DBQ=" + DBPath + "\0");
		}

		/// <summary>
		/// Modify a System DSN
		/// </summary>
		/// <param name="DSName">Name of DSN</param>
		/// <param name="DBPath">Path to database</param>
		/// <returns>True if successfull</returns>
		public bool EditSystemDSN(string DSName, string DBPath)
		{
			return SQLConfigDataSource((IntPtr)0, 5, "Microsoft Access Driver (*.MDB)\0", 
				"DSN=" + DSName + "\0Uid=Admin\0pwd=\0DBQ=" + DBPath + "\0");
		}

		/// <summary>
		/// Remove a System DSN
		/// </summary>
		/// <param name="DSName">Name of DSN</param>
		/// <returns>True if successfull</returns>
		public bool RemoveSystemDSN(string DSName)
		{
			return SQLConfigDataSource((IntPtr)0, 6, "Microsoft Access Driver (*.MDB)\0", 
				"DSN=" + DSName + "\0Uid=Admin\0pwd=\0DBQ=\0");
		}

		/// <summary>
		/// Win32 API Import
		/// </summary>
		[
		DllImport("ODBCCP32.dll")
		]
		private static extern bool SQLConfigDataSource(IntPtr parent, int request, string driver, string attributes);
	}

// ****************************************************************
// * Valid values for "request":                                  *
// *                                                              *
// * 1 - ODBC_ADD_DSN (use this to add a user DSN)                *
// * 2 - ODBC_CONFIG_DSN (use this to configure a user DSN)       *
// * 3 - ODBC_REMOVE_DSN (use this to remove a user DSN)          *
// * 4 - ODBC_ADD_SYS_DSN (use this to add a system DSN)          *
// * 5 - ODBC_CONFIG_SYS_DSN (use this to configure a system DSN) *
// * 6 - ODBC_REMOVE_SYS_DSN (use this to remove a system DSN)    *
// ****************************************************************
}
