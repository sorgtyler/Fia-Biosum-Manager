using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
namespace FIA_Biosum_Manager
{	
	/*********************************************************************************************************
	 **DbFileItem                          
	 *********************************************************************************************************/
	public class DbFileItem
	{
		
		public TableLinkItem_Collection m_oTableLinkItemCollection1=new TableLinkItem_Collection();
        public TableItem_Collection m_oTableItemCollection1 = new TableItem_Collection();

        public DbFileItem()
        {
            _TableItemCollection = m_oTableItemCollection1;
            _TableLinkItemCollection = m_oTableLinkItemCollection1;
        }
		
		private string _strDbFileName="";
		public string DbFileName
		{
			get {return _strDbFileName;}
			set {_strDbFileName=value;}
		}
        private string _strTableType="";
        public string TableType
        {
            get { return _strTableType; }
            set { _strTableType = value; }
        }
		
		private string _strDirectory="";
		public string Directory
		{
			get {return _strDirectory;}
			set {_strDirectory=value;}
		}
        public string FullPath
        {
            get {return Directory + "\\" + DbFileName;}
        }
		
        private TableItem_Collection _TableItemCollection;
        public  TableItem_Collection TableCollection
		{
		    get {return _TableItemCollection;}
		    set {_TableItemCollection=value;}
		}
        private TableLinkItem_Collection _TableLinkItemCollection;
        public  TableLinkItem_Collection TableLinkCollection
		{
		    get {return _TableLinkItemCollection;}
		    set {_TableLinkItemCollection=value;}
		}
        private System.Data.OleDb.OleDbConnection _oConn = null;
        public System.Data.OleDb.OleDbConnection Connection
        {
            get { return _oConn; }
            set { _oConn = value; }
        }
        private void InstantiateConnectionObject(ado_data_access p_oAdo)
        {

            if (_oConn == null)
            {
            }
            else
            {
                p_oAdo.CloseConnection(_oConn);
                _oConn.Dispose();
                _oConn = null;
            }
            _oConn = new System.Data.OleDb.OleDbConnection();
            
        }
        public void OpenConnection(ado_data_access p_oAdo)
        {
            InstantiateConnectionObject(p_oAdo);
            p_oAdo.OpenConnection(p_oAdo.getMDBConnString(FullPath, "", ""), ref _oConn);
        }

		
	}
	public class DbFileItem_Collection : System.Collections.CollectionBase
	{
		public DbFileItem_Collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(DbFileItem p_oItem)
		{
			// vérify if object is not already in
			if (this.List.Contains(p_oItem))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(p_oItem);
 
			// return collection
			//return this;
		}
		public void Remove(int index)
		{
			// Check to see if there is a widget at the supplied index.
			if (index > Count - 1 || index < 0)
				// If no widget exists, a messagebox is shown and the operation 
				// is canColumned.
			{
				System.Windows.Forms.MessageBox.Show("Index not valid!");
			}
			else
			{
				List.RemoveAt(index); 
			}
		}
		public DbFileItem Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (DbFileItem) List[Index];
		}

	}
	
	public class TableLinkItem
  {
		
				
		private string _strDbFileName="";
		public string DbFileName
		{
			get {return _strDbFileName;}
			set {_strDbFileName=value;}
		}
		
		private string _strDirectory="";
		public string Directory
		{
			get {return _strDirectory;}
			set {_strDirectory=value;}
		}
         public string FullPath
        {
            get {return Directory + "\\" + DbFileName;}
        }
		private string _strTableName="";
		public string TableName
		{
			get {return _strTableName;}
			set {_strTableName=value;}
		}
       

		private string _strLinkedTableName="";
		public string LinkedTableName
		{
			get {return _strLinkedTableName;}
			set {_strLinkedTableName=value;}
		}

        private int _intTypeId = -1;
        public int TypeId
        {
            get { return _intTypeId; }
            set { _intTypeId = value; }
        }
        private string _strTypeDesc = "";
        public string TypeDescription
        {
            get { return _strTypeDesc; }
            set { _strTypeDesc = value; }
        }
        private string _strFVSOutputTableName = "";
        public string FVSOutputTableName
        {
            get { return _strFVSOutputTableName; }
            set { _strFVSOutputTableName = value; }
        }
        private bool _bFVSOutputTable = false;
        public bool FVSOutputTable
        {
            get { return _bFVSOutputTable; }
            set { _bFVSOutputTable = value; }
        }
        private bool _bFVSOutputSeqNumMatrixTable = false;
        public bool FVSOutputSeqNumMatrixTable
        {
            get { return _bFVSOutputSeqNumMatrixTable; }
            set { _bFVSOutputSeqNumMatrixTable = value; }
        }

		
		
		
	}
	public class TableLinkItem_Collection : System.Collections.CollectionBase
	{
		public TableLinkItem_Collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(TableLinkItem p_oItem)
		{
			// vérify if object is not already in
			if (this.List.Contains(p_oItem))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(p_oItem);
 
			// return collection
			//return this;
		}
		public void Remove(int index)
		{
			// Check to see if there is a widget at the supplied index.
			if (index > Count - 1 || index < 0)
				// If no widget exists, a messagebox is shown and the operation 
				// is canColumned.
			{
				System.Windows.Forms.MessageBox.Show("Index not valid!");
			}
			else
			{
				List.RemoveAt(index); 
			}
		}
		public TableLinkItem Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (TableLinkItem) List[Index];
		}

	}
	
	
	public class TableItem
    {
		private string _strTableName="";
		public string TableName
		{
			get {return _strTableName;}
			set {_strTableName=value;}
		}
		
	}
	public class TableItem_Collection : System.Collections.CollectionBase
	{
		public TableItem_Collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(TableItem p_oItem)
		{
			// vérify if object is not already in
			if (this.List.Contains(p_oItem))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(p_oItem);
 
			// return collection
			//return this;
		}
		public void Remove(int index)
		{
			// Check to see if there is a widget at the supplied index.
			if (index > Count - 1 || index < 0)
				// If no widget exists, a messagebox is shown and the operation 
				// is canColumned.
			{
				System.Windows.Forms.MessageBox.Show("Index not valid!");
			}
			else
			{
				List.RemoveAt(index); 
			}
		}
		public TableItem Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (TableItem) List[Index];
		}

	}
}
	