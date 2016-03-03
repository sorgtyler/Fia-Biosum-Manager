using System;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for ListViewAlternateColors.
	/// </summary>
	public class ListViewAlternateBackgroundColors
	{
		System.Windows.Forms.ListView _oListView=null;
		System.Drawing.Color _oBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
		System.Drawing.Color _oAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
		System.Drawing.Color _oForegroundColor = frmMain.g_oGridViewRowForegroundColor;
		System.Drawing.Color _oAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
		System.Drawing.Color _oSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
		System.Drawing.Color _oSelectedRowForegroundColor = System.Drawing.Color.White;
		public ListViewAlternateBackgroundColors.Row_Collection m_oRowCollection = new Row_Collection();

		bool _bCustomFullRowSelect=false;
		string[] m_strColumnNoUpdateArray=null;

		public int m_intSelectedRow=-1;



		public ListViewAlternateBackgroundColors()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		public System.Windows.Forms.ListView ReferenceListView
		{
			set {_oListView = value;}
			get {return _oListView;}
		}
		public System.Drawing.Color ReferenceBackgroundColor
		{
			set {_oBackgroundColor=value;}
			get {return _oBackgroundColor;}
		}
		public System.Drawing.Color ReferenceAlternateBackgroundColor
		{
			set {_oAlternateBackgroundColor=value;}
			get {return _oAlternateBackgroundColor;}
		}
		public System.Drawing.Color ReferenceForegroundColor
		{
			set {_oForegroundColor=value;}
			get {return _oForegroundColor;}
		}
		public System.Drawing.Color ReferenceAlternateForegroundColor
		{
			set {_oAlternateForegroundColor=value;}
			get {return _oAlternateForegroundColor;}
		}
		public System.Drawing.Color ReferenceSelectedRowBackgroundColor
		{
			set {_oSelectedRowBackgroundColor=value;}
			get {return _oSelectedRowBackgroundColor;}
		}
		public System.Drawing.Color ReferenceSelectedRowForegroundColor
		{
			set {_oSelectedRowForegroundColor=value;}
			get {return _oSelectedRowForegroundColor;}
		}
		public bool CustomFullRowSelect
		{
			get {return _bCustomFullRowSelect;}
			set {_bCustomFullRowSelect=value;}
		}
		public void ColumnsToNotUpdate(string p_strList)
		{
	        this.m_strColumnNoUpdateArray = frmMain.g_oUtils.ConvertListToArray(p_strList,",");
		}
		public void ListView()
		{
			int x;
			for (x=0;x<=ReferenceListView.Items.Count-1;x++)
			{
				ListViewItem(ReferenceListView.Items[x]);
			}
		}
		public void InitializeRowCollection()
		{
			int x;
			if (m_oRowCollection != null)
			{
				for (x=m_oRowCollection.Count-1;x>=0;x--)
					m_oRowCollection.Remove(x);

			}
			m_oRowCollection = new Row_Collection();
		}
		public void InitializeRows()
		{
			int x,y;
			if (m_oRowCollection != null)
			{
				for (x=m_oRowCollection.Count-1;x>=0;x--)
					m_oRowCollection.Remove(x);

			}
			m_oRowCollection = new Row_Collection();
			for (x=0;x<=ReferenceListView.Items.Count-1;x++)
			{
				ListViewAlternateBackgroundColors.RowItem oItem = new RowItem();
				for (y=0;y<=ReferenceListView.Items[x].SubItems.Count-1;y++)
				{
					ListViewAlternateBackgroundColors.ColumnItem oColumnItem  = new ColumnItem();
					oColumnItem.ReferenceBackgroundColor = this.ReferenceBackgroundColor;
					oColumnItem.ReferenceAlternateBackgroundColor = this.ReferenceAlternateBackgroundColor;
					oColumnItem.ReferenceSelectedRowBackgroundColor = this.ReferenceSelectedRowBackgroundColor;
					oItem.m_oColumnCollection.Add(oColumnItem);
				}
				m_oRowCollection.Add(oItem);
				
			}
		}
		public void AddRow()
		{
			ListViewAlternateBackgroundColors.RowItem oItem = new RowItem();
			this.m_oRowCollection.Add(oItem);
		}
		public void AddColumn(int p_intRowIndex)
		{
			ListViewAlternateBackgroundColors.ColumnItem oColumnItem  = new ColumnItem();
			oColumnItem.ReferenceBackgroundColor = this.ReferenceBackgroundColor;
			oColumnItem.ReferenceAlternateBackgroundColor = this.ReferenceAlternateBackgroundColor;
			oColumnItem.ReferenceSelectedRowBackgroundColor = this.ReferenceSelectedRowBackgroundColor;
			this.m_oRowCollection.Item(p_intRowIndex).m_oColumnCollection.Add(oColumnItem);
		}
		public void AddColumns(int p_intRowIndex,int p_intColumnCount)
		{
			for (int x=0;x<=p_intColumnCount-1;x++)
			{
				AddColumn(p_intRowIndex);
			}
		}
		
		
		public void ListViewItem(System.Windows.Forms.ListViewItem p_oItem)
		{
			int x,y;
			if (this. m_intSelectedRow != -1 && CustomFullRowSelect)
			{
				//color the unselected row
				if (p_oItem.Selected && p_oItem.Index != this.m_intSelectedRow)
				{
					for (x=0;x<=this.ReferenceListView.Items[m_intSelectedRow].SubItems.Count-1;x++)
					{
						ListViewSubItem(m_intSelectedRow,x,ReferenceListView.Items[m_intSelectedRow].SubItems[x],false);
					}

				}
			}
			for (x=0;x<=p_oItem.SubItems.Count-1;x++)
			{
				ListViewSubItem(p_oItem.Index,x,p_oItem.SubItems[x],p_oItem.Selected);
			}
			if (p_oItem.Selected) m_intSelectedRow=p_oItem.Index;
			
		}
		public void ListViewSubItem(int p_intRow,int p_intCol, System.Windows.Forms.ListViewItem.ListViewSubItem oSubItem,bool p_bSelected)
		{
			if (this.m_oRowCollection.Item(p_intRow).m_oColumnCollection.Item(p_intCol).UpdateColumn)
			{
				if (p_bSelected && this.CustomFullRowSelect)
				{
					//oSubItem.BackColor = this.ReferenceSelectedRowBackgroundColor;
					oSubItem.BackColor = this.m_oRowCollection.Item(p_intRow).m_oColumnCollection.Item(p_intCol).ReferenceSelectedRowBackgroundColor;
					oSubItem.ForeColor = ReferenceSelectedRowForegroundColor;
				}
				else
				{
					if (p_intRow % 2 == 0)
					{
						//oSubItem.BackColor = ReferenceBackgroundColor;
						oSubItem.BackColor = m_oRowCollection.Item(p_intRow).m_oColumnCollection.Item(p_intCol).ReferenceBackgroundColor;
						oSubItem.ForeColor = ReferenceForegroundColor;
					}
					else
					{
						//oSubItem.BackColor = ReferenceAlternateBackgroundColor;
						oSubItem.BackColor = this.m_oRowCollection.Item(p_intRow).m_oColumnCollection.Item(p_intCol).ReferenceAlternateBackgroundColor;
						oSubItem.ForeColor = ReferenceAlternateForegroundColor;
					}
				}
			}
		}
		//THREAD SAFE DELEGATE CALLS
		public void DelegateListView()
		{
			int x;
			int intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(ReferenceListView,"Count",false);
			for (x=0;x<=intCount-1;x++)
			{
				DelegateListViewItem(ReferenceListView.Items[x]);
			}
		}
		public void DelegateListViewItem(System.Windows.Forms.ListViewItem p_oItem)
		{
			int x;
            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(p_oItem.ListView, false);
			int intCount = (int)frmMain.g_oDelegate.GetListViewItemSubItemsPropertyValue(p_oItem,"Count",false);
			int intRow = (int)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,p_oItem.Index,"Index",false);
			bool bSelected = (bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,p_oItem.Index,"Selected",false);

			if (this. m_intSelectedRow != -1 && CustomFullRowSelect)
			{
				//color the unselected row
				if (bSelected && intRow != this.m_intSelectedRow)
				{
					for (x=0;x<=intCount-1;x++)
					{
						if (this.m_intSelectedRow <= this.ReferenceListView.Items.Count-1)
							DelegateListViewSubItem(ReferenceListView.Items[m_intSelectedRow],x);
					}
				}
			}

			for (x=0;x<=intCount-1;x++)
			{
				DelegateListViewSubItem(p_oItem,x);
			}

			if (bSelected) m_intSelectedRow=intRow;
			
		}
		public void DelegateListViewSubItem(System.Windows.Forms.ListViewItem p_oItem,int p_intCol)
		{
			//int intRow = (int)frmMain.g_oDelegate.GetListViewItemPropertyValue(p_oItem.ListView,p_oItem.Index,"Index",false);
            
            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(p_oItem.ListView, false);
            int intIndex = (int)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,p_oItem.Index,"Index",false);
            System.Windows.Forms.ListViewItem oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, intIndex, false);
            int intRow = (int)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, oLvItem.Index, "Index", false);
			
			if (this.m_oRowCollection.Item(intRow).m_oColumnCollection.Item(p_intCol).UpdateColumn)
			{
				if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,p_oItem.Index,"Selected",false) && CustomFullRowSelect)
				{
					//frmMain.g_oDelegate.SetListViewSubItemPropertyValue(p_oItem.ListView,intRow,p_intCol,"BackColor",this.ReferenceSelectedRowBackgroundColor);
					frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv,intRow,p_intCol,"BackColor",this.m_oRowCollection.Item(intRow).m_oColumnCollection.Item(p_intCol).ReferenceSelectedRowBackgroundColor);
					frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv,intRow,p_intCol,"ForeColor",this.ReferenceSelectedRowForegroundColor);
					//frmMain.g_oDelegate.SetListViewSubItemPropertyValue(p_oItem.ListView,intRow,p_intCol,"ForeColor",ReferenceForegroundColor);
				}
				else
				{
					if (intRow % 2 == 0)
					{
						//frmMain.g_oDelegate.SetListViewSubItemPropertyValue(p_oItem.ListView,intRow,p_intCol,"BackColor",ReferenceBackgroundColor);
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv,intRow,p_intCol,"BackColor",this.m_oRowCollection.Item(intRow).m_oColumnCollection.Item(p_intCol).ReferenceBackgroundColor);
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv,intRow,p_intCol,"ForeColor",ReferenceForegroundColor);
					}
					else
					{
						//frmMain.g_oDelegate.SetListViewSubItemPropertyValue(p_oItem.ListView,intRow,p_intCol,"BackColor",ReferenceAlternateBackgroundColor);
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv,intRow,p_intCol,"BackColor",this.m_oRowCollection.Item(intRow).m_oColumnCollection.Item(p_intCol).ReferenceAlternateBackgroundColor);
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv,intRow,p_intCol,"ForeColor", ReferenceAlternateForegroundColor);
					}
				}
			}
		}
        public void DelegateListViewSubItem(System.Windows.Forms.ListViewItem p_oItem, int p_intRow,int p_intCol)
        {
            int intRow = p_intRow;
            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(p_oItem.ListView, false);
            System.Windows.Forms.ListViewItem oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, intRow, false);
            int intIndex = (int)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, intRow, "Index", false);
            if (this.m_oRowCollection.Item(intRow).m_oColumnCollection.Item(p_intCol).UpdateColumn)
            {
                if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, intIndex, "Selected", false) && CustomFullRowSelect)
                {
                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, intRow, p_intCol, "BackColor", this.m_oRowCollection.Item(intRow).m_oColumnCollection.Item(p_intCol).ReferenceSelectedRowBackgroundColor);
                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, intRow, p_intCol, "ForeColor", this.ReferenceSelectedRowForegroundColor);
                }
                else
                {
                    if (intRow % 2 == 0)
                    {
                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, intRow, p_intCol, "BackColor", this.m_oRowCollection.Item(intRow).m_oColumnCollection.Item(p_intCol).ReferenceBackgroundColor);
                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, intRow, p_intCol, "ForeColor", ReferenceForegroundColor);
                    }
                    else
                    {
                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, intRow, p_intCol, "BackColor", this.m_oRowCollection.Item(intRow).m_oColumnCollection.Item(p_intCol).ReferenceAlternateBackgroundColor);
                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, intRow, p_intCol, "ForeColor", ReferenceAlternateForegroundColor);
                    }
                }
            }
        }
		public class RowItem
		{
			public ListViewAlternateBackgroundColors.Column_Collection m_oColumnCollection = new Column_Collection();

			public RowItem()
			{
			}
		}


		public class Row_Collection : System.Collections.CollectionBase
		{
			public Row_Collection()
			{
				//
				// TODO: Add constructor logic here
				//
			}

			public void Add(ListViewAlternateBackgroundColors.RowItem m_oRowItem)
			{
				// vérify if object is not already in
				if (this.List.Contains(m_oRowItem))
					throw new InvalidOperationException();
 
				// adding it
				this.List.Add(m_oRowItem);
 
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
			public ListViewAlternateBackgroundColors.RowItem Item(int Index)
			{
				// The appropriate item is retrieved from the List object and
				// explicitly cast to the Widget type, then returned to the 
				// caller.
				return (ListViewAlternateBackgroundColors.RowItem) List[Index];
			}
		}
		public class ColumnItem
		{
			public System.Drawing.Color _oBackgroundColor=frmMain.g_oGridViewBackgroundColor;
			public System.Drawing.Color _oAlternateBackgroundColor=frmMain.g_oGridViewAlternateRowBackgroundColor;
			public System.Drawing.Color _oSelectedBackgroundColor=frmMain.g_oGridViewSelectedRowBackgroundColor;
			public bool _bUpdateColumn=true;

			public ColumnItem()
			{
			}

			public System.Drawing.Color ReferenceBackgroundColor
			{
				set {_oBackgroundColor=value;}
				get {return _oBackgroundColor;}
			}
			public System.Drawing.Color ReferenceAlternateBackgroundColor
			{
				set {_oAlternateBackgroundColor=value;}
				get {return _oAlternateBackgroundColor;}
			}
			public bool UpdateColumn
			{
				set {_bUpdateColumn=value;}
				get {return _bUpdateColumn;}
			}
			
			public System.Drawing.Color ReferenceSelectedRowBackgroundColor
			{
				set {_oSelectedBackgroundColor=value;}
				get {return _oSelectedBackgroundColor;}
			}

		}
		public class Column_Collection : System.Collections.CollectionBase
		{
			public Column_Collection()
			{
				//
				// TODO: Add constructor logic here
				//
			}

			public void Add(ListViewAlternateBackgroundColors.ColumnItem m_oColumnItem)
			{
				// vérify if object is not already in
				if (this.List.Contains(m_oColumnItem))
					throw new InvalidOperationException();
 
				// adding it
				this.List.Add(m_oColumnItem);
 
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
			public ListViewAlternateBackgroundColors.ColumnItem Item(int Index)
			{
				// The appropriate item is retrieved from the List object and
				// explicitly cast to the Widget type, then returned to the 
				// caller.
				return (ListViewAlternateBackgroundColors.ColumnItem) List[Index];
			}
		}
	}
}
