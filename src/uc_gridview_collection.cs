using System;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_gridview_collection.
	/// </summary>
	public class uc_gridview_collection : System.Collections.CollectionBase
	{
		public uc_gridview_collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(uc_gridview uc_gridview1)
		{
			// vérify if object is not already in
			if (this.List.Contains(uc_gridview1))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(uc_gridview1);
 
			// return collection
			//return this;
		}
		public void Remove(int index)
		{
			// Check to see if there is a widget at the supplied index.
			if (index > Count - 1 || index < 0)
				// If no widget exists, a messagebox is shown and the operation 
				// is cancelled.
			{
				System.Windows.Forms.MessageBox.Show("Index not valid!");
			}
			else
			{
				List.RemoveAt(index); 
			}
		}
		public uc_gridview Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (uc_gridview) List[Index];
		}



	}
}

