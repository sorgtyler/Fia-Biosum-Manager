using System;
using System.Windows.Forms;
public  struct PrintHeaderFormat
		{
			public char[] chrLabel;
			public char[] chrFiller;
			public char[] chrDesc;
		}
public  struct PrintFieldFormat
		{
			public char[] chrFieldNumber;
			public char[] chrFiller;
			public char[] chrFieldName;
			public char[] chrFiller2;
			public char[] chrFieldType;
		
		}
namespace FIA_Biosum_Manager
{
	
	/// <summary>
	/// Summary description for TableStructure.
	/// </summary>
	
	public class TableStructure
	{
		
		private System.Data.DataTable m_table;
		
		private PrintHeaderFormat m_printheader;
		private PrintFieldFormat  m_printfield;
       
        public TableStructure()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		public TableStructure(System.Data.DataTable p_table,string strAction)
		{
			
			this.m_table = p_table; //this.m_ds.Tables(strDataSet);

			if (strAction == "VIEW")
			{
				this.m_printheader = new PrintHeaderFormat();
				this.m_printheader.chrLabel = new char[22];
				this.m_printheader.chrFiller = new char[2];
				this.m_printheader.chrDesc = new char[90];
				this.m_printheader.chrDesc.Initialize();
				this.m_printheader.chrFiller.Initialize();
				this.m_printheader.chrLabel.Initialize();

                this.m_printfield = new PrintFieldFormat();
                this.m_printfield.chrFieldNumber = new char[5];
				this.m_printfield.chrFiller = new char[3];
				this.m_printfield.chrFieldName = new char[50];
				this.m_printfield.chrFiller2 = new char[3];
				this.m_printfield.chrFieldType = new char[25];
			
				ReadOnlyListStruc();
			}

			
		}
		private void ReadOnlyListStruc()
		{
			string strLargestString="";
			string str="";
			int x=0;

			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.uc_select_list_item1.listBox1.Font = new System.Drawing.Font("Courier New", 8);
			frmTemp.Text = "Table Structure";

			frmTemp.uc_select_list_item1.lblMsg.Text= "Table structure of " + this.m_table.TableName;
			frmTemp.uc_select_list_item1.lblMsg.Visible = true;
		
			frmTemp.uc_select_list_item1.Dock = System.Windows.Forms.DockStyle.Fill;
						
			frmTemp.uc_project1.Visible=false;
			frmTemp.uc_select_list_item1.listBox1.Items.Clear();
			//list the header information
			utils p_utils = new utils();

			//print header information
			//dataset name
			p_utils.LoadStringToCharArray("Structure For DataSet:", ref this.m_printheader.chrLabel,true,0);
			p_utils.LoadStringToCharArray("  ", ref this.m_printheader.chrFiller,true,0);
            p_utils.LoadStringToCharArray(this.m_table.TableName ,ref this.m_printheader.chrDesc,true,0);
            p_utils.LoadCharArrayToString(ref str, this.m_printheader.chrLabel,false);
			p_utils.LoadCharArrayToString(ref str, this.m_printheader.chrFiller,false);
            p_utils.LoadCharArrayToString(ref str, this.m_printheader.chrDesc,false);
    		frmTemp.uc_select_list_item1.listBox1.Items.Add(str);
			strLargestString = str.Trim();
            
			//number of records
			str="";
			p_utils.LoadStringToCharArray("Number Of Records:", ref this.m_printheader.chrLabel,true,0);
			p_utils.LoadStringToCharArray("  ", ref this.m_printheader.chrFiller,true,0);
			p_utils.LoadStringToCharArray(this.m_table.Rows.Count.ToString() ,ref this.m_printheader.chrDesc,true,0);
			p_utils.LoadCharArrayToString(ref str, this.m_printheader.chrLabel,true);
			p_utils.LoadCharArrayToString(ref str, this.m_printheader.chrFiller,false);
			p_utils.LoadCharArrayToString(ref str, this.m_printheader.chrDesc,false);
			frmTemp.uc_select_list_item1.listBox1.Items.Add(str);
			if (str.Trim().Length > strLargestString.Trim().Length) strLargestString = str.Trim();

            //field information column headers
            p_utils.LoadStringToCharArray("Field", ref this.m_printfield.chrFieldNumber,true,0);
			p_utils.LoadStringToCharArray("   ", ref this.m_printfield.chrFiller,true,0);
            p_utils.LoadStringToCharArray("Field Name", ref this.m_printfield.chrFieldName,true,0);
            p_utils.LoadStringToCharArray("   ", ref this.m_printfield.chrFiller2,true,0);
            p_utils.LoadStringToCharArray("Type", ref this.m_printfield.chrFieldType,true,0);
            p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFieldNumber,true);
            p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFiller,false);
			p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFieldName,false);
			p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFiller2,false);
			p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFieldType,false);
			frmTemp.uc_select_list_item1.listBox1.Items.Add(str);
			if (str.Trim().Length > strLargestString.Trim().Length) strLargestString = str;

			//detail
			for (x=0; x<=this.m_table.Columns.Count-1;x++)
			{
				p_utils.LoadStringToCharArray(Convert.ToString(x+1), ref this.m_printfield.chrFieldNumber,true,0);
				p_utils.LoadStringToCharArray("   ", ref this.m_printfield.chrFiller,true,0);
				p_utils.LoadStringToCharArray(this.m_table.Columns[x].ColumnName.ToString(), ref this.m_printfield.chrFieldName,true,0);
				p_utils.LoadStringToCharArray("   ", ref this.m_printfield.chrFiller2,true,0);
				p_utils.LoadStringToCharArray(this.m_table.Columns[x].DataType.ToString(), ref this.m_printfield.chrFieldType,true,0);
				p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFieldNumber,true);
				p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFiller,false);
				p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFieldName,false);
				p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFiller2,false);
				p_utils.LoadCharArrayToString(ref str, this.m_printfield.chrFieldType,false);
				frmTemp.uc_select_list_item1.listBox1.Items.Add(str);
				if (str.Trim().Length > strLargestString.Trim().Length) strLargestString = str;
			}


			frmTemp.uc_select_list_item1.Initialize_Width(strLargestString + "***********");
			frmTemp.uc_select_list_item1.Visible=true;
			frmTemp.uc_select_list_item1.btnOK.Text = "Print";
			DialogResult result = frmTemp.ShowDialog();
			
			if (result == DialogResult.OK) 
			{
				printing p_oPrint = new printing();
				p_oPrint.PrintListBoxContents("FIA Biosum Data Set Structure", frmTemp.uc_select_list_item1.listBox1);
				p_oPrint = null;

			}

			
			frmTemp.Close();
			frmTemp = null;

		}
		

	}
}
