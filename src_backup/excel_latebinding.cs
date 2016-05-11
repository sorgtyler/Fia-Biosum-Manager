using System;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows.Forms;
using FIA_Biosum_Manager;

//using System.Management;

namespace excel_latebinding
{
    
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public class excel_latebinding
	{
    public delegate void SaveWorkbookCallback(object oBook, string strFilePathAndName);
		private utils m_utils=new utils();
		private bool _bUseRandomData=true;

		//estimation engine table used for report
		private System.Data.DataTable m_oDtEstEng;
		//declare the row that will contain the datarow found using oPopEvalExSearch
		private System.Data.DataRow m_oDataRow; 

		Type oExcelAppType;

        const int xlAutoActivate = 3;
        const int xlAutoClose = 2;
        const int xlAutoDeactivate = 4;
        const int xlAutoOpen = 1;


		const int xlAutomatic=-4105;
		const int xlManual =-4135;
		const int xlUpward =-4171;
		const int xlWait   =2;

		//excel horizontal alignments
		const int xlAlignCenter=-4108;
		const int xlAlignLeft = -4131;
		const int xlAlignRight = -4152;
		const int xlAlignCenterAcrossSelection = 7;

		//excel vertial alignments
		const int xlVAlignBottom=-4107;
		const int xlVAlignCenter=-4108;
		const int xlVAlignDistributed=-4117;
		const int xlVAlignJustify=-4130;
		const int xlVAlignTop=-4160;

		//excel line weight
		const int xlHairLine=1;
		const int xlMedium = -4138;
		const int xlThick = 4;
		const int xlThin = 2;

		//excel line style
		const int xlNone=-4142;
		const int xlContinuous=1;
		const int xlDash=-4115;
		const int xlDashDot=4;
		const int xlDashDotDot=5;
		const int xlDot=-4118;
		const int xlDouble=-4119;
		const int xlSlantDashDot=13;

		//excel line placement
		const int xlDiagonalDown=5;
		const int xlDiagonalUp=6;
		const int xlEdgeLeft=7;
		const int xlEdgeTop=8;
		const int xlEdgeBottom=9;
		const int xlEdgeRight=10;
		const int xlInsideVertical=11;
		const int xlInsideHorizontal=12;


		object oExcelApp;
		object oBook;
		object oBooks;
		object oSheets;
		object oSheet;
		object oCells;
		object oRange;
		object oCharacters;
		object oFont;
		object oFootnoteFont;
		object oInterior;
		object oColumns;
		object oRows;
		object oBorders;
		object oPageSetup;
		long hWnd;
		public string m_strError;
		public int m_intError;

		private int m_intCurrentSheetIndex;
		private string m_strCurrentSheetName="";
		private string _strPrinterName;
		private bool _bLandscape;
		private int _intLeftMargin;
		private int _intRightMargin;
		private int _intTopMargin;
		private int _intBottomMargin;
		private string _strExcelFileName="";
        private bool _bDisplayAlerts=true;
		private object oWScript;
		Type oWScriptType;


	

		public excel_latebinding()
		{
			//
			// TODO: Add constructor logic here
			//

			oWScriptType = Type.GetTypeFromProgID("WScript.Network");
			oWScript = Activator.CreateInstance(oWScriptType);

	
		}
		~excel_latebinding()
		{
            if (oPageSetup != null) Marshal.ReleaseComObject(oPageSetup);
			if (oBorders != null)   Marshal.ReleaseComObject(oBorders);
			if (oRows != null) Marshal.ReleaseComObject(oRows);
			if (oColumns != null) Marshal.ReleaseComObject(oColumns);
			if (oInterior != null) Marshal.ReleaseComObject(oInterior);
			if (oFont != null) Marshal.ReleaseComObject(oFont);
			if (oRange != null) Marshal.ReleaseComObject(oRange);
			if (oCells != null) Marshal.ReleaseComObject(oCells);
			if (oSheet != null) Marshal.ReleaseComObject(oSheet);
			if (oSheets != null) Marshal.ReleaseComObject(oSheets);
			if (oBook != null) Marshal.ReleaseComObject(oBook);
			if (oBooks != null) Marshal.ReleaseComObject(oBooks);
			if (oExcelApp != null) Marshal.ReleaseComObject(oExcelApp);


		}
        public void ReleaseComObjects()
        {
            if (oPageSetup != null) Marshal.ReleaseComObject(oPageSetup);
            if (oBorders != null) Marshal.ReleaseComObject(oBorders);
            if (oRows != null) Marshal.ReleaseComObject(oRows);
            if (oColumns != null) Marshal.ReleaseComObject(oColumns);
            if (oInterior != null) Marshal.ReleaseComObject(oInterior);
            if (oFont != null) Marshal.ReleaseComObject(oFont);
            if (oRange != null) Marshal.ReleaseComObject(oRange);
            if (oCells != null) Marshal.ReleaseComObject(oCells);
            if (oSheet != null) Marshal.ReleaseComObject(oSheet);
            if (oSheets != null) Marshal.ReleaseComObject(oSheets);
            if (oBook != null) Marshal.ReleaseComObject(oBook);
            if (oBooks != null) Marshal.ReleaseComObject(oBooks);
            if (oExcelApp != null) Marshal.ReleaseComObject(oExcelApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();



        }
		public void StartExcelApplication()
		{
			m_strError="";
			m_intError=0;
			try
			{
				oExcelAppType = Type.GetTypeFromProgID("Excel.Application");
				oExcelApp = Activator.CreateInstance(oExcelAppType);
				this.SetProperty(oExcelApp,"DisplayAlerts",DisplayAlerts);
                this.SetProperty(oExcelApp, "EnableEvents", true);
				
				
			}
			catch (Exception err)
			{
				m_strError=err.Message.ToString();
				m_intError=-1;
				MessageBox.Show(m_strError,"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
			}
		}
        public void OpenWorkBook()
        {
            try
            {
                oBooks = GetProperty(oExcelApp, "Workbooks");
                oBook = this.InvokeMethod(oBooks, "Open", ExcelFileName);
                this.InvokeMethod(oBook,"RunAutoMacros", (object)xlAutoOpen);
                

                oSheets = GetProperty(oBook, "Worksheets");
                string strTitle = (string)oExcelAppType.InvokeMember("Caption", BindingFlags.GetProperty, null, oExcelApp, null);
                hWnd = m_utils.FindWindowLike((System.IntPtr)0, strTitle, "*", true, true);
            }
            catch (Exception err)
            {
                m_strError = err.Message.ToString();
                m_intError = -1;
                MessageBox.Show(m_strError, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        public void OpenNewWorkBook()
        {
            try
            {
                oBook = this.InvokeMethod(oBooks, "Add", "");
                oSheets = GetProperty(oBook, "Worksheets");
                string strTitle = (string)oExcelAppType.InvokeMember("Caption", BindingFlags.GetProperty, null, oExcelApp, null);
                hWnd = m_utils.FindWindowLike((System.IntPtr)0, strTitle, "*", true, true);
            }
            catch (Exception err)
            {
                m_strError = err.Message.ToString();
                m_intError = -1;
                MessageBox.Show(m_strError, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        public void RunMacro(string p_strMacroName)
        {
            oExcelApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default |
                     System.Reflection.BindingFlags.InvokeMethod, null, oExcelApp, new object[] {p_strMacroName});
            
        }

		
		#region private Wrappers
		private void SetProperty(object obj,string sProperty,object oValue)
		{
			object[] oParam=new object[1];
			oParam[0]=oValue;
			obj.GetType().InvokeMember(sProperty,BindingFlags.SetProperty, null, obj, oParam );
		}
		private object GetProperty(object obj,string sProperty,object oValue)
		{
			object[] oParam=new object[1];
			oParam[0]=oValue;
			return obj.GetType().InvokeMember(sProperty,BindingFlags.GetProperty, null, obj, oParam );
		}
		private object GetProperty(object obj,string sProperty,object oValue1,object oValue2)
		{
			object[] oParam=new object[2];
			oParam[0]=oValue1;
			oParam[1]=oValue2;
			return obj.GetType().InvokeMember(sProperty,BindingFlags.GetProperty, null, obj, oParam );
		}
		private object GetProperty(object obj,string sProperty)
		{
			return obj.GetType().InvokeMember(sProperty,BindingFlags.GetProperty, null, obj, null );
		}
		private object InvokeMethod(object obj,string sProperty,object[] oParam)
		{
			return obj.GetType().InvokeMember(sProperty,BindingFlags.InvokeMethod, null, obj, oParam);
		}
		private object InvokeMethod(object obj,string sProperty,object oValue)
		{
			object[] oParam=new object[1];
			oParam[0]=oValue;
			return obj.GetType().InvokeMember(sProperty,BindingFlags.InvokeMethod, null, obj, oParam );
		}
		#endregion

		
		public void Show()
		{
			try
			{
				this.SetProperty(oExcelApp,"Visible",true);
				if ((int)this.GetProperty(oBooks,"Count")==0)
				{
					InitializeWorkbook();
				}
				if (m_intError==0)
				{
					m_utils.SetFocus(hWnd);
					oExcelAppType.InvokeMember("WindowState",BindingFlags.SetProperty,null,oExcelApp, new object [] {(int)1});
					m_utils.BringWindowToTop(hWnd);
					
				}
			}
			catch (Exception err)
			{
				m_strError=err.Message.ToString();
				m_intError=-1;
				MessageBox.Show(m_strError,"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);

			}
		}
		public void Hide()
		{
			try
			{
				this.SetProperty(oExcelApp,"Visible",true);
				if ((int)this.GetProperty(oBooks,"Count")==0)
				{
					InitializeWorkbook();
				}
				if (m_intError==0)
				{
					m_utils.SetFocus(hWnd);
                    this.SetProperty(oExcelApp, "Visible", false);
					m_utils.BringWindowToTop(hWnd);
					
				}
			}
			catch (Exception err)
			{
				m_strError=err.Message.ToString();
				m_intError=-1;
				MessageBox.Show(m_strError,"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);

			}
		}
		private void InitializeWorkbook()
		{
			m_intError=0;
			m_strError="";
			try
			{
				oBooks=GetProperty(oExcelApp,"Workbooks");
				oBook = this.InvokeMethod(oBooks,"Open", ExcelFileName);
				
				
				oSheets = GetProperty(oBook,"Worksheets");
				

			}
			catch (Exception e)
			{
				m_strError=e.Message.ToString();
				m_intError=-1;
				MessageBox.Show(m_strError,"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
			}
		}
		public void SaveAsWorkbook(string p_strFilePathAndName)
		{
			m_intError=0;
			m_strError="";
			try
			{
                
				oBook.GetType().InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod, null, oBook, new object[] {p_strFilePathAndName});
			}
			catch (Exception e)
			{
				m_intError=-1;
				m_strError="Excel_Latebinding.SaveAsWorkbook:" + e.Message;
				MessageBox.Show(m_strError,"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
			}
		}
        public void SaveWorkbook()
        {
            m_intError = 0;
            m_strError = "";
            try
            {
                oBook.GetType().InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, oBook, null);
            }
            catch (Exception e)
            {
                m_intError = -1;
                m_strError = "Excel_Latebinding.SaveWorkbook:" + e.Message;
                MessageBox.Show(m_strError, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        public void SaveAndExit()
        {
            SaveWorkbook();
            QuitExcelApplication();
        }
		public void DeleteAllSheets()
		{
			oSheets = GetProperty(oBook,"Worksheets");
			if ((int)this.GetProperty(oSheets,"Count") > 0)
			{

				//delete all but the first sheet
				for (int x=(int)this.GetProperty(oSheets,"Count");x>1;x--)
				{
						
					oSheet = GetProperty(oSheets,"Item",x);
					oSheet.GetType().InvokeMember("Delete",BindingFlags.InvokeMethod,null,oSheet,null);
						
				}
					
			}
		}
		public void AddSheet(string strSheetName)
		{
			try
			{
				//add a sheet
				oSheets.GetType().InvokeMember("Add",BindingFlags.InvokeMethod,null,oSheets,null);
				//get the last sheet just added
				oSheet = GetProperty(oSheets,"Item",1);
				//assign a name to the sheet
				SetProperty(oSheet,"Name",strSheetName);
				oCells = oSheet.GetType().InvokeMember("Cells",BindingFlags.InvokeMethod | BindingFlags.GetProperty,null,oSheet,null);
			}
			catch (Exception e)
			{
				m_intError=-1;
				m_strError="excel_latebinding:AddSheet:" + e.Message;
				MessageBox.Show(m_strError,"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
			}
		}
		public void AssignSheetName(int p_intSheetIndex,string p_strNewSheetName)
		{
			m_intError=0;
			m_strError="";
			try
			{
				oSheet = GetProperty(oSheets,"Item",p_intSheetIndex);
				SetProperty(oSheet,"Name",p_strNewSheetName);
			}
			catch (Exception e)
			{
				m_intError=-1;
				m_strError=e.Message;
				MessageBox.Show(m_strError,"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
			}
		}
		public void AssignCurrentSheet(string strSheetName)
		{
			m_intError=0;
			m_strError="";
			try
			{
				this.m_intCurrentSheetIndex=-1;
				string strName;
				for (int x=(int)this.GetProperty(oSheets,"Count");x>0;x--)
				{
						
					oSheet = GetProperty(oSheets,"Item",x);
					strName = Convert.ToString(GetProperty(oSheet,"Name"));
					if (strName.Trim().ToUpper()==strSheetName.Trim().ToUpper())
					{
						this.m_intCurrentSheetIndex=x;
						this.m_strCurrentSheetName=strSheetName;
						break;
					}
						
				}
			}
			catch (Exception e)
			{
				m_intError=-1;
				m_strError=e.Message;
				MessageBox.Show(m_strError,"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
			}

		}
        public void DeleteSheet(string p_strSheetName)
        {
            string strName = "";
            oSheets = GetProperty(oBook, "Worksheets");
            if ((int)this.GetProperty(oSheets, "Count") > 0)
            {

                
                for (int x = (int)this.GetProperty(oSheets, "Count"); x > 1; x--)
                {
                    
                    oSheet = GetProperty(oSheets, "Item", x);
                    strName = Convert.ToString(GetProperty(oSheet, "Name"));
                    if (strName.Trim().ToUpper() == p_strSheetName.Trim().ToUpper())
                    {
                       oSheet.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, oSheet, null);
                       break;
                    }

                }

            }
        }
        public void UnprotectSheet()
        {
            this.InvokeMethod(oSheet, "Unprotect", null);
        }
		public void QuitExcelApplication()
		{
			try
			{
				if (oPageSetup != null)
				{
					Marshal.ReleaseComObject(oPageSetup);
				}
				oPageSetup = null;
				// CLOSE OUR EXCEL APPLICATION OBJECT

                this.SetProperty(oExcelApp, "DisplayAlerts", this.DisplayAlerts);
                this.SetProperty(oExcelApp, "EnableEvents", false);
				this.oExcelApp.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.IgnoreReturn | System.Reflection.BindingFlags.Instance | 
					System.Reflection.BindingFlags.InvokeMethod, null, this.oExcelApp, null);


			}
			catch
			{
			}
			
		}
        /// <summary>
        /// Assign a cell value to the current sheet.
        /// </summary>
        /// <param name="p_strCellRange">Format is CCCRRR:CCCRRR where CCC=column=(Alpha), RRR=row=(Numeric), :=range separator</param>
        /// <param name="p_strValue"></param>
        public void AssignCellValue(string p_strCellRange, string p_strCellValue)
        {
            object[] RowColParam = new Object[2];
            string[] strRangeArray = frmMain.g_oUtils.ConvertListToArray(p_strCellRange, ":");
            strRangeArray[0] = strRangeArray[0].Trim();
            strRangeArray[1] = strRangeArray[1].Trim();
            RowColParam[0] = (object)strRangeArray[0];
            RowColParam[1] = (object)strRangeArray[1];
            oRange = oSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, oSheet, RowColParam);
            oRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, oRange, new object[] { p_strCellValue});
        }
		
		
		private string GetCellColumn(string strColRow)
		{
			string str="";
			string strNum="0123456789";
			int x;
			for (x=0;x<=strColRow.Trim().Length-1;x++)
			{
				if (strNum.IndexOf(strColRow.Substring(x,1),0) >= 0)
				{
					break;
				}
				else
				{
					str=str+strColRow.Substring(x,1);
				}
			}
			return str;
				
		}
		private string GetCellRow(string strColRow)
		{
			string str="";
			string strNum="0123456789";
			int x;
			for (x=0;x<=strColRow.Trim().Length-1;x++)
			{
				if (strNum.IndexOf(strColRow.Substring(x,1),0) < 0)
				{
				}
				else
				{
					str=str+strColRow.Substring(x,1);
				}
			}
			return str;
		}
		private string GetColumn(int intCol)
		{
			string str="";
			try
			{
				
				string strAlpha="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
				if (intCol < 26)
				{
					str=strAlpha.Substring(intCol,1);
				}
				else if (intCol < 52)
				{

					str="A" + strAlpha.Substring(intCol-26,1);
				}
				else if (intCol < 78)
				{
					str="B" + strAlpha.Substring(intCol-52,1);
				}
			}
			catch (Exception e)
			{
				this.m_intError=-1;
				this.m_strError="excel_latebinding.GetColumn:" + e.Message;
				MessageBox.Show("excel_latebinding.GetColumn:" + e.Message,"Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
			return str;
		}
		private string GetSEColumn(string strCol)
		{
			try
			{
				string strAlpha="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
				int intIndex = strAlpha.IndexOf(strCol.ToUpper(),0);
				return strAlpha.Substring(intIndex+1,1);
			}
			catch (Exception e)
			{
				m_intError=-1;
				m_strError=e.Message;
				MessageBox.Show(e.Message,"Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
			return "";
		}
		public int GetColor(string strColor)
		{
			switch (strColor.Trim())
			{
				case "Black": return 1;
                case "White": return 2;
				case "DarkGray": return 56;
				case "Gray": return 48;
				case "LightGray": return 15;
			    case "DarkRed": return 9;
			    case "Red": return 3;
			    case "LightRed": return 22;
			    case "DarkGreen": return 51;
			    case "Green": return 10;
			    case "LightGreen": return 35;
			    case "SeaGreen" :  return 50;
			    case "OliveGreen": return 52;
			    case "BrightGreen": return 4;
			    case "DarkBlue": return 11;
			    case "Blue": return 5;
			    case "LightBlue": return 41;
			    case "SkyBlue": return 33;
			    case "PaleBlue": return 37;
			    case "Indigo": return 55;
			    case "BlueGray": return 47;
			    case "Violet": return 13;
			    case "Plum": return 54;
			    case "Lavender": return 39;
			    case "DarkTeal": return 49;
			    case "Teal": return 14;
			    case "Aqua": return 42;
			    case "Turquoise": return 28;
			    case "LightTurquoise": return 34;
			    case "DarkYellow": return 12;
			    case "Lime": return 43;
			    case "Yellow": return 6;
			    case "LightYellow": return 36;
				case "Magenta": return 7;
			    case "Cyan": return 8;
			    case "DarkPurple": return 21;
			    case "Purple": return 29;
			    case "LightPurple": return 24;
			    case "Pink": return 26;
			    case "Rose": return 38;
				case "Orange": return 46;
			    case "LightOrange": return 45;
			    case "Gold": return 44;
			    case "Brown": return 53;
			    case "Tan": return 40;
				default: return 2;
			}
		}
		public string GetColor(int p_intColor)
		{
			

			switch (p_intColor)
			{
				case 1: return "Black";
				case 2: return "White";
				case 56: return "DarkGray";
				case 48: return "Gray";
				case 15: return "LightGray";
				case 9: return "DarkRed";
				case 3: return "Red";
				case 22: return "LightRed";
				case 51: return "DarkGreen";
				case 10: return "Green";
				case 35: return "LightGreen";
				case 50: return "SeaGreen";
				case 52: return "OliveGreen";
				case 4: return "BrightGreen";
				case 11: return "DarkBlue";
				case 5: return "Blue";
				case 41: return "LightBlue";
				case 33: return "SkyBlue";
				case 37: return "PaleBlue";
				case 55: return "Indigo";
				case 47: return "BlueGray";
				case 13: return "Violet";
				case 54: return "Plum";
				case 39: return "Lavender";
				case 49: return "DarkTeal";
				case 14: return "Teal";
				case 42: return "Aqua";
				case 28: return "Turquoise";
				case 34: return "LightTurquoise";
				case 12: return "DarkYellow";
				case 43: return "Lime";
				case 6: return "Yellow";
				case 36: return "LightYellow";
				case 7: return "Magenta";
				case 8: return "Cyan";
				case 21: return "DarkPurple";
				case 29: return "Purple";
				case 24: return "LightPurple";
				case 26: return "Pink";
				case 38: return "Rose";
				case 46: return "Orange";
				case 45: return "LightOrange";
				case 44: return "Gold";
				case 53: return "Brown";
				case 40: return "Tan";
				default: return "White";
			}
		}
		/// <summary>
		/// the value saved in the background and foreground color properties is the menu array item number. This function
		/// translates the menu array item number to the excel recognized color value.
		/// </summary>
		/// <param name="p_intEnum"></param>
		/// <returns></returns>
		public int GetColorByEnum(int p_intEnum)
		{
			switch (p_intEnum)
			{
				case 0: return 1;   //black
				case 1: return 2;   //white
				case 2: return 56;  //Dark Gray
				case 3: return 48;  //Gray
				case 4: return 15;  //Light Gray
				case 5: return 9;   //Dark Red
				case 6: return 3;   //red
				case 7: return 22;  //Light Red
				case 8: return 51;  //Dark Green
				case 9: return 10;  //Green
				case 10: return 35; //Light Green
				case 11: return 50; //Sea Green
				case 12: return 52; //OliveGreen
				case 13: return 4;  //BrightGreen
				case 14: return 11; //DarkBlue
				case 15: return 5;  //Blue
				case 16: return 41; //LightBlue
				case 17: return 33; //SkyBlue
				case 18: return 37; //PaleBlue
				case 19: return 55; //Indigo
				case 20: return 47; //bluegray
				case 21: return 13; //violet
				case 22: return 54; //plum
				case 23: return 39; //lavender
				case 24: return 49; //DarkTeal
				case 25: return 14; //teal
				case 26: return 42; //Aqua
				case 27: return 8;  //turquoise
				case 28: return 34; //LightTurquoise
				case 29: return 12; //DarkYellow
				case 30: return 43; //Lime
				case 31: return 6;  //Yellow
				case 32: return 36; //LightYellow
				case 33: return 7;  //Magenta
				case 34: return 8;  //Cyan
				case 35: return 21; //DarkPurple
				case 36: return 29; //Purple
				case 37: return 24; //LightPurple
				case 38: return 7;  //Pink
				case 39: return 38; //Rose
				case 40: return 46; //Orange
				case 41: return 45; //LightOrange
				case 42: return 44; //Gold
				case 43: return 53; //Brown
				case 44: return 40; //Tan
				default: return 2;
			}
		}

		private int GetHorizontalAlignment(string str)
		{
			switch (str.Trim().ToUpper())
			{
				case "LEFT": return excel_latebinding.xlAlignLeft;
				case "CENTER": return excel_latebinding.xlAlignCenter;
				case "RIGHT": return excel_latebinding.xlAlignRight;
				case "CENTERACROSSSELECTION": return excel_latebinding.xlAlignCenterAcrossSelection;
				default: return excel_latebinding.xlAutomatic;
			}
		}
		private int GetVerticalAlignment(string str)
		{
			switch (str.Trim().ToUpper())
			{
				case "BOTTOM": return excel_latebinding.xlVAlignBottom;
				case "CENTER": return excel_latebinding.xlVAlignCenter;
				case "DISTRIBUTED": return excel_latebinding.xlVAlignDistributed;
				case "JUSTIFY": return excel_latebinding.xlVAlignJustify;
				case "TOP": return excel_latebinding.xlVAlignTop;
				default: return excel_latebinding.xlVAlignBottom;
			}
		}
		private int GetLineStyle(string str)
		{
			switch (str.Trim().ToUpper())
			{
				case "NONE": return excel_latebinding.xlNone;
				case "CONTINUOUS": return excel_latebinding.xlContinuous;
				case "DASH": return excel_latebinding.xlDash;
				case "DASHDOT": return excel_latebinding.xlDashDot;
				case "DASHDOTDOT": return excel_latebinding.xlDashDotDot;
				case "DOT": return excel_latebinding.xlDot;
				case "DOUBLE": return excel_latebinding.xlDouble;
				case "SLANTDASHDOT": return excel_latebinding.xlSlantDashDot;
				default: return excel_latebinding.xlNone;
			}
		}
		private int GetLineWeight(string str)
		{
			switch (str.Trim().ToUpper())
			{
				case "HAIRLINE": return excel_latebinding.xlHairLine;
				case "MEDIUM": return excel_latebinding.xlMedium;
				case "THICK": return excel_latebinding.xlThick;
				case "THIN": return excel_latebinding.xlThin;
                default: return excel_latebinding.xlMedium;
			}
		}

		private int GetLinePlacement(string str)
		{
			switch (str.Trim().ToUpper())
			{
				case "DIAGONALDOWN": return excel_latebinding.xlDiagonalDown;
				case "DIAGONALUP": return excel_latebinding.xlDiagonalUp;
				case "EDGELEFT": return excel_latebinding.xlEdgeLeft;
				case "EDGERIGHT": return excel_latebinding.xlEdgeRight;
				case "EDGETOP": return excel_latebinding.xlEdgeTop;
				case "EDGEBOTTOM": return excel_latebinding.xlEdgeBottom;
				case "INSIDEVERTICAL": return excel_latebinding.xlInsideVertical;
				case "INSIDEHORIZONTAL": return excel_latebinding.xlInsideHorizontal;
				default: return excel_latebinding.xlEdgeBottom;
			}
		}
		/// <summary>
		/// cell string is formatted as COLUMNROW (ex. A1 for column A Row 1).  This routine
		/// splits the cell string into the appropriate row and column value
		/// </summary>
		/// <param name="strCell">Cell value. Example, "A1" for column A and Row 1.</param>
		/// <param name="strCol">Column value</param>
		/// <param name="intRow">Row value</param>
		private void SplitCellRowCol(string strCell,ref string strCol, ref int intRow)
		{
            int x;
            string strAlpha="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string strC="";
		    string strR="";
			string str=strCell.ToUpper();
			for (x=0;x<=str.Length-1;x++)
			{ 
				if (strAlpha.IndexOf(str.Substring(x,1),0) >=0)
				{
					strC=strC + str.Substring(x,1);
				}
				else
				{
					strR=strR + str.Substring(x,1);
				}

			}
			strCol=strC;
			intRow=Convert.ToInt32(strR);
		}

		
	
		public string PrinterName
		{
			get {return _strPrinterName;}
			set {_strPrinterName=value;}
		}
		public bool Landscape
		{
			get {return _bLandscape;}
			set {_bLandscape=value;}
		}
		public int LeftMargin
		{
			get {return _intLeftMargin;}
			set {_intLeftMargin=value;}
		}
		public int RightMargin
		{
			get {return _intRightMargin;}
			set {_intRightMargin=value;}
		}
		public int TopMargin
		{
			get {return _intTopMargin;}
			set {_intTopMargin=value;}
		}
		public int BottomMargin
		{
			get {return _intBottomMargin;}
			set {_intBottomMargin=value;}
		}
		
		public string ExcelFileName
		{
			set {_strExcelFileName=value;}
			get {return _strExcelFileName;}
		}
        public bool DisplayAlerts
        {
            set { _bDisplayAlerts = value; }
            get { return _bDisplayAlerts; }
        }




	
	}

}
