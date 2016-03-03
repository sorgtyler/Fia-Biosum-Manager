using System;
using System.Windows.Forms;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for txtWholesAndDecimals.
	/// </summary>
	/// 
	public class txtNumeric:System.Windows.Forms.TextBox
	{
		public bool bInitialize=true;
		public bool bEdit = false;
		private int intWholeMaxLen=4;       //default value
		private int intDecimalMaxLen=2;     //default value
		private int intWholeCurLen=0;
		private int intDecimalCurLen=0;
		private string strLastKey = "";
		private FIA_Biosum_Manager.frmDialog  m_frmDialog;
		private FIA_Biosum_Manager.frmCoreScenario m_frmScenario;

		public txtNumeric(int intWholeMaxLength, int intDecimalMaxLength)
		{
		
			
			this.SelectionStart = 1;
			// + 2 is for the "$" and "." characters
			this.MaxLength = intWholeMaxLength + 
				Convert.ToInt32(intWholeMaxLength/3) +   //take care of commas
				intDecimalMaxLength                                       
				+ 1;    
            this.intWholeMaxLen = intWholeMaxLength;
			this.intDecimalMaxLen = intDecimalMaxLength;

			//
			// TODO: Add constructor logic here
			//
		}
		public txtNumeric(int intWholeMaxLength, int intDecimalMaxLength,ref FIA_Biosum_Manager.uc_scenario_costs p_uc)
			
		{
		    
			MessageBox.Show(p_uc.ParentForm.Text);
			
			this.SelectionStart = 1;
			// + 2 is for the "$" and "." characters
			this.MaxLength = intWholeMaxLength + 
				Convert.ToInt32(intWholeMaxLength/3) +   //take care of commas
				intDecimalMaxLength                                       
				+ 2;    
			this.intWholeMaxLen = intWholeMaxLength;
			this.intDecimalMaxLen = intDecimalMaxLength;

			//
			// TODO: Add constructor logic here
			//
		}
		public FIA_Biosum_Manager.frmCoreScenario p_frmScenario
		{
			set 
			{
				this.m_frmScenario = value;
			}
		}
		public FIA_Biosum_Manager.frmDialog p_frmDialog
		{
			set
			{
			 	this.m_frmDialog = value;
			}
		}
		protected override void OnKeyPress(System.Windows.Forms.KeyPressEventArgs e)
		{
			this.strLastKey="";
			this.bEdit=true;
			this.AllowNumericOnly(e);
			//base.OnKeyPress (e);
		}
		protected void AllowNumericOnly(System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				//if (this.SelectionStart == 0) 
				//{
				//	e.Handled=true;
				//}
				//else
				//{
					
				if ((this.intWholeCurLen==this.intWholeMaxLen &&
					this.SelectionStart <= 
					this.Text.IndexOf(".",0) && this.intDecimalMaxLen > 0) ||
					(this.intWholeCurLen==this.intWholeMaxLen && this.intDecimalMaxLen==this.intDecimalCurLen))
				{
					e.Handled=true;
				}
				else 
				{

					if (this.intDecimalCurLen==this.intDecimalMaxLen &&  
						this.SelectionStart > 
						this.Text.IndexOf(".",0) && this.intDecimalMaxLen > 0)
					{
						e.Handled=true;
					}
					else
					{
						this.strLastKey = Convert.ToString(e.KeyChar);
						//if (this.m_frmScenario.btnSave.Enabled==false) this.m_frmScenario.btnSave.Enabled=true;
						// Digits are OK
						//if (((FIA_Biosum_Manager.frmScenario)this.Parent.Parent).btnSave.Enabled == false) 
						//	((FIA_Biosum_Manager.frmScenario)this.Parent.Parent).btnSave.Enabled=true;
					}
						
				}
				//}
			}
			else if (e.KeyChar == '\b')
			{
				if (this.SelectionStart == 0) 
				{
					e.Handled=true;
				}
				else 
				{
					if (this.Text.Substring(this.SelectionStart - 1,1) != "." &&
						this.Text.Substring(this.SelectionStart - 1,1) != "$")
					{
						//if (this.m_frmScenario.btnSave.Enabled==false) this.m_frmScenario.btnSave.Enabled=true;
						//back space is okay
						//if (((frmScenario)this.ParentForm).btnSave.Enabled == false) 
						//	((frmScenario)this.ParentForm).btnSave.Enabled=true;

						this.strLastKey="b";

					
					}
					else 
					{
						if (this.Text.Substring(this.SelectionStart - 1,1) == ".")
							this.SelectionStart = this.SelectionStart - 1;
						e.Handled=true;
					}
				}
			
				
			}
			else if (e.KeyChar == '.')
			{
				if (this.Text.IndexOf(".",0) >= 0)
				{
					e.Handled=true;
				}
			}
			else
			{

				e.Handled = true;
			}
		}
		protected override void OnTextChanged(EventArgs e)
		{
			this.FormatString();

			

			//base.OnTextChanged (e);
		}
		public void FormatString()
		{
			int x;
			int intWholeCount=0;
			int intDecimalCount=0;
			int intCommaCount=0;
			//int intDecimalCount = 0;
			//bool bCountWhole = true;
			bool bWholeEdit = true;
			string strWhole="";
			string strWholeNoComma="";
			string strDecimal="";
			int intCursorMove = 0;
			string strNewWhole="";
			string strNewWholeNoComma="";
			//string strNewDecimal="";
			//bool bNewChanges=false;
			int intRemainder=0;
			int intSelection=this.SelectionStart;

			if (bEdit == false) 
			{
				bEdit = true;
				return;
			}
            

			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			

			//make sure the Whole sign is always a part of the string
			//if (this.Text.IndexOf("$",0) < 0)
			//{
				
			//	this.bEdit=false;
			//	this.Text = "$" + this.Text;
			//	this.SelectionStart = intSelection;
			//}

			//make sure the decimal is always a part of the string
			
			if (this.Text.IndexOf(".",0) < 0 && this.intDecimalMaxLen > 0)
			{
				this.bEdit=false;
				//check the cursor position
				if (this.Text.Trim().Length - 1 >= intSelection)
				{
					//user deleted the "." and the string had Decimal values
					this.Text = 
						this.Text.Substring(0,this.SelectionStart) + "." + 
						this.Text.Substring(this.SelectionStart,this.Text.Length - this.SelectionStart);
				}
				else 
				{
					//user deleted the "." and the string did not have Decimal values
					this.Text = 
						this.Text.Substring(0,this.SelectionStart) + ".";
				}
				
				this.SelectionStart = intSelection;
			}

			//determine whether Wholes or Decimals are being edited
			if (intSelection <= this.Text.IndexOf(".",0))
			{
				bWholeEdit = true;
			}
			else
			{
				bWholeEdit = false;
			}

			if (this.intDecimalMaxLen > 0)
			{
				//get the formatted Whole string without the "$" sign
				strWhole = this.Text.Substring(0,this.Text.IndexOf(".",0));
			}
			else
			{
				strWhole = this.Text;
			}

			//get the Whole string with only numbers
			strWholeNoComma = strWhole.Replace(",","");

			//check to see if there are any Decimals
			if (this.intDecimalMaxLen > 0)
			{
				if (this.Text.Trim().Length > (this.Text.IndexOf(".",0)+1))
				{
					//get the Decimal string
					strDecimal= this.Text.Substring(this.Text.IndexOf(".",0)+1,this.Text.Trim().Length - this.Text.IndexOf(".",0)-1);
				}
			}
			else
			{
				strDecimal="";
			}
			//get the number of Whole digits
			intWholeCount = strWholeNoComma.Length;
			//get the number of Decimal digits
			intDecimalCount = strDecimal.Length;
             
			//load the values in the shared variables
			this.intWholeCurLen = intWholeCount;
			this.intDecimalCurLen = intDecimalCount;

			//get the number of commas in the formatted Whole string
			intCommaCount = strWhole.Trim().Length - intWholeCount;

			//format the new Whole string
			for (x=intWholeCount-1; x>=0;x--)
			{
				if (strNewWhole.Trim().Length > 0)
				{
					//check to see if a comma needs to be added
					System.Math.DivRem(strNewWholeNoComma.Trim().Length,3,out intRemainder);
					if (intRemainder == 0) strNewWhole = "," + strNewWhole;
				}
				strNewWhole = strWholeNoComma.Substring(x,1) + strNewWhole;
				strNewWholeNoComma = strWholeNoComma.Substring(x,1) + strNewWholeNoComma;
			}
			//MessageBox.Show(strNewWhole + " " + strWhole + " " + strDecimal);
			//check to see if the new formatted string is different than what is currently
			//in the text field
			if (strNewWhole != strWhole)
			{
				if (this.bInitialize==false)
				{
//					if (this.m_frmScenario.btnSave.Enabled==false) this.m_frmScenario.btnSave.Enabled=true;
				}
				//change the value of the text field to the new formatted string
				this.bEdit = false;
				if (this.intDecimalMaxLen > 0)
				{
					this.Text = strNewWhole + "." + strDecimal;
				}
				else 
				{
					this.Text = strNewWhole;
				}
			}

			//position the cursor based on the last key pressed and its current position
			switch (this.strLastKey.Trim())
			{
				case "":
					intCursorMove =  strNewWhole.Trim().Length - strWhole.Trim().Length;
					this.SelectionStart = intSelection + intCursorMove;
					break;
				case "b":
					if (intSelection != 1)
						intCursorMove =  strNewWhole.Trim().Length - strWhole.Trim().Length;
					this.SelectionStart = intSelection + intCursorMove;
					break;
				default:
					intCursorMove =  strNewWhole.Trim().Length - strWhole.Trim().Length;
					if (bWholeEdit && intWholeCount == this.intWholeMaxLen && (strWhole.Trim().Length == intSelection))
					{
						this.SelectionStart = intSelection + intCursorMove + 1;
					}
					else 
					{
						this.SelectionStart = intSelection + intCursorMove;
					}
					break;

			}
			this.strLastKey="";

		}
		public void InitializeText()
		{
			this.bInitialize = true;
			this.SelectionStart=1;
			this.SelectionLength = 0;

			
			
			if (this.Text == null || this.Text.Length ==0)
			{
				this.bEdit = false;
				this.Text = "$.00";
				this.intDecimalCurLen = 2;
				this.intWholeCurLen=0;
			}
			else 
			{
				this.bEdit=false;
				this.Text.Replace("$","");
				if (this.Text == "0") 
				{
					this.bEdit = false;
					this.Text = "$.00";
					this.intDecimalCurLen = 2;
					this.intWholeCurLen=0;
				}
				else if (Convert.ToDouble(this.Text) == 0)
				{
					this.bEdit = false;
					this.Text = "$.00";
					this.intDecimalCurLen = 2;
					this.intWholeCurLen=0;
				}
				else
				{
					this.bEdit = false;
					this.Text = "$" + this.Text;
					if (this.Text.IndexOf(".",0) < 0)
					{
						this.bEdit=false;
						this.Text = this.Text + ".00";
					}
					this.FormatString();
					
				}
			}
			this.bInitialize=false;
			
		}
		public void InitializeText(string strNumber)
		{
			//string strBuild1="";
			//string strBuild2="";
			if (strNumber.Trim().Length==0)
			{
				if (this.intDecimalMaxLen > 0)
				{
					this.Text = ".";
				}
				return;
			}
			if (strNumber.IndexOf(".") < 0)
			{
				this.Text = strNumber + ".";
				return;
			}  
		}
		

	}
}
