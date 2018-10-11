using System;
using System.Windows.Forms;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for txtDollarsAndCents.
	/// </summary>
	/// 
	public class txtDollarsAndCents:System.Windows.Forms.TextBox
	{
		public bool bInitialize=true;
		public bool bEdit = false;
		private int intDollarMaxLen=4;
		private int intCentMaxLen=2;
		private int intDollarCurLen=0;
		private int intCentCurLen=0;
		private string strLastKey = "";
		private FIA_Biosum_Manager.frmOptimizerScenario _frmCoreScenario;

		public txtDollarsAndCents(int intDollarMaxLength, int intCentMaxLength)
		{
		
			
			this.SelectionStart = 1;
			// + 2 is for the "$" and "." characters
			this.MaxLength = intDollarMaxLength + 
				Convert.ToInt32(intDollarMaxLength/3) +   //take care of commas
				intCentMaxLength                                       
				+ 2;    
            this.intDollarMaxLen = intDollarMaxLength;
			this.intCentMaxLen = intCentMaxLength;

			//
			// TODO: Add constructor logic here
			//
		}
		public txtDollarsAndCents(int intDollarMaxLength, int intCentMaxLength,ref FIA_Biosum_Manager.uc_optimizer_scenario_costs p_uc)
			
		{
		    
			MessageBox.Show(p_uc.ParentForm.Text);
			
			this.SelectionStart = 1;
			// + 2 is for the "$" and "." characters
			this.MaxLength = intDollarMaxLength + 
				Convert.ToInt32(intDollarMaxLength/3) +   //take care of commas
				intCentMaxLength                                       
				+ 2;    
			this.intDollarMaxLen = intDollarMaxLength;
			this.intCentMaxLen = intCentMaxLength;

			//
			// TODO: Add constructor logic here
			//
		}
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceCoreScenarioForm
		{
			set 
			{
				_frmCoreScenario = value;
			}
			get
			{
				return _frmCoreScenario;
			}
		}
		protected override void OnKeyPress(System.Windows.Forms.KeyPressEventArgs e)
		{
			this.strLastKey="";
			this.bEdit=true;
			this.AllowNumericOnly(e);
		}
		protected void AllowNumericOnly(System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				if (this.SelectionStart == 0) 
				{
					e.Handled=true;
				}
				else
				{
					
					if (this.intDollarCurLen==this.intDollarMaxLen &&
						this.SelectionStart <= 
						this.Text.IndexOf(".",0))
					{
						e.Handled=true;
					}
					else 
					{
						if (this.intCentCurLen==this.intCentMaxLen &&  
							this.SelectionStart > 
							this.Text.IndexOf(".",0))
						{
							e.Handled=true;
						}
						else
						{
							this.strLastKey = Convert.ToString(e.KeyChar);
							try
							{
								if (ReferenceCoreScenarioForm != null)
									ReferenceCoreScenarioForm.m_bSave=true;
							}
							catch
							{
							}
							// Digits are OK
							
						}
						
					}
				}
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
						try
						{
							ReferenceCoreScenarioForm.m_bSave=true;
						}
						catch
						{
						}
						//back space is okay
						if (ReferenceCoreScenarioForm != null)
							ReferenceCoreScenarioForm.m_bSave=true;

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
			else
			{
				e.Handled = true;
			}
		}
		protected override void OnTextChanged(EventArgs e)
		{
			this.FormatString();

		}
		private void txtRailHaulCost_Enter(object sender, System.EventArgs e)
		{
		
		}
		public void FormatString()
		{
			int x;
			int intDollarCount=0;
			int intCentCount=0;
			int intCommaCount=0;
			bool bDollarEdit = true;
			string strDollar="";
			string strDollarNoComma="";
			string strCent="";
			int intCursorMove = 0;
			string strNewDollar="";
			string strNewDollarNoComma="";
			int intRemainder=0;
			int intSelection=this.SelectionStart;

			if (bEdit == false) 
			{
				bEdit = true;
				return;
			}
            


			//make sure the dollar sign is always a part of the string
			if (this.Text.IndexOf("$",0) < 0)
			{
				
				this.bEdit=false;
				this.Text = "$" + this.Text;
				this.SelectionStart = intSelection;
			}

			//make sure the decimal is always a part of the string
			if (this.Text.IndexOf(".",0) < 0)
			{
				this.bEdit=false;
				//check the cursor position
				if (this.Text.Trim().Length - 1 >= intSelection)
				{
					//user deleted the "." and the string had cent values
					this.Text = 
						this.Text.Substring(0,this.SelectionStart) + "." + 
						this.Text.Substring(this.SelectionStart,this.Text.Length - this.SelectionStart);
				}
				else 
				{
					//user deleted the "." and the string did not have cent values
					this.Text = 
						this.Text.Substring(0,this.SelectionStart) + ".";
				}
				
				this.SelectionStart = intSelection;
			}

			//determine whether dollars or cents are being edited
			if (intSelection <= this.Text.IndexOf(".",0))
			{
				bDollarEdit = true;
			}
			else
			{
				bDollarEdit = false;
			}

			//get the formatted dollar string without the "$" sign
			strDollar = this.Text.Substring(1,this.Text.IndexOf(".",0)-1);

			//get the dollar string with only numbers
			strDollarNoComma = strDollar.Replace(",","");

			//check to see if there are any cents
			if (this.Text.Trim().Length > (this.Text.IndexOf(".",0) + 1))
			{
				//get the cent string
				strCent= this.Text.Substring(this.Text.IndexOf(".",0)+1,this.Text.Trim().Length - this.Text.IndexOf(".",0)-1);
			}
            
			//get the number of dollar digits
			intDollarCount = strDollarNoComma.Length;
			//get the number of cent digits
			intCentCount = strCent.Length;
             
			//load the values in the shared variables
			this.intDollarCurLen = intDollarCount;
			this.intCentCurLen = intCentCount;

			//get the number of commas in the formatted dollar string
			intCommaCount = strDollar.Trim().Length - intDollarCount;

			//format the new dollar string
			for (x=intDollarCount-1; x>=0;x--)
			{
				if (strNewDollar.Trim().Length > 0)
				{
					//check to see if a comma needs to be added
					System.Math.DivRem(strNewDollarNoComma.Trim().Length,3,out intRemainder);
					if (intRemainder == 0) strNewDollar = "," + strNewDollar;
				}
				strNewDollar = strDollarNoComma.Substring(x,1) + strNewDollar;
				strNewDollarNoComma = strDollarNoComma.Substring(x,1) + strNewDollarNoComma;
			}
			//MessageBox.Show(strNewDollar + " " + strDollar);
			//check to see if the new formatted string is different than what is currently
			//in the text field
			if (strNewDollar != strDollar)
			{
				if (this.bInitialize==false)
				{
					//if (this.m_frmScenario.btnSave.Enabled==false) this.m_frmScenario.btnSave.Enabled=true;
					if (ReferenceCoreScenarioForm != null)
							ReferenceCoreScenarioForm.m_bSave=true;
				}
				//change the value of the text field to the new formatted string
				this.bEdit = false;
				this.Text = "$" + strNewDollar + "." + strCent;
			}

			//position the cursor based on the last key pressed and its current position
			switch (this.strLastKey.Trim())
			{
				case "":
					intCursorMove =  strNewDollar.Trim().Length - strDollar.Trim().Length;
					this.SelectionStart = intSelection + intCursorMove;
					break;
				case "b":
					if (intSelection != 1)
						intCursorMove =  strNewDollar.Trim().Length - strDollar.Trim().Length;
					this.SelectionStart = intSelection + intCursorMove;
					break;
				default:
					intCursorMove =  strNewDollar.Trim().Length - strDollar.Trim().Length;
					if (bDollarEdit && intDollarCount == this.intDollarMaxLen && (strDollar.Trim().Length + 1 == intSelection))
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
				this.intCentCurLen = 2;
				this.intDollarCurLen=0;
			}
			else 
			{
				this.bEdit=false;
				this.Text.Replace("$","");
				if (this.Text == "0") 
				{
					this.bEdit = false;
					this.Text = "$.00";
					this.intCentCurLen = 2;
					this.intDollarCurLen=0;
				}
				else if (Convert.ToDouble(this.Text) == 0)
				{
					this.bEdit = false;
					this.Text = "$.00";
					this.intCentCurLen = 2;
					this.intDollarCurLen=0;
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

	}
}
