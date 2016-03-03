using System;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for run_core_scenario.
	/// </summary>
	public class run_core_scenario
	{
		public FIA_Biosum_Manager.frmCoreScenario m_frmScenario;
		public FIA_Biosum_Manager.frmMain m_frmMain;
		public FIA_Biosum_Manager.frmRunCoreScenario m_frmRunCoreScenario;
		
		private bool m_bAbortProcess=false;
		
		public run_core_scenario(FIA_Biosum_Manager.frmMain p_frmMain, FIA_Biosum_Manager.frmCoreScenario p_frmScenario, FIA_Biosum_Manager.frmRunCoreScenario p_frmRunCoreScenario)
		{
			this.m_frmMain = p_frmMain;
			this.m_frmScenario = p_frmScenario;
			this.m_frmRunCoreScenario = p_frmRunCoreScenario;

			while (this.m_bAbortProcess==false)
			{
				System.Windows.Forms.Application.DoEvents();
			}
		    MessageBox.Show("out of while loop");


			
		}

		~run_core_scenario()
		{
		}
		
		public bool AbortProcess
		{
			get
			{
				return this.m_bAbortProcess;
			}
			set
			{
				this.m_bAbortProcess = value;
				
			}
		}
	}
}
