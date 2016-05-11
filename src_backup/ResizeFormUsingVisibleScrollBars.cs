using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Resize the client form until the Horizontal and Vertical scroll bars are no longer visible
	/// </summary>
	public class ResizeFormUsingVisibleScrollBars
	{
		private ScrollBars _visibleScrollbars = ScrollBars.None;
		public event EventHandler VisibleScrollbarsChanged;
		private Control _oControlToResize=null;
		private Control _oScrollBarParentControl=null;
		private int _MaxHt=800;
		private int _MaxWd=800;
		private bool _bResizeWd=true;
		private bool _bResizeHt=true;

		public ResizeFormUsingVisibleScrollBars()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		public int MaximumHeight
		{
			get {return this._MaxHt;}
			set {this._MaxHt = value;}
		}
		public int MaximumWidth
		{
			get {return this._MaxWd;}
			set {this._MaxWd = value;}
		}
		public bool ResizeHeight
		{
			get {return this._bResizeHt;}
			set {this._bResizeHt=value;}
		}
		public bool ResizeWidth
		{
			get {return this._bResizeWd;}
			set {this._bResizeWd=value;}
		}
		public Control ControlToResize
		{
			set {_oControlToResize=value;}
			get {return _oControlToResize;}
		}
		public Control ScrollBarParentControl
		{
			set 
			{
				this._oScrollBarParentControl = value;
			    this._oScrollBarParentControl.Resize += new System.EventHandler(this.ScrollBarParentControl_Resize);
			}
			get {return this._oScrollBarParentControl;}
		}
		public void ResizeControl()
		{
			try
			{
				int x=ControlToResize.Width,y=ControlToResize.Height;

				for (;;)
				{
                    bool bAdjustmentMade = false;
					if (this.ResizeHeight==false && this.ResizeWidth==false) break;
					if ((x >= this.MaximumWidth ||  this.ResizeWidth==false) && 
						(y >= this.MaximumHeight || this.ResizeHeight==false)) break;
					if (VisibleScrollbars==System.Windows.Forms.ScrollBars.Both)
					{
						if (this.ResizeWidth)
						{
							if (x < this.MaximumWidth)
							{
								x=x + 10;
								ControlToResize.Width = x;
                                bAdjustmentMade = true;
							}
						}
						if (this.ResizeHeight)
						{
							if (y < this.MaximumHeight)
							{
								y=y+10;
								ControlToResize.Height = y + 10;
                                bAdjustmentMade = true;
							}
						}
						
					}
					else if (VisibleScrollbars==System.Windows.Forms.ScrollBars.Horizontal && 
						this.ResizeWidth)
					{
						if (x < this.MaximumWidth)
						{
							x=x + 10;
							ControlToResize.Width = x;
                            bAdjustmentMade = true;
						}
						
					}
					else if (VisibleScrollbars==System.Windows.Forms.ScrollBars.Vertical &&
						this.ResizeHeight)
					{
						if (y < this.MaximumHeight)
						{
							y=y+10;
							ControlToResize.Height = y + 10;
                            bAdjustmentMade = false;
						}
					}
                    if (bAdjustmentMade == false) break;
                    else break;
				
				}
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message);
			}
			
		}
		

		private ScrollBars VisibleScrollbars
		{
			get { return _visibleScrollbars; }
			set
			{
				if (_visibleScrollbars != value)
				{
					_visibleScrollbars = value;
					OnVisibleScrollbarsChanged(EventArgs.Empty);
				}
			}
		}

		
		protected virtual void OnVisibleScrollbarsChanged(EventArgs e)
		{
			if (VisibleScrollbarsChanged != null)
				VisibleScrollbarsChanged(this, e);
		}
	    private ScrollBars GetVisibleScrollbars(Control ctl)
		{
			int wndStyle = Win32.GetWindowLong(ctl.Handle, Win32.GWL_STYLE);
			bool hsVisible = (wndStyle & Win32.WS_HSCROLL) != 0;
			bool vsVisible = (wndStyle & Win32.WS_VSCROLL) != 0;

			if (hsVisible)
				return vsVisible ? ScrollBars.Both : ScrollBars.Horizontal;
			else
				return vsVisible ? ScrollBars.Vertical : ScrollBars.None;
		}
		protected void ScrollBarParentControl_Resize(object sender, System.EventArgs e)

		{
			//base.OnResize(e);
			VisibleScrollbars = GetVisibleScrollbars(ScrollBarParentControl);
		}


	}
	public class Win32
	{
		// offset of window style value
		public const int GWL_STYLE = -16;

		// window style constants for scrollbars
		public const int WS_VSCROLL = 0x00200000;
		public const int WS_HSCROLL = 0x00100000;

		[DllImport("user32.dll", SetLastError = true)]
		public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
	}
}
