using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Windows.Forms;
using System.Data;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for printing.
	/// </summary>
	public class printing
	{
		private PrintDocument m_printDoc;
		private PrintDocument m_printReportDoc;
		private PrintDocument m_printReportGroupDoc;
		private System.Windows.Forms.FontDialog m_dlgFont;
		private System.Windows.Forms.PrintDialog m_dlgPrint;
		private System.Windows.Forms.PrintPreviewDialog m_dlgPrintPreview;
		private System.Drawing.Font m_fontReportTitle;
		private System.Drawing.Font m_fontHeader;
		private System.Drawing.Font m_fontColumns;
		private System.Drawing.Font m_fontBody;
		private System.Drawing.Font m_fontReportDate;
		private System.Drawing.Font m_fontSummary;
		private System.Drawing.Font m_fontSummaryHeader;
		private System.Drawing.Font m_fontSummaryColumn;
		private System.Drawing.Font m_fontPageNumber;
		private System.Drawing.Font m_fontEndOfReport;

		private string m_strPrintType="";
		private string m_strHeader="";
		private System.Windows.Forms.ListBox m_oListBox;
		private System.Data.DataTable m_dt;
		private System.Data.DataView m_dv;

		private System.Data.DataSet m_dsReportConfig;
		private int m_intPageNumber=0;
		
		private int m_intCurrRow=0;
		private int m_intCurrSummaryRow=0;
		private int m_intCurrSummaryRow2=0;
		private int m_intCurrRow2 = 0;
		private int m_intCurrCol=0;
		private string m_strCurrGroup="";
		
		private string m_strNewGroup="";
		private string m_strGroup1Value="";
		private string m_strGroup2Value="";
		private string m_strGroup3Value="";
		private int m_intReportSummaryLineTotal=0;
		
		private int m_intGroupSummaryLineTotal=0;
		
		
		private bool m_bNewGroup=true;
		private bool m_bPrintingReportSummary=false;
		private bool m_bPrintingGroupSummary=false;
		private bool m_bSummaryDoneButNewPage=false;
		private bool m_bDetailLineNewPage=false;
		private int m_intCounter=0;

		private FIA_Biosum_Manager.frmTherm m_frmTherm;
			
		//private System.Drawing m_draw;

		public printing()
		{
			this.m_printDoc = new PrintDocument();
			this.m_printReportDoc = new PrintDocument();
			this.m_printReportGroupDoc = new PrintDocument();
			this.m_dlgFont = new FontDialog();
			this.m_dlgPrint = new PrintDialog();
			this.m_dlgPrintPreview = new PrintPreviewDialog();
			this.m_frmTherm = new frmTherm();
			this.m_frmTherm.btnCancel.Click += new System.EventHandler(this.ThermCancel);
			this.m_frmTherm.Hide();

            
			
			
			

			this.m_printDoc.PrintPage += new PrintPageEventHandler(printDoc_PrintPage);
			this.m_printReportDoc.PrintPage += new PrintPageEventHandler(printReportDoc_PrintPage);
			this.m_printReportGroupDoc.PrintPage += new PrintPageEventHandler(printReportGroupDoc_PrintPage);
			//
			// TODO: Add constructor logic here
			//


		}
		private void printDoc_PrintPage(Object sender, PrintPageEventArgs e)
		{
			float headerHeight;
			float linesPerPage;
			float yPosition;
			//float count;
			switch (this.m_strPrintType)
			{
				case "LISTBOX":
					

					this.m_fontHeader = new Font("Courier New", 12, System.Drawing.FontStyle.Bold);
					this.m_fontBody = this.m_oListBox.Font;

					// Determine the height of the header, based on the selected font.
					headerHeight = this.m_fontHeader.GetHeight(e.Graphics);

					// Determine the number of lines of body text that can be printed per page, taking
					// into account the presence of the header and the size of the selected body font.
					linesPerPage = (e.MarginBounds.Height - headerHeight)/ this.m_fontBody.GetHeight(e.Graphics);


					// Used to store the position at which the next body line
					// should be printed.
					yPosition = 0;

					// Used to store the number of lines printed so far on the
					// current page.
					//count = 0;
					// Print the page header, as specified by the user in the form.
					// Use the header font for this line only.
					e.Graphics.DrawString(this.m_strHeader, this.m_fontHeader, Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top, new StringFormat());

					for (int x = 0; x <= this.m_oListBox.Items.Count-1;x++)
					{
						

						// Determine the position at which to print.  Since this report prints one line
						// at a time, only the height (or Y coordinate) needs to be calculated, because
						// every line will begin at the far left.  The Y coordinate must take into 
						// consideration the header, the height of each line of body text, and the 
						// number of body lines printed so far on this page.
						yPosition = e.MarginBounds.Top + headerHeight + ((x+1) * this.m_fontBody.GetHeight(e.Graphics));

						// Draw the line of text on the page using the body font specified by the user.
						//e.Graphics.DrawString(line, bodyFont, Brushes.Black, e.MarginBounds.Left, yPosition, new StringFormat());
						e.Graphics.DrawString(m_oListBox.Items[x].ToString(), this.m_fontBody,Brushes.Black,e.MarginBounds.Left,yPosition, new StringFormat());
						
					}
					e.HasMorePages = false;


					break;
			}
            





             
		}

		private void printReportDoc_PrintPage(Object sender, PrintPageEventArgs e)
		{
			float titleHeight;     //title height
			float dateHeight;      //date height
			float pagenumHeight;   //page number height
			float linesPerPage;    
			float columnHeight;    //column name height
			float bodyHeight;      //body of the report
			float SummaryHeight;
			float SummaryHeightColumn;
			float ypos;
			float xpos;
			float linecount=0;
			int x=0;
			int y=0;
			int z=0;
			int intWidth=0;
			int intWidth2 = 0;
			float floatWidth=0;
			float floatWidth2=0;
			
			//string strValue2="";
			System.Drawing.SizeF sizeFWidth;
			System.Drawing.SizeF sizeFWidth2;
			switch (this.m_strPrintType)
			{
				case "REPORT":
					
					Pen pen = new Pen(System.Drawing.Color.Black);

					// Used to store the number of record detail lines printed so far on the
					// current page.
					linecount=0;
					

					// Determine the height of the elements in the report
					titleHeight = this.m_fontReportTitle.GetHeight(e.Graphics);
					pagenumHeight = this.m_fontPageNumber.GetHeight(e.Graphics) * 2;
					dateHeight = this.m_fontReportDate.GetHeight(e.Graphics);
					columnHeight = this.m_fontColumns.GetHeight(e.Graphics);
					bodyHeight = this.m_fontBody.GetHeight(e.Graphics);
					SummaryHeight = this.m_fontSummary.GetHeight(e.Graphics);
					SummaryHeightColumn = this.m_fontSummaryColumn.GetHeight(e.Graphics);

					// Used to store the position at which the next body line
					// should be printed.
					ypos = e.MarginBounds.Top;
					xpos = e.MarginBounds.Left;

					//print a line 
					e.Graphics.DrawLine(pen, e.MarginBounds.Left, e.MarginBounds.Top,
						e.MarginBounds.Right, e.MarginBounds.Top);

					// Determine the number of lines of body text that can be printed per page, taking
					// into account the presence of the header and the size of the selected body font.
					if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["type"]=="C")
					{
						//each column is printed on the same line
						if (this.m_intPageNumber == 0)
						{
							//1st page of report so print the report header
							linesPerPage = (e.MarginBounds.Height - titleHeight - pagenumHeight-dateHeight-columnHeight)/ bodyHeight;
							//print the report title
							e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_title"].Rows[0]["title"].ToString(), this.m_fontReportTitle, Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top, new StringFormat());
							ypos = ypos + titleHeight; // +  this.m_fontBody.GetHeight(e.Graphics);
							//print the date
							e.Graphics.DrawString(System.DateTime.Now.ToShortDateString(), this.m_fontReportDate, Brushes.Black, e.MarginBounds.Left, ypos, new StringFormat());
							ypos = ypos + pagenumHeight;
							//draw a line
							e.Graphics.DrawLine(pen, e.MarginBounds.Left, ypos,
								e.MarginBounds.Right, ypos);

						}
						else
						{
							linesPerPage = (e.MarginBounds.Height - pagenumHeight-columnHeight)/ this.m_fontBody.GetHeight(e.Graphics);
						}

						//print the column header
						for (x=0; x<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;x++)
						{
							//get the width of the columns longest value
							sizeFWidth = e.Graphics.MeasureString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["largeststring"].ToString().Trim() + "***", this.m_fontColumns);
							//get the width of the columns name
							sizeFWidth2 = e.Graphics.MeasureString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim() + "***", this.m_fontColumns);
							intWidth = (int)sizeFWidth.Width;
							intWidth2 = (int)sizeFWidth2.Width;
							floatWidth = (float)sizeFWidth.Width;
							floatWidth2 = (float)sizeFWidth2.Width;

							//see if column width is wider than the widest column value
							//if (intWidth2 > intWidth)
							//if the column name width is greater than the longest value than use the column name width
							if (floatWidth2 > floatWidth)
							{
								//if this column name will print beyond the margin 
								//than drop this column from the report
								//if (xpos + intWidth2 <= e.MarginBounds.Width)
								if (xpos + floatWidth2 <= e.MarginBounds.Width)
								{
									//this is within the margins so lets keep this column
									//assign the beginning position of the column
									this.m_dsReportConfig.Tables["report_fields"].Rows[x]["Xpos"] = xpos;
									//print the column name
									e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim(), this.m_fontColumns, Brushes.Black, xpos, ypos, new StringFormat());
									//xpos = xpos + intWidth2;
									xpos = xpos + floatWidth2;
								
								}
								else 
								{
									break;
								}

							}
							else
							{
								//if the width of the longest value in the column
								//is greater than the margin width than drop
								//this column from the report
								//if (xpos + intWidth <= e.MarginBounds.Width)
								if (xpos + floatWidth <= e.MarginBounds.Width)
								{
									//this is within the margins so lets keep this column
									//assign the beginning position of the column
									this.m_dsReportConfig.Tables["report_fields"].Rows[x]["Xpos"] = xpos;
									//print the column name
									e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim(), this.m_fontColumns, Brushes.Black, xpos, ypos, new StringFormat());
									//xpos = xpos + intWidth;
									xpos = xpos + floatWidth;
								
								}
								else 
								{
									break;
								}
							}
						
						}
						ypos = ypos + columnHeight + 2;
					}
					else
					{
						//each column name is printed on a separate line
						if (this.m_intPageNumber == 0)
						{
							linesPerPage = (e.MarginBounds.Height - titleHeight - pagenumHeight-dateHeight)/ bodyHeight;
							//print the report title
							e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_title"].Rows[0]["title"].ToString(), this.m_fontReportTitle, Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top, new StringFormat());
							ypos = ypos + titleHeight; // +  this.m_fontBody.GetHeight(e.Graphics);
							//print the date
							e.Graphics.DrawString(System.DateTime.Now.ToShortDateString(), this.m_fontReportDate, Brushes.Black, e.MarginBounds.Left, ypos, new StringFormat());
							ypos = ypos + pagenumHeight;
							//draw a line
							e.Graphics.DrawLine(pen, e.MarginBounds.Left, ypos,
								e.MarginBounds.Right, ypos);

						}
						else
						{
							linesPerPage = (e.MarginBounds.Height - pagenumHeight)/ bodyHeight;
						}
						//get the width of the longest column name
						sizeFWidth = e.Graphics.MeasureString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["largeststring"].ToString().Trim() + ":**", this.m_fontColumns);
						floatWidth = (float)sizeFWidth.Width;
						for (x=0;x <= this.m_dsReportConfig.Tables["Report_Fields"].Rows.Count-1;x++)
						{
							//assign the position for the column values to print
							this.m_dsReportConfig.Tables["Report_Fields"].Rows[x]["xpos"] = e.MarginBounds.Left + floatWidth;
						}


					}
					//draw a line
					e.Graphics.DrawLine(pen, e.MarginBounds.Left, ypos,
						e.MarginBounds.Right, ypos);

					this.m_intPageNumber++;

					//loop through each row in the dataview
					for (x=this.m_intCurrRow; x<=this.m_dv.Table.Rows.Count-1; x++)
					{
						//loop thru the report config data set and only
						//print the columns from the dataview that are in
						//the report config dataset
						for (y=this.m_intCurrRow2;y <= this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;y++)
						{
							//loop through the dataview columns
							for (z=0;z<=this.m_dv.Table.Columns.Count-1;z++)
							{
								//if the dataview column name equals the report config column name
								//then print the column
								if (this.m_dsReportConfig.Tables["report_fields"].Rows[y]["field_name"].ToString().Trim().ToUpper() ==
									this.m_dv.Table.Columns[z].ColumnName.ToString().Trim().ToUpper())
								{
									
									if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["type"] == "C")
									{
										//all columns on the same row
										if (this.m_dsReportConfig.Tables["report_fields"].Rows[y]["xpos"].ToString().Trim() != "-1")

										{
											if (this.m_dv.Table.Rows[x][z] != System.DBNull.Value)
											{
												//print the column value
												if ((string)this.m_dsReportConfig.Tables["report_fields"].Rows[y]["stringdatatype_YN"] == "N" ||
													(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[y]["candomath"] == true)
												{
													//right justify for numeric items
													System.Drawing.SizeF sizeFWidthLargestString = e.Graphics.MeasureString(this.m_dsReportConfig.Tables["report_fields"].Rows[y]["largeststring"].ToString().Trim(), this.m_fontBody);
													float floatWidthLargestString = (float)sizeFWidthLargestString.Width;
													string strValue = String.Format("{0:f3}",Math.Round(Convert.ToDecimal(this.m_dv.Table.Rows[x][z]),3));
													System.Drawing.SizeF fltValueF = e.Graphics.MeasureString(strValue, this.m_fontBody);
													float fltValue  = (float)fltValueF.Width;
													if (z==0)  //first column
													{
													
														e.Graphics.DrawString(strValue, this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"] + (float)((float)floatWidthLargestString - (float)fltValue),ypos, new StringFormat());                                         
													
													}
													else
													{
														e.Graphics.DrawString(strValue, this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"] + (float)((float)floatWidthLargestString - (float)fltValue),ypos, new StringFormat());                                         
														//e.Graphics.DrawString(strValue, this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"] + (float)floatWidth - fltValue,ypos, new StringFormat());                                         
													}
												}
												else
												{
													//left justify
													e.Graphics.DrawString(this.m_dv.Table.Rows[x][z].ToString().Trim(), this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"],ypos, new StringFormat());                                         
												}
											}
										}
									}
									else
									{
										//each column has its own row
										//print the column name
										e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_fields"].Rows[y]["field_name"].ToString().Trim() + ":  ", this.m_fontColumns, Brushes.Black, e.MarginBounds.Left, ypos, new StringFormat());
										//print the column value
										e.Graphics.DrawString(this.m_dv.Table.Rows[x][z].ToString().Trim(), this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"],ypos, new StringFormat());                                         
										linecount++;
										//check to see if we need to start a new page
										if (linecount >= linesPerPage)
										{
											//start a new page so save the current 
											//dataview column number and 
											//increment the report config 
											//record number to print the
											//next column
											this.m_intCurrCol=z;
											this.m_intCurrRow2 = y+1;
											break;
										}
										ypos = ypos + bodyHeight;
									}
								}
							}
							if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["type"] == "R")
							{
								//each column printed on a separate row
								if (linecount >= linesPerPage)
								{
									break;
								}
								else
								{
									//initialize to print all columns
									//this will be greater than zero for a
									//new page
									this.m_intCurrCol=0;
								}
							}
						}
						
						linecount++;
						ypos = ypos + bodyHeight;
						
						
						if (linecount >= linesPerPage)
						{
							//print the page in the middle of the page at the bottom
							sizeFWidth = e.Graphics.MeasureString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber);
							
							xpos = (e.PageBounds.Width/2)- (float)(sizeFWidth.Width * .50) ;
							ypos = e.PageBounds.Height - (pagenumHeight * 2);
							e.Graphics.DrawString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber, Brushes.Black, xpos, ypos, new StringFormat());
							if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["type"] == "C")
							{
								//increment by one to print the new row on the next page
								this.m_intCurrRow=x+1;
							}
							else 
							{
								//save the current row to print the same row but a new column on the 
								//next page
								this.m_intCurrRow = x;
							}
							e.HasMorePages = true;
							break;
						}
						else
						{
							this.m_intCurrRow2=0;
							e.HasMorePages = false;
						}

					}
					//see about printing the report summary lines
					if (x > this.m_dv.Table.Rows.Count-1)
					{
						this.m_intCurrRow = x;
						if (this.m_intReportSummaryLineTotal > 0)
						{
							PrintSummary_Style1(ref sender, ref e, ref ypos, bodyHeight, pagenumHeight,ref linecount, linesPerPage);
						}
					}
					//this takes care of the last page printed
					if (e.HasMorePages == false)
					{

						ypos = ypos + (bodyHeight * 2);
						e.Graphics.DrawString("**End Of Report**", this.m_fontEndOfReport,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
						
						sizeFWidth = e.Graphics.MeasureString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber);
						xpos = (e.PageBounds.Width/2)- (float)(sizeFWidth.Width * .50) ;
						ypos = e.PageBounds.Height - (pagenumHeight * 2);
						e.Graphics.DrawString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber, Brushes.Black, xpos, ypos, new StringFormat());
						this.m_intPageNumber=0;
						this.m_intCurrCol = 0;
						this.m_intCurrRow=0;
						this.m_intCurrRow2=0;
						
					}
					break;
			}  
		}
		private void PrintSummary_Style1(ref Object sender, 
			ref PrintPageEventArgs e, 
			ref float ypos, 
			float bodyHeight, 
			float pagenumHeight,
			ref float linecount,
			float linesPerPage)
		{
			float floatWidth=0;
			string str="";
			string strValue1="";
			string strValue2="";
			System.Drawing.SizeF sizeFWidth;
			int x=0;
			int y=0;
			float xpos;
           
			//comment
			ypos = ypos + bodyHeight;
			e.Graphics.DrawString("Report Summary", this.m_fontSummaryHeader,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
			linecount++;

			sizeFWidth = e.Graphics.MeasureString("Record Count And Pct:**", this.m_fontSummary);
			floatWidth = (float)sizeFWidth.Width;
			for (x=m_intCurrSummaryRow;x<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;x++)
			{
				for (y=m_intCurrSummaryRow2;y<=this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;y++)
				{
					if (this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim().ToUpper() == 
						this.m_dsReportConfig.Tables["report_results"].Rows[y]["field_name"].ToString().Trim().ToUpper())
					{
								  	
						ypos = ypos + bodyHeight;
						e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_results"].Rows[y]["field_name"].ToString().Trim(), this.m_fontSummaryColumn,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
						linecount++;
                                
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["candomath"] == true)
						{
							//min and max
							str="";
							strValue1="";
							strValue2="";
							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["min"] == true)
							{
								str="Min";
								strValue1 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["min_report_summary"].ToString().Trim();
								if (strValue1.IndexOf(".",0) >= 0 &&
									strValue1.Trim().Length > strValue1.IndexOf(".",0) + 3)
									strValue1 = strValue1.Substring(0,strValue1.IndexOf(".",0) + 3);

							}    

							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["max"] == true)
							{
								if (str.Trim()=="Min")
								{
									str += " and Max";
									strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["max_report_summary"].ToString().Trim();
									if (strValue2.IndexOf(".",0) >= 0 &&
										strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
										strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);

								}
								else 
								{
									str="Max";
									strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["max_report_summary"].ToString().Trim();
									if (strValue2.IndexOf(".",0) >= 0 &&
										strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
										strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);

								}
					  	  				
					  	  				
							} 
							if (str.Trim().Length > 0)
							{
                                                
								ypos = (ypos) + (bodyHeight);
								e.Graphics.DrawString(str, this.m_fontSummary,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
								xpos = (e.MarginBounds.Left)+ (float)(sizeFWidth.Width) ;
								e.Graphics.DrawString(strValue1, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
								if (strValue2.Trim().Length > 0)
								{
									xpos = xpos + (float)(sizeFWidth.Width) ;
									e.Graphics.DrawString(strValue2, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         

								}
								linecount++;
										
							}
							//sum and average
							str="";
							strValue1="";
							strValue2="";
							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["sum"] == true)
							{
								str="Sum";
								strValue1 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["sum_report_summary_total"].ToString().Trim();
								if (strValue1.IndexOf(".",0) >= 0 &&
									strValue1.Trim().Length > strValue1.IndexOf(".",0) + 3)
									strValue1 = strValue1.Substring(0,strValue1.IndexOf(".",0) + 3);

							}    

							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["avg_sum"] == true)
							{
								if (str.Trim()=="Sum")
								{
									str += " and Avg";
									strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["avg_sum_report_summary_total"].ToString().Trim();
									if (strValue2.IndexOf(".",0) >= 0 &&
										strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
										strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);

								}
								else 
								{
									str="Avg";
									strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["avg_sum_report_summary_total"].ToString().Trim();
									if (strValue2.IndexOf(".",0) >= 0 &&
										strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
										strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);

								}
										
					  	  				
							} 
							if (str.Trim().Length > 0)
							{
                                                
										
								ypos = (ypos) + (bodyHeight);
								e.Graphics.DrawString(str, this.m_fontSummary,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
								xpos = (e.MarginBounds.Left)+ (float)(sizeFWidth.Width) ;
								e.Graphics.DrawString(strValue1, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
								if (strValue2.Trim().Length > 0)
								{
									xpos = xpos + (float)(sizeFWidth.Width) ;
									e.Graphics.DrawString(strValue2, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         

								}
								linecount++;
							}

						}
						//count and percent
						str="";
						strValue1="";
						strValue2="";
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["count"] == true)
						{
							str="Count";
							strValue1 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["count_report_summary_total"].ToString().Trim();
							if (strValue1.IndexOf(".",0) >= 0 &&
								strValue1.Trim().Length > strValue1.IndexOf(".",0) + 3)
								strValue1 = strValue1.Substring(0,strValue1.IndexOf(".",0) + 3);

						}    

						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["pct_count"] == true)
						{
							if (str.Trim()=="Count")
							{
								str += " and Pct";
								strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["pct_count_report_summary_total"].ToString().Trim();
								//limit numeric values to 2 places after the decimal
								if (strValue2.IndexOf(".",0) >= 0 &&
									strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
									strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);
							}
							else 
							{
								str="Pct";
								strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["pct_count_report_summary_total"].ToString().Trim();
								if (strValue2.IndexOf(".",0) >= 0 &&
									strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
									strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);

							}
										
					  	  				
						} 
						if (str.Trim().Length > 0)
						{
                                                
										
							ypos = (ypos) + (bodyHeight);
							e.Graphics.DrawString(str, this.m_fontSummary,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
							xpos = (e.MarginBounds.Left)+ (float)(sizeFWidth.Width) ;
							e.Graphics.DrawString(strValue1, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
							if (strValue2.Trim().Length > 0)
							{
								xpos = xpos + (float)(sizeFWidth.Width) ;
								e.Graphics.DrawString(strValue2, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         

							}
							linecount++;
							if (linecount >= linesPerPage)
							{
								//start a new page so save the current 
								//dataview column number and 
								//increment the report config 
								//record number to print the
								//next column
								this.m_intCurrSummaryRow=x;
								this.m_intCurrSummaryRow2 = y;
								break;
							}
						}
						break;
					}
				}
				if (linecount >= linesPerPage)
				{
					//print the page in the middle of the page at the bottom
					sizeFWidth = e.Graphics.MeasureString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber);
					this.m_bPrintingReportSummary=true;
					xpos = (e.PageBounds.Width/2)- (float)(sizeFWidth.Width * .50) ;
					ypos = e.PageBounds.Height - (pagenumHeight * 2);
					e.Graphics.DrawString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber, Brushes.Black, xpos, ypos, new StringFormat());
					e.HasMorePages=true;
								  
					break;
				}

			}
			if (e.HasMorePages==false)
			{
				this.m_bPrintingReportSummary=false;
				ypos = ypos + (bodyHeight * 2);
				str= "Total Records: " + this.m_dsReportConfig.Tables["report_results"].Rows[0]["count_report_summary_record_total"].ToString();
				e.Graphics.DrawString(str, this.m_fontSummary,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
				linecount++;
			}
			
		}
    
		private void PrintGroupSummary_Style1(ref Object sender, 
			ref PrintPageEventArgs e, 
			ref float ypos, 
			float bodyHeight, 
			float pagenumHeight,
			ref float linecount,
			float linesPerPage)
			
		{
			float floatWidth=0;
			string str="";
			string strValue1="";
			string strValue2="";
			System.Drawing.SizeF sizeFWidth;
			int x=0;
			int y=0;
			float xpos;
			int intFirstGroupRow=-1;
			int intLastGroupRow=0;
			//bool bPrintRecordCount = false;
			float linecount2=0;

			this.m_bSummaryDoneButNewPage = false;
			if (linecount + 4 >= linesPerPage)
				linecount += 4;
			else
			{        

				/*********************************************************
				 **get the first and last record of the current group
				 *********************************************************/
				for(x=0;x<=this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;x++)
				{
					str=this.m_dsReportConfig.Tables["report_results"].Rows[x]["group_field1_value"].ToString().Trim() + 
						this.m_dsReportConfig.Tables["report_results"].Rows[x]["group_field2_value"].ToString().Trim() + 
						this.m_dsReportConfig.Tables["report_results"].Rows[x]["group_field3_value"].ToString().Trim();
					if (str.Trim().ToUpper() == this.m_strCurrGroup.Trim().ToUpper())
					{
						if (intFirstGroupRow==-1)
						{
							intFirstGroupRow = x;
							intLastGroupRow=x;
						}
						else 
						{
							intLastGroupRow = x;
						}
					}
					else
					{
						if (intLastGroupRow > 0) break;
					
 				  }
        }
			
				if (this.m_intCurrSummaryRow2 == 0) this.m_intCurrSummaryRow2 = intFirstGroupRow;

				/***************************************************************************
				 **lets return if we have already printed the group summary but because
				 ** of a new page event the group summary tries to print again
				 ***************************************************************************/
				if (m_intCurrSummaryRow2 > intLastGroupRow) 
				{
					x = this.m_dsReportConfig.Tables["report_results"].Rows.Count;
					this.m_intCurrSummaryRow=0;
					this.m_intCurrSummaryRow2=0;
					this.m_bPrintingGroupSummary=false;
					return;
			     
				}
    	
				ypos = ypos + bodyHeight;
				e.Graphics.DrawString("Group Summary", this.m_fontSummaryHeader,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
				linecount++;
                linecount2=linecount;
				sizeFWidth = e.Graphics.MeasureString("Record Count And Pct:**", this.m_fontSummary);
				floatWidth = (float)sizeFWidth.Width;
				for (x=m_intCurrSummaryRow;x<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;x++)
				{
					for (y=m_intCurrSummaryRow2;y<=intLastGroupRow;y++)
					{
						if (this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim().ToUpper() == 
							this.m_dsReportConfig.Tables["report_results"].Rows[y]["field_name"].ToString().Trim().ToUpper())
						{
							
							ypos = ypos + bodyHeight;
							e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_results"].Rows[y]["field_name"].ToString().Trim(), this.m_fontSummaryColumn,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
							linecount++;
                                
							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["candomath"] == true)
							{
								//min and max
								str="";
								strValue1="";
								strValue2="";
								//if (this.m_strGroupPrint.IndexOf("minmax",0) < 0)
								//{
								if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["min"] == true)
								{
									str="Min";
									strValue1 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["min_group_summary"].ToString().Trim();
									if (strValue1.IndexOf(".",0) >= 0 &&
										strValue1.Trim().Length > strValue1.IndexOf(".",0) + 3)
										strValue1 = strValue1.Substring(0,strValue1.IndexOf(".",0) + 3);
								}
								if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["max"] == true)
								{
									if (str.Trim()=="Min")
									{
										str += " and Max";
										strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["max_group_summary"].ToString().Trim();
										if (strValue2.IndexOf(".",0) >= 0 &&
											strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
											strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);
									}
									else 
									{
										str="Max";
										strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["max_group_summary"].ToString().Trim();
										if (strValue2.IndexOf(".",0) >= 0 &&
											strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
											strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);
									}
								} 
								if (str.Trim().Length > 0)
								{
									ypos = (ypos) + (bodyHeight);
									e.Graphics.DrawString(str, this.m_fontSummary,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
									xpos = (e.MarginBounds.Left)+ (float)(sizeFWidth.Width) ;
									e.Graphics.DrawString(strValue1, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
									if (strValue2.Trim().Length > 0)
									{
										xpos = xpos + (float)(sizeFWidth.Width) ;
										e.Graphics.DrawString(strValue2, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
									}
									linecount++;
								}
								
								//sum and average
								str="";
								strValue1="";
								strValue2="";
								if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["sum"] == true)
								{
									str="Sum";
									strValue1 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["sum_group_summary_total"].ToString().Trim();
									if (strValue1.IndexOf(".",0) >= 0 &&
										strValue1.Trim().Length > strValue1.IndexOf(".",0) + 3)
										strValue1 = strValue1.Substring(0,strValue1.IndexOf(".",0) + 3);
								}    
								if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["avg_sum"] == true)
								{
									if (str.Trim()=="Sum")
									{
										str += " and Avg";
										strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["avg_sum_group_summary_total"].ToString().Trim();
										if (strValue2.IndexOf(".",0) >= 0 &&
											strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
											strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);
									}
									else 
									{
										str="Avg";
										strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["avg_sum_group_summary_total"].ToString().Trim();
										if (strValue2.IndexOf(".",0) >= 0 &&
											strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
											strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);
									}
								} 
								if (str.Trim().Length > 0)
								{
									ypos = (ypos) + (bodyHeight);
									e.Graphics.DrawString(str, this.m_fontSummary,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
									xpos = (e.MarginBounds.Left)+ (float)(sizeFWidth.Width) ;
									e.Graphics.DrawString(strValue1, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
									if (strValue2.Trim().Length > 0)
									{
										xpos = xpos + (float)(sizeFWidth.Width) ;
										e.Graphics.DrawString(strValue2, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
									}
									linecount++;
								}
							}
							//group count and percent
							str="";
							strValue1="";
							strValue2="";
							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["count"] == true)
							{
								str="Count";
								strValue1 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["count_group_summary_total"].ToString().Trim();
								if (strValue1.IndexOf(".",0) >= 0 &&
									strValue1.Trim().Length > strValue1.IndexOf(".",0) + 3)
									strValue1 = strValue1.Substring(0,strValue1.IndexOf(".",0) + 3);
							}    
							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["pct_count"] == true)
							{
								if (str.Trim()=="Count")
								{
									str += " and Pct";
									strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["pct_count_group_summary_total"].ToString().Trim();
									//limit numeric values to 2 places after the decimal
									if (strValue2.IndexOf(".",0) >= 0 &&
										strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
										strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);
								}
								else 
								{
									str="Pct";
									strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[y]["pct_count_group_summary_total"].ToString().Trim();
									if (strValue2.IndexOf(".",0) >= 0 &&
										strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
										strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);
								}
							} 
							if (str.Trim().Length > 0)
							{
								ypos = (ypos) + (bodyHeight);
								e.Graphics.DrawString(str, this.m_fontSummary,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
								xpos = (e.MarginBounds.Left)+ (float)(sizeFWidth.Width) ;
								e.Graphics.DrawString(strValue1, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
								if (strValue2.Trim().Length > 0)
								{
									xpos = xpos + (float)(sizeFWidth.Width) ;
									e.Graphics.DrawString(strValue2, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
								}
								linecount++;
							}
													
							if (linecount >= linesPerPage)
							{
								//start a new page so save the current 
								//dataview column number and 
								//increment the report config 
								//record number to print the
								//next column
								this.m_intCurrSummaryRow2 = y + 1;
								break;
							}
				   							  
				   

						}
						if (linecount >= linesPerPage)
						{
							
							break;
						}
						

					}
					if (linecount >= linesPerPage) break;
					
				}
				if ((y >= intLastGroupRow && intLastGroupRow > 0) || linecount2==linecount)
				{
					//group record count and percent compared to the report count
					str="";
					strValue1="";
					strValue2="";
					if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["count_group_records_YN"] == "Y")
					{
						str="Record Count";
						strValue1 = this.m_dsReportConfig.Tables["report_results"].Rows[intFirstGroupRow]["count_group_summary_record_total"].ToString().Trim();
						if (strValue1.IndexOf(".",0) >= 0 &&
							strValue1.Trim().Length > strValue1.IndexOf(".",0) + 3)
							strValue1 = strValue1.Substring(0,strValue1.IndexOf(".",0) + 3);
					}    
					if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["pct_count_group_records_YN"] == "Y")
					{
						if (str.Trim()=="Record Count")
						{
							str += " and Pct";
							strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[intFirstGroupRow]["pct_group_summary_record_total"].ToString().Trim();
							//limit numeric values to 2 places after the decimal
							if (strValue2.IndexOf(".",0) >= 0 &&
								strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
								strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);
						}
						else 
						{
							str="Record Pct";
							strValue2 = this.m_dsReportConfig.Tables["report_results"].Rows[intFirstGroupRow]["pct_group_summary_record_total"].ToString().Trim();
							if (strValue2.IndexOf(".",0) >= 0 &&
								strValue2.Trim().Length > strValue2.IndexOf(".",0) + 3)
								strValue2 = strValue2.Substring(0,strValue2.IndexOf(".",0) + 3);
						}
					} 
					if (str.Trim().Length > 0)
					{
						ypos = (ypos) + (bodyHeight * 2);
						e.Graphics.DrawString(str, this.m_fontSummary,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
						xpos = (e.MarginBounds.Left)+ (float)(sizeFWidth.Width) ;
						e.Graphics.DrawString(strValue1, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
						if (strValue2.Trim().Length > 0)
						{
							xpos = xpos + (float)(sizeFWidth.Width) ;
							e.Graphics.DrawString(strValue2, this.m_fontSummary,Brushes.Black, xpos,ypos, new StringFormat());                                         
						}
						linecount++;
						linecount++;
						if (linecount >= linesPerPage)
						{
							this.m_intCurrSummaryRow=x;
							if (this.m_intCurrRow2 != y + 1)
							   this.m_intCurrSummaryRow2 = y + 1;
						}
					}
				


				}

			}
			if (linecount >= linesPerPage)
			{
				if  (y >= intLastGroupRow && intLastGroupRow > 0) 
				{
					this.m_bSummaryDoneButNewPage = true;
					

				}
			
				//print the page in the middle of the page at the bottom
				sizeFWidth = e.Graphics.MeasureString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber);
						
				xpos = (e.PageBounds.Width/2)- (float)(sizeFWidth.Width * .50) ;
				ypos = e.PageBounds.Height - (pagenumHeight * 2);
				e.Graphics.DrawString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber, Brushes.Black, xpos, ypos, new StringFormat());
				e.HasMorePages=true;
			}
			
			/*****************************************************************************
			 **if all the report_fields rows have not been processed and the 
			 **and the current report_fields field_name is not equal to the field_name of
			 **the last row of the current group in the report_results table
			 **then we need to do some more group summary processing.
			 *****************************************************************************/
			if (x<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1  &&
			    	this.m_dsReportConfig.Tables["report_results"].Rows[intLastGroupRow]["field_name"].ToString().Trim().ToUpper()
			    	   !=
			    	this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim().ToUpper())
			{
				
				this.m_bPrintingGroupSummary=true;
				this.m_intCurrSummaryRow = x;

			}
			else 
			{  

				this.m_intCurrSummaryRow=0;
				this.m_intCurrSummaryRow2=0;
				this.m_bPrintingGroupSummary=false;
			
			}
			ypos = (ypos) + (bodyHeight * 2);
			linecount = linecount + 2;

			
		   
									     
									   


		}

		public void PrintDocumentWithDialog()
		{
			//PrintDialog dlg = new PrintDialog();
			this.m_dlgPrint.Document = this.m_printDoc;
			if (this.m_dlgPrint.ShowDialog() == DialogResult.OK)
			{
				
				this.m_printDoc.Print();
			}
		}
		public void PrintListBoxContents(string strHeader, System.Windows.Forms.ListBox p_oListBox)
		{
			this.m_strHeader = strHeader;
			this.m_oListBox = new ListBox();
			this.m_oListBox = p_oListBox;
			this.m_strPrintType = "LISTBOX";
			this.m_dlgPrint.Document = this.m_printDoc;
			if (this.m_dlgPrint.ShowDialog() == DialogResult.OK)
			{
				try
				{
					this.m_printDoc.Print();
				}
				catch (Exception e)
				{
					MessageBox.Show(e.Message,"Print Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
				}
			}
			
		}
		public void PrintReport(System.Windows.Forms.DataGrid p_dg, System.Data.DataSet p_dsReportConfig)
		{
			ado_data_access p_ado = new ado_data_access();
			System.Windows.Forms.CurrencyManager p_cm = (System.Windows.Forms.CurrencyManager)p_dg.BindingContext[p_dg.DataSource,p_dg.DataMember];
			System.Data.DataView p_dv = (DataView)p_cm.List;
			FIA_Biosum_Manager.frmTherm p_therm = new frmTherm();
			p_therm.progressBar1.Minimum = 1;
			p_therm.progressBar1.Maximum = p_dv.Count;
		    p_therm.progressBar1.Value=1;
            System.Data.DataTable p_dtNew = p_ado.ConvertDataViewToDataTable(p_dv,ref this.m_frmTherm);
			p_therm=null;
			p_ado = null;
		    this.PrintReport(p_dtNew,p_dsReportConfig);
		}
		public void PrintReport(System.Data.DataTable p_dt, System.Data.DataSet p_dsReportConfig)
		{
			string str="";
			int x=0;
			int y=0;
			int z=0;
			
			System.Data.DataView p_dv;
			System.Data.DataTable p_dtNew;
			this.m_dt = new DataTable();
			this.m_dt = p_dt.Copy();
			this.m_dsReportConfig = p_dsReportConfig;
			this.m_strPrintType = "REPORT";
			utils p_utils = new utils();		
			ado_data_access p_ado = new ado_data_access();
			
            
			
			
            
			this.m_fontReportTitle = new Font("Courier", 12, System.Drawing.FontStyle.Bold);
			this.m_fontBody = new Font("Courier", 8, System.Drawing.FontStyle.Regular);
			this.m_fontReportDate = new Font("Courier", 8, System.Drawing.FontStyle.Italic);
			this.m_fontPageNumber = new Font("Courier", 8, System.Drawing.FontStyle.Italic);
			this.m_fontColumns = new Font("Courier", 10, System.Drawing.FontStyle.Bold);
			this.m_fontSummary = new Font("Courier", 8, System.Drawing.FontStyle.Regular);
			this.m_fontSummaryColumn = new Font("Courier",8,System.Drawing.FontStyle.Underline);
			this.m_fontSummaryHeader = new Font("Courier",8,System.Drawing.FontStyle.Italic);
			this.m_fontEndOfReport = new Font("Courier",8,System.Drawing.FontStyle.Bold);

			this.m_frmTherm.Show();
			this.m_frmTherm.Focus();
			this.m_frmTherm.Text = "Processing Report";
			this.m_frmTherm.Refresh();
			this.m_frmTherm.progressBar1.Minimum = 1;
			this.m_frmTherm.AbortProcess = false;
			this.m_intCounter=0;


			//-----step 1------//

			this.m_frmTherm.progressBar1.Maximum = this.m_dt.Columns.Count;
			//remove columns that are not on the report
			// and get the datatype of the columns that are on the report
			for (y=0;y<=this.m_dt.Columns.Count-1;y++)
			{
				System.Windows.Forms.Application.DoEvents();
				if (this.m_frmTherm.AbortProcess == true) 
				{
					p_utils = null;
					p_ado = null;
					return;
				}
				this.m_intCounter++;
				this.m_frmTherm.Increment(this.m_intCounter);
				for (x=0; x<= this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;x++)
				{
					if (this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim().ToUpper() ==
						this.m_dt.Columns[y].ColumnName.ToString().Trim().ToUpper())
					{
						//save the datatype
						this.m_dsReportConfig.Tables["report_fields"].Rows[x]["datatype"] = this.m_dt.Columns[y].DataType.FullName.ToString().Trim();
						break;
					}
					
					
				}
				if (x <= this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1)
				{
					//save whether the field is a string or not
					this.m_dsReportConfig.Tables["report_fields"].Rows[x]["stringdatatype_YN"] =
						p_ado.getIsTheFieldAStringDataType(this.m_dt.Columns[y].DataType.FullName.ToString()); 

				}
					//remove the column because it is not on the report
				else 
				{
					this.m_dt.Columns.Remove(this.m_dt.Columns[y].ColumnName.ToString());
					y=y-1;
				}
			}
            

			//-----step 2----

			//find the string with the longest length for each column
			//make sure that the fields the user selects to do math can do math

			//check if all the fields are on one row or each field has its own row
			if (this.m_dsReportConfig.Tables["record_layout"].Rows[0]["type"].ToString() == "C")
			{
				//all the fields are on one row
				//find the largest string for each report column for measuring column widths
				//go through all the rows in the datatable. Limit numeric to 4 decimal positions
				this.m_frmTherm.progressBar1.Maximum += this.m_dt.Rows.Count;
				for (x=0; x<= this.m_dt.Rows.Count-1; x++)
				{
    		     	System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) 
					{
						p_utils = null;
						p_ado = null;
						return;
					}
					this.m_intCounter++;
					this.m_frmTherm.Increment(this.m_intCounter);

					//go through each column in the row
					//loop agenda:
					//	1. initialize the column position of field on the report to -1
					//  2. those columns that the user selected to perform numeric functions
					//       (ie sum, sum avg) then make sure the column values can perform math
					//  3. limit numeric values to a maximum of 2 numbers after the decimal
					//	4. get the columns largest string value for formatting column width
					for (y=0;y<=this.m_dt.Columns.Count-1;y++)
					{
						//see if the column is one to print
						for (z=0; z<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;z++)
						{
							if (this.m_dsReportConfig.Tables["report_fields"].Rows[z]["field_name"].ToString().Trim().ToUpper() ==
								this.m_dt.Columns[y].ColumnName.ToString().Trim().ToUpper())
							{
								//initialize the xpos of the printed column
								this.m_dsReportConfig.Tables["report_fields"].Rows[z]["xpos"] = -1;

								str = this.m_dt.Rows[x][y].ToString().Trim();
								//if the user wants to do math then make
								//sure that the string only contains numbers
								if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[z]["candomath"]==true)    
								{
									if (p_utils.IsTheStringNumeric(str)==false)
									{
										this.m_dsReportConfig.Tables["report_fields"].Rows[z]["candomath"] = false;
									}
								}
								if ((string)this.m_dsReportConfig.Tables["report_fields"].Rows[z]["stringdatatype_YN"] == "N" ||
									(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[z]["candomath"] == true)
								{
									//right justify for numeric items
									if (this.m_dt.Rows[x][y] == System.DBNull.Value)
									{
										str = "0.000";
									}
									else
									{
										str = String.Format("{0:f3}",Math.Round(Convert.ToDecimal(this.m_dt.Rows[x][y]),3));
									}
								}
								else
								{
									str = this.m_dt.Rows[x][y].ToString().Trim();
								}
								

								if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[z]["candomath"] == true)
								{
									//limit numeric values to 2 places after the decimal
									if (str.IndexOf(".",0) >= 0 &&
										str.Trim().Length > str.IndexOf(".",0) + 3)
											str = str.Substring(0,str.IndexOf(".",0) + 3);
								}
								//compare string lengths to get the longest string for that field
								if (this.m_dsReportConfig.Tables["report_fields"].Rows[z]["largeststring"].ToString().Trim().Length <
									str.Trim().Length)
								{
									this.m_dsReportConfig.Tables["report_fields"].Rows[z]["largeststring"] = str;
								}

								

							}
						}
					}
				}
			}
			else 
			{
				//each field has its own row
				//get the longest field name
				str="";
				this.m_frmTherm.progressBar1.Maximum += this.m_dsReportConfig.Tables["report_fields"].Rows.Count;
				for (x=0; x <= this.m_dsReportConfig.Tables["Report_Fields"].Rows.Count-1;x++)
				{
			
					System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) 
					{
						p_utils = null;
						p_ado = null;
						return;
					}
					this.m_intCounter++;
					this.m_frmTherm.Increment(this.m_intCounter);

					if (str.Trim().Length < this.m_dsReportConfig.Tables["Report_Fields"].Rows[x]["field_name"].ToString().Trim().Length)
					{
						str = this.m_dsReportConfig.Tables["Report_Fields"].Rows[x]["field_name"].ToString().Trim();
						if (p_utils.IsTheStringNumeric(str) == true &&
							str.IndexOf(".",0) >= 0 &&
							str.Trim().Length > str.IndexOf(".",0) + 20)
							str = str.Substring(0,str.IndexOf(".",0) + 20);

					}

				}
				this.m_frmTherm.progressBar1.Maximum += this.m_dsReportConfig.Tables["report_fields"].Rows.Count;
				//save the value to the largeststring field 
				for (x=0; x <= this.m_dsReportConfig.Tables["Report_Fields"].Rows.Count-1;x++)
				{
			
					System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) 
					{
						p_utils = null;
						p_ado = null;
						return;
					}
					this.m_intCounter++;
					this.m_frmTherm.Increment(this.m_intCounter);
					this.m_dsReportConfig.Tables["Report_Fields"].Rows[x]["largeststring"] = str;
				}
			}

			p_utils = null;
			
			p_dv = this.m_dt.DefaultView;
			
			
			//create the group and sort string
			str="";
			this.m_frmTherm.progressBar1.Maximum += this.m_dsReportConfig.Tables["report_fields"].Rows.Count;
			for (x=0;x<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;x++)
			{
				System.Windows.Forms.Application.DoEvents();
				if (this.m_frmTherm.AbortProcess == true) 
				{
					p_dv.Dispose();
					p_dv = null;
					p_ado = null;
					return;
				}
				this.m_intCounter++;
				this.m_frmTherm.Increment(this.m_intCounter);

				if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["group"]==true ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["sort"]==true)
				{
					if (str.Trim().Length == 0)
					{
						str = "[";
					}
					else 
					{
						str = str + ",[";
					}
					str += this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim() + "] " + 
						this.m_dsReportConfig.Tables["report_fields"].Rows[x]["sortorder"].ToString().Trim();
				}
			}

			//if there is a group or sort field than 
			//recreate the table with the new view
			if (str.Trim().Length > 0)
			{

				p_dv.Sort = str;
				p_dtNew = p_ado.ConvertDataViewToDataTable(p_dv,ref this.m_frmTherm);
				if (this.m_frmTherm.AbortProcess == true) 
				{
					p_dv.Dispose();
					p_dv = null;
					p_ado = null;
					return;
				}

				//create a new view from the new table just created
				this.m_dv = p_dtNew.DefaultView;
			}
			else
			{
				
				this.m_dv = p_dv;
			}
			getNumberOfSummaryLines();
			if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["group_YN"] == "Y")
			{
				getGroupReportTotals();
			}
			else
			{
				getReportTotals();
			}

			if (this.m_frmTherm.AbortProcess == true) 
			{
				p_dv.Dispose();
				p_dv = null;
				p_ado = null;
				return;
			}

			this.m_frmTherm.progressBar1.Value = this.m_frmTherm.progressBar1.Maximum;
			this.m_frmTherm.Hide();
			
			
				if ((bool)this.m_dsReportConfig.Tables["Report_Fields"].Rows[0]["group"]==true)
				{ 
					this.m_strCurrGroup="";
					this.m_strNewGroup="";
					//check and set orientation to print portrait or landscape
					if ((string)this.m_dsReportConfig.Tables["orientation"].Rows[0]["type"] == "P")
						this.m_printReportGroupDoc.DefaultPageSettings.Landscape=false;
					else  this.m_printReportGroupDoc.DefaultPageSettings.Landscape=true;

					this.m_printReportGroupDoc.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.AllPages;
					if ((string)this.m_dsReportConfig.Tables["printpreview"].Rows[0]["printpreview_yn"] == "Y") 
					{
						this.m_dlgPrintPreview.Document = this.m_printReportGroupDoc;
					}
					else
					{
						this.m_dlgPrint.Document = this.m_printReportGroupDoc;
					}

				}
				else 
				{
					//check and set orientation to print portrait or landscape
					if ((string)this.m_dsReportConfig.Tables["orientation"].Rows[0]["type"] == "P")
						this.m_printReportDoc.DefaultPageSettings.Landscape=false;
					else  this.m_printReportDoc.DefaultPageSettings.Landscape=true;
					this.m_printReportDoc.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.AllPages;

					/***********************************************************
					 **if the user does not want to print the detail lines 
					 **than make the current row greater than the number of 
					 **records in the view
					 ***********************************************************/
					if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["print_detail_line_YN"] == "N")
					{
						this.m_intCurrRow = this.m_dv.Table.Rows.Count + 1;
					}
					if ((string)this.m_dsReportConfig.Tables["printpreview"].Rows[0]["printpreview_yn"] == "Y") 
					{
						this.m_dlgPrintPreview.Document = this.m_printReportDoc;
					}
					else
					{
						this.m_dlgPrint.Document = this.m_printReportDoc;
					}
				}
			if ((string)this.m_dsReportConfig.Tables["printpreview"].Rows[0]["printpreview_yn"] == "Y") 
			{
				this.m_dlgPrintPreview.ShowDialog();
			}
			else
			{

				if (this.m_dlgPrint.ShowDialog() == DialogResult.OK)
				{
				
					if ((bool)this.m_dsReportConfig.Tables["Report_Fields"].Rows[0]["group"]==true)
					{
						this.m_printReportGroupDoc.Print();
 
					}
					else
					{
						this.m_printReportDoc.Print();
					}
				}
			}

			
			
			


		}
		private void getNumberOfSummaryLines()
		{
			int x=0;
			int intCurCount=0;
			int intCurCount2=0;
			this.m_intReportSummaryLineTotal=0;
			this.m_intGroupSummaryLineTotal=0;
		  
			if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["print_report_totals_YN"] == "Y" ||
				(string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["print_group_summary_YN"] == "Y")
			{
		  	
				for (x=0;x<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;x++)
				{
					intCurCount=intCurCount2;
					if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["candomath"] == true)
					{
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["sum"] == true ||
							(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["avg_sum"] == true) 
						{
							intCurCount2++;
						}
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["min"] == true ||
							(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["max"] == true) 
						{
							intCurCount2++;
						}
		    	 
					}
					if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["count"] == true ||
						(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["pct_count"] == true) 
					{
						intCurCount2++;
					}
   	     
					//if there is a summary for the field then add a line to print its header
					if (intCurCount != intCurCount2)
					{
						intCurCount2++;
					}
				}
		    
				if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["print_report_totals_YN"] == "Y")
					this.m_intReportSummaryLineTotal = intCurCount2;
				if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["print_group_summary_YN"] == "Y")
				{
					this.m_intGroupSummaryLineTotal = intCurCount2;
					if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["count_group_records_YN"] == "Y" ||
						(string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["pct_count_group_records_YN"] == "Y")
					{
						this.m_intGroupSummaryLineTotal = this.m_intReportSummaryLineTotal + 1;
					}
				}	  
			}
    	
		}
		private void getGroupReportTotals()
		{
			string str="";
			
			int intRptCfgRow=0;
			int intDvRow=0;
			int intDvCol=0;
			int intRptResultLastRow=0;
			int intRptResultFirstRow=0;
			
			bool bFirstTime=true;

            this.m_frmTherm.progressBar1.Maximum += this.m_dv.Table.Rows.Count;
			//loop through each row in the dataview
			for (intDvRow=0; intDvRow<=this.m_dv.Table.Rows.Count; intDvRow++)
			{
				System.Windows.Forms.Application.DoEvents();
				if (this.m_frmTherm.AbortProcess == true) 
				{
					return;
				}
				this.m_intCounter++;
				this.m_frmTherm.Increment(this.m_intCounter);

				//lets see if the group changed
				if (intDvRow < this.m_dv.Table.Rows.Count)	this.m_bNewGroup = NewGroup(intDvRow);

				if (intDvRow == this.m_dv.Table.Rows.Count)
				{
					FinalizeGroupCounts();   //take care of averages and update report total counts

					//if there are no more rows to process than break free from the loop
					if (intDvRow == this.m_dv.Table.Rows.Count) break;
				}
                
				
				//loop thru the report config data set and only
				//process the columns from the dataview that are in
				//the report config dataset
				for (intRptCfgRow=0;intRptCfgRow <= this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;intRptCfgRow++)
				{
					//loop through the dataview columns
					for (intDvCol=0;intDvCol<=this.m_dv.Table.Columns.Count-1;intDvCol++)
					{
						//if the dataview column name equals the report config column name
						//then update group counts
						if (this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["field_name"].ToString().Trim().ToUpper() ==
							this.m_dv.Table.Columns[intDvCol].ColumnName.ToString().Trim().ToUpper())
						{
							
							if (this.m_bNewGroup==true)
							{
								//add all the other columns to the group record
								AddGroupRecord();
								//get the first and last of the group records
								intRptResultLastRow=
									this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;
								for (intRptResultFirstRow=this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;
									intRptResultFirstRow>=0;
									intRptResultFirstRow--)
								{
									str = this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultFirstRow]["group_field1_value"].ToString().Trim() +
										this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultFirstRow]["group_field2_value"].ToString().Trim() + 
										this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultFirstRow]["group_field3_value"].ToString().Trim();
                                      
									if (str.Trim().ToUpper() != this.m_strNewGroup.Trim().ToUpper())
									{
										break;
									}
									   
								}
								intRptResultFirstRow=intRptResultFirstRow+1;
								bFirstTime=true;
							}
							this.m_bNewGroup=false;
							this.m_strCurrGroup=this.m_strNewGroup;
							//}
							if (this.m_bNewGroup==false)
							{
								//this is where we update the report counts, etc...
								if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["sum"]==true ||
									(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["max"]==true ||
									(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["min"]==true ||
									(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["avg_sum"]==true ||
									(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["count"]==true ||
									(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["pct_count"] == true)
								{
									IncrementGroupCounts(intRptCfgRow, intDvRow,intDvCol, ref bFirstTime,intRptResultFirstRow, intRptResultLastRow);
								}
							}
						}
					}
				}
				bFirstTime=false;
				
				this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultFirstRow]["count_group_summary_record_total"] = 
					Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultFirstRow]["count_group_summary_record_total"]) + 1;
				

				if (this.m_bNewGroup==true)
				{
					this.m_bNewGroup=false;
					this.m_strCurrGroup=this.m_strNewGroup;
				}
			}
			
		
		}

	

		

		private bool NewGroup(int p_intCurrRow)
		{
			int intRptCfgRow=0;
			int intDvRow=0;
			string strGroup1Value="";
			string strGroup2Value="";
			string strGroup3Value="";
			string strValue="";
			
			//lets see if the group has changed
			//
			this.m_strNewGroup="";
			for (intRptCfgRow=0;intRptCfgRow<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;intRptCfgRow++)
			{
				//check only columns that are part of a group
				if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["group"]==true)
				{
					for (intDvRow=0;intDvRow<=this.m_dv.Table.Columns.Count-1;intDvRow++)
					{

						//see if the column is part of the group 
						if (this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["field_name"].ToString().Trim().ToUpper() ==
							this.m_dv.Table.Columns[intDvRow].ColumnName.ToString().Trim().ToUpper())
						{
							//save the value of the group column;
							
							
							strValue = this.m_dv.Table.Rows[p_intCurrRow][intDvRow].ToString().Trim();
							if (strValue.Trim().Length ==0) strValue="null";
							if (strGroup1Value.Trim().Length ==0)
								strGroup1Value = strValue;
							else if (strGroup2Value.Trim().Length == 0)
								strGroup2Value = strValue;
							else if (strGroup3Value.Trim().Length == 0)
								strGroup3Value = strValue;
							this.m_strNewGroup += strValue;
							
						}
					}
				}
				else
				{
					break;
				}
						    
			}
			//check to see if the values in the group columns changed
			if (this.m_strNewGroup != this.m_strCurrGroup)
			{
				this.m_strGroup1Value=strGroup1Value;
				this.m_strGroup2Value=strGroup2Value;
				this.m_strGroup3Value=strGroup3Value;
				return true;
			}
			return false;


		}

		private void AddGroupRecord()
		{
			int a;
			//int b;
			int c;
			int intGroupCount=1;
			string strGroup1="";
			string strGroup2="";
			string strGroup3="";
			System.Data.DataRow p_row;

			//create a new group record
			int intCurrGroup = this.m_dsReportConfig.Tables["report_results"].Rows.Count;
			p_row = this.m_dsReportConfig.Tables["report_results"].NewRow();
			p_row["results_type"] = "G";  // column results
			p_row["group_field1_value"] = "";
			p_row["group_field2_value"] = "";
			p_row["group_field3_value"]="";
			p_row["field_name"] = "";
			p_row["sum_group_summary_total"] = 0;
			p_row["sum_report_summary_total"] = 0;
			p_row["avg_sum_group_summary_total"] = 0;
			p_row["avg_sum_report_summary_total"] = 0;
			p_row["max_group_summary"] = 0;
			p_row["max_report_summary"] = 0;
			p_row["min_group_summary"] = 0;
			p_row["min_report_summary"] = 0;
			p_row["count_group_summary_total"] = 0;
			p_row["count_report_summary_total"] = 0;
			p_row["pct_count_group_summary_total"] = 0;
			p_row["pct_count_report_summary_total"] = 0;
			p_row["count_group_summary_record_total"] = 0;
			p_row["count_report_summary_record_total"] = 0;
			p_row["pct_group_summary_record_total"] = 0;
			p_row["pct_report_summary_record_total"] = 0;
			p_row["min_init"]=false;
			
			this.m_dsReportConfig.Tables["report_results"].Rows.Add(p_row);

			//populate the record with the group values
			for (a=0;a<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;a++)
			{
				if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[a]["group"]==true)
				{
					switch (intGroupCount)
					{
						case 1:
							strGroup1= this.m_strGroup1Value; //this.m_dsReportConfig.Tables["report_fields"].Rows[a]["value"].ToString().Trim();
							this.m_dsReportConfig.Tables["report_results"].Rows[intCurrGroup]["group_field1_value"] = 
								strGroup1;
								
							intGroupCount=2;
							break;
						case 2:
							strGroup2= this.m_strGroup2Value; //this.m_dsReportConfig.Tables["report_fields"].Rows[a]["value"].ToString().Trim();
							this.m_dsReportConfig.Tables["report_results"].Rows[intCurrGroup]["group_field2_value"] = 
								strGroup2;
								
							intGroupCount=3;
							break;
						case 3:
							strGroup3= this.m_strGroup3Value; //this.m_dsReportConfig.Tables["report_fields"].Rows[a]["value"].ToString().Trim();
							this.m_dsReportConfig.Tables["report_results"].Rows[intCurrGroup]["group_field3_value"] = 
								strGroup3;
							break;
					}
				}

			}
			
			
			for (c=0; c<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;c++)
			{
				
				if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[c]["sum"] == true ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[c]["max"] == true ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[c]["min"] == true ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[c]["count"] == true  ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[c]["avg_sum"] == true ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[c]["pct_count"] == true)
				{
					p_row = this.m_dsReportConfig.Tables["report_results"].NewRow();
					p_row["results_type"] = "G";  // column results
					p_row["group_field1_value"] = strGroup1;
					p_row["group_field2_value"] = strGroup2;
					p_row["group_field3_value"]=strGroup3;
					p_row["field_name"] =
						this.m_dsReportConfig.Tables["report_fields"].Rows[c]["field_name"].ToString().Trim();
					p_row["sum_group_summary_total"] = 0;
					p_row["sum_report_summary_total"] = 0;
					p_row["avg_sum_group_summary_total"] = 0;
					p_row["avg_sum_report_summary_total"] = 0;
					p_row["max_group_summary"] = 0;
					p_row["max_report_summary"] = 0;
					p_row["min_group_summary"] = 0;
					p_row["min_report_summary"] = 0;
					p_row["count_group_summary_total"] = 0;
					p_row["count_report_summary_total"] = 0;
					p_row["pct_count_group_summary_total"] = 0;
					p_row["pct_count_report_summary_total"] = 0;
					p_row["count_group_summary_record_total"] = 0;
					p_row["count_report_summary_record_total"] = 0;
					p_row["pct_group_summary_record_total"] = 0;
					p_row["pct_report_summary_record_total"] = 0;
					p_row["min_init"]=false;
					this.m_dsReportConfig.Tables["report_results"].Rows.Add(p_row);
				}
				
			}
			

			

		}
		private void IncrementGroupCounts(int intRptCfgRow, int intDvRow, int intDvCol,ref  bool bFirstTime,int intRptResultFirstRow, int intRptResultLastRow)
		{
			double dbl=0;
			double dbl2=0;
			int x;


			//go through all the columns in the record 
			// and update sum, counts, max, and min values
			for (x=intRptResultFirstRow;x<=intRptResultLastRow;x++)
			{
				if (this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["field_name"].ToString().Trim().ToUpper()==
					this.m_dsReportConfig.Tables["report_results"].Rows[x]["field_name"].ToString().Trim().ToUpper())
				{
					if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["candomath"]==true)
					{
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["sum"]==true &&
							this.m_dv.Table.Rows[intDvRow][intDvCol].ToString().Trim().Length > 0)
						{
							dbl = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["sum_group_summary_total"]);
							dbl2 = Convert.ToDouble(this.m_dv.Table.Rows[intDvRow][intDvCol]);
							dbl = dbl + dbl2;
							this.m_dsReportConfig.Tables["report_results"].Rows[x]["sum_group_summary_total"] = dbl;
						}
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["max"]==true &&
							this.m_dv.Table.Rows[intDvRow][intDvCol].ToString().Trim().Length > 0)
						{
							dbl = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["max_group_summary"]);
							dbl2 = Convert.ToDouble(this.m_dv.Table.Rows[intDvRow][intDvCol]);
							if (dbl2 > dbl) 
								this.m_dsReportConfig.Tables["report_results"].Rows[x]["max_group_summary"] = dbl2;

						}
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["min"]==true &&
							this.m_dv.Table.Rows[intDvRow][intDvCol].ToString().Trim().Length > 0)
						{
							dbl = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_group_summary"]);
							dbl2 = Convert.ToDouble(this.m_dv.Table.Rows[intDvRow][intDvCol]);
							if (dbl2 < dbl || (bool)this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_init"] == false)
							{
								this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_group_summary"] = dbl2;
								if ((bool)this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_init"] == false)
									this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_init"] = true;
							}
						}
					}
					//group count
					//if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["count"]==true && 
					if (this.m_dv.Table.Rows[intDvRow][intDvCol].ToString().Trim().Length > 0)
					{
						dbl = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_group_summary_total"]);
						dbl += 1;
						this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_group_summary_total"] = dbl;
					}
				}
			}

		}
		//take care of averages and update report total counts, sums, etc 
		private void FinalizeGroupCounts()
		{
			int x;
			double dblTtlRecords=0;
			double dblTtlColumnSum=0;
			
			double dblTtlColumnCount=0;
			double dblColumnMax=0;
			double dblColumnMin=0;
			string strGroup="";
			
			string strField="";
			string strCurField="";
			int intRptCfgRow=0;
			int intFirstFieldRecord=0;
			int intGroupRecord=0;
			
			bool bFirstTime=true;
			
			int intCount=0;
			
			int intRptResultsRow=0;
			//--------------SECTION 1--------------------
			/**********************************************************************
			 **accumulate all the group summary record counts and save the value
			 ** to the report summary record count_report_summary_record_total
			 **********************************************************************/
			// to the count_report_record_total field
			for (x=0;x<=this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;x++)
			{
				strGroup = this.m_dsReportConfig.Tables["report_results"].Rows[x]["group_field1_value"].ToString().Trim() + 
					this.m_dsReportConfig.Tables["report_results"].Rows[x]["group_field2_value"].ToString().Trim() + 
					this.m_dsReportConfig.Tables["report_results"].Rows[x]["group_field3_value"].ToString().Trim();
				strField = this.m_dsReportConfig.Tables["report_results"].Rows[x]["field_name"].ToString().Trim();
				if (strGroup.Length > 0 && strField.Length == 0)
				{
					//get total number of records in report
					dblTtlRecords += Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_group_summary_record_total"]);
				}
				else
				{
					if (strField.Length > 0)
					{

					}
				}

			}
			this.m_dsReportConfig.Tables["report_results"].Rows[0]["count_report_summary_record_total"] = dblTtlRecords;
			
			//------------SECTION 2-----------------
			/*************************************************************
			 **average the group record counts.  1)Divide the 
			 **  number of records in each group by the total number
			 **  of records for the report and save the value to
			 **  pct_group_summary_record_total.
			 **2)Divide the number of values for each field in a single
			 **  group and divide by the total number of records in
			 **  the group and save the value to pct_count_group_summary_total
			 *************************************************************/
			for (x=0;x<=this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;x++)
			{
				strGroup = this.m_dsReportConfig.Tables["report_results"].Rows[x]["group_field1_value"].ToString().Trim() + 
					this.m_dsReportConfig.Tables["report_results"].Rows[x]["group_field2_value"].ToString().Trim() + 
					this.m_dsReportConfig.Tables["report_results"].Rows[x]["group_field3_value"].ToString().Trim();
				strField = this.m_dsReportConfig.Tables["report_results"].Rows[x]["field_name"].ToString().Trim();
				if (strGroup.Length > 0 && strField.Length == 0)
				{
					intCount=x;
					this.m_dsReportConfig.Tables["report_results"].Rows[x]["pct_group_summary_record_total"] = 
						(Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_group_summary_record_total"]) / dblTtlRecords) * 100;
				}
				else if (strGroup.Length > 0 && strField.Length > 0 && (int)this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_group_summary_total"] > 0)
				{
					this.m_dsReportConfig.Tables["report_results"].Rows[x]["pct_count_group_summary_total"] = 
						(Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_group_summary_total"]) 
						/ 
						Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[intCount]["count_group_summary_record_total"])) * 100;
				}
			}


			//--------------SECTION 3-----------------
			intGroupRecord=-1;
			/*******************************************************************
			 **loop thru the report config data set and figure
			 ** accumalated sums, counts, max, and min from group summary
			 ** and save the values to report summary. 
			 *******************************************************************/
			for (intRptCfgRow=0;intRptCfgRow <= this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;intRptCfgRow++)
			{
				//loop through the results table
				for (intRptResultsRow=0;intRptResultsRow<=this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;intRptResultsRow++)
				{
					/*****************************************
					 **initialize variables for a new field
					 *****************************************/
					dblTtlColumnSum = 0;
					dblColumnMax = 0;
					dblColumnMin = 0;
					dblTtlColumnCount = 0;
					intCount=0;
					strGroup = this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultsRow]["group_field1_value"].ToString().Trim() + 
						this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultsRow]["group_field2_value"].ToString().Trim() + 
						this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultsRow]["group_field3_value"].ToString().Trim();
					strField = this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultsRow]["field_name"].ToString().Trim();
					if (strGroup.Length > 0 && strField.Length == 0 && intRptResultsRow > intGroupRecord)
					{
						intGroupRecord = intRptResultsRow;
					}
					else
					{
						/***************************************************************************
						 **if the report results field name equals the report config field name
						 **then accumulate and analyze group counts,sums,max, and min
						 ***************************************************************************/
						if (this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["field_name"].ToString().Trim().ToUpper() ==
							this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultsRow]["field_name"].ToString().Trim().ToUpper())
						
						{
							intFirstFieldRecord = intRptResultsRow;
							strCurField = this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["field_name"].ToString().Trim();
							bFirstTime=true;
							/*************************************************************
							 **loop through all the reports results rows and 
							 **accumulate sum, count, min, and max from 
							 **all group summary fields with the current field name
							 *************************************************************/
							for (x=intFirstFieldRecord; x<= this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;x++)
							{
								strField = this.m_dsReportConfig.Tables["report_results"].Rows[x]["field_name"].ToString().Trim();
								if (strField == strCurField)
								{
									/*********************************************
									 **check if the column can do math
									 *********************************************/
									if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["candomath"] == true)
									{
										if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["sum"]==true || 
											(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["avg_sum"]==true) 
										{
											/*****************************************
											 **accumulate group summary sums
											 *****************************************/
											dblTtlColumnSum += Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["sum_group_summary_total"]);
											intCount++;
										}
										if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["max"]==true )
										{
											/******************************************
											 **analyze if this is the reports maximum value
											 ******************************************/
											if (dblColumnMax < Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["max_group_summary"]))
												dblColumnMax = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["max_group_summary"]);
										}
										if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["min"]==true )
										{
											/******************************************
											 **analyze if this is the reports minimum value
											 ******************************************/
											if (bFirstTime==true &&
												(bool)this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_init"] == true)
											{
												dblColumnMin = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_group_summary"]);
												bFirstTime=false;
											}
											else
											{
												if (dblColumnMin > Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_group_summary"])&&
													(bool)this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_init"] == true)
													dblColumnMin = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_group_summary"]);
											}
										}
									}
									/*************************************
									 **accumulate the fields record count
									 *************************************/
									dblTtlColumnCount += Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_group_summary_total"]);

									/*************************************
									 **average the sum group counts
									 *************************************/
									if ((int)this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_group_summary_total"] > 0 &&
										(double)this.m_dsReportConfig.Tables["report_results"].Rows[x]["sum_group_summary_total"] > 0)
									{
										this.m_dsReportConfig.Tables["report_results"].Rows[x]["avg_sum_group_summary_total"] = 
											(Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["sum_group_summary_total"])
											/ Convert.ToInt32(this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_group_summary_total"]));
									}
								}
								
							}
							/********************************************************************
							 **okay, all the group summary rows for the current field name have
							 **  been summed and counted so lets now update the 
							 **  report summary fields.
							 ********************************************************************/
							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["sum"]==true || 
								(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["avg_sum"]==true) 
							{
								this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["sum_report_summary_total"] = dblTtlColumnSum;

								//	dblTtlColumnSum / intCount;
								this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["avg_sum_report_summary_total"] = 
									Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["sum_report_summary_total"]) 
									/
									dblTtlColumnCount; //Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["count_report_summary_total"]);

							}
							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["max"]==true )
							{
								this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["max_report_summary"] = dblColumnMax;
							}
							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["min"]==true )
							{
								this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["min_report_summary"] = dblColumnMin;
							}
							this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["count_report_summary_total"] = dblTtlColumnCount;
							this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["pct_count_report_summary_total"] = 
								(dblTtlColumnCount  / Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[0]["count_report_summary_record_total"])) * 100;

								    
							break;
						}
					}
				}
			}

			this.m_dsReportConfig.WriteXml("c:\\temp\\fia_biosum_report.xml");

		}
		private void getReportTotals()
		{
			
			int x=0;
			
			int intRptCfgRow=0;
			int intDvRow=0;
			int intDvCol=0;
			
			//add the top record for report totals
			this.AddRecord("");
 
			//add all the fields that perform math
			for (x=0; x<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;x++)
			{
				if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["sum"] == true ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["max"] == true ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["min"] == true ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["count"] == true  ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["avg_sum"] == true ||
					(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["pct_count"] == true)
				{
					this.AddRecord(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim());
				}
			}

            this.m_frmTherm.progressBar1.Maximum += this.m_dv.Table.Rows.Count;
			//loop through each row in the dataview
			for (intDvRow=0; intDvRow<=this.m_dv.Table.Rows.Count; intDvRow++)
			{
				System.Windows.Forms.Application.DoEvents();
				this.m_intCounter++;
				this.m_frmTherm.Increment(this.m_intCounter);
				if (intDvRow == this.m_dv.Table.Rows.Count)
				{
					FinalizeCounts();   //take care of averages and update report total counts
					//if there are no more rows to process than break free from the loop
					if (intDvRow == this.m_dv.Table.Rows.Count) break;
				}
                
				
				//loop thru the report config data set and only
				//process the columns from the dataview that are in
				//the report config dataset
				for (intRptCfgRow=0;intRptCfgRow <= this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;intRptCfgRow++)
				{
					//loop through the dataview columns
					for (intDvCol=0;intDvCol<=this.m_dv.Table.Columns.Count-1;intDvCol++)
					{
						//see if the dataview column name equals the report config column name
						if (this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["field_name"].ToString().Trim().ToUpper() ==
							this.m_dv.Table.Columns[intDvCol].ColumnName.ToString().Trim().ToUpper())
						{
							if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["sum"]==true ||
								(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["max"]==true ||
								(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["min"]==true ||
								(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["avg_sum"]==true ||
								(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["count"]==true ||
								(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["pct_count"] == true)
							{
								IncrementCounts(intRptCfgRow, intDvRow,intDvCol);
							}
							
						}
					}
				}
				this.m_dsReportConfig.Tables["report_results"].Rows[0]["count_report_summary_record_total"] = 
					Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[0]["count_report_summary_record_total"]) + 1;
			}
		
		}

		private void AddRecord(string strField)
		{
			
						
			System.Data.DataRow p_row;

			//create a new  record
			p_row = this.m_dsReportConfig.Tables["report_results"].NewRow();
			p_row["results_type"] = "R";  // column results
			p_row["group_field1_value"] = "";
			p_row["group_field2_value"] = "";
			p_row["group_field3_value"]="";
			p_row["field_name"] = strField;
			p_row["sum_group_summary_total"] = 0;
			p_row["sum_report_summary_total"] = 0;
			p_row["avg_sum_group_summary_total"] = 0;
			p_row["avg_sum_report_summary_total"] = 0;
			p_row["max_group_summary"] = 0;
			p_row["max_report_summary"] = 0;
			p_row["min_group_summary"] = 0;
			p_row["min_report_summary"] = 0;
			p_row["count_group_summary_total"] = 0;
			p_row["count_report_summary_total"] = 0;
			p_row["pct_count_group_summary_total"] = 0;
			p_row["pct_count_report_summary_total"] = 0;
			p_row["count_group_summary_record_total"] = 0;
			p_row["count_report_summary_record_total"] = 0;
			p_row["pct_group_summary_record_total"] = 0;
			p_row["pct_report_summary_record_total"] = 0;
			p_row["min_init"]=false;
			this.m_dsReportConfig.Tables["report_results"].Rows.Add(p_row);
		}

		private void IncrementCounts(int intRptCfgRow, int intDvRow, int intDvCol)
		{
			double dbl=0;
			double dbl2=0;
			int x;


			//go through all the columns in the record 
			// and update sum, counts, max, and min values
			for (x=1;x<=this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;x++)
			{
				if (this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["field_name"].ToString().Trim().ToUpper()==
					this.m_dsReportConfig.Tables["report_results"].Rows[x]["field_name"].ToString().Trim().ToUpper())
				{
					if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["candomath"]==true)
					{
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["sum"]==true &&
							this.m_dv.Table.Rows[intDvRow][intDvCol].ToString().Trim().Length > 0)
						{
							dbl = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["sum_report_summary_total"]);
							dbl2 = Convert.ToDouble(this.m_dv.Table.Rows[intDvRow][intDvCol]);
							dbl = dbl + dbl2;
							this.m_dsReportConfig.Tables["report_results"].Rows[x]["sum_report_summary_total"] = dbl;
						}
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["max"]==true &&
							this.m_dv.Table.Rows[intDvRow][intDvCol].ToString().Trim().Length > 0)
						{
							dbl = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["max_report_summary"]);
							dbl2 = Convert.ToDouble(this.m_dv.Table.Rows[intDvRow][intDvCol]);
							if (dbl2 > dbl) 
								this.m_dsReportConfig.Tables["report_results"].Rows[x]["max_report_summary"] = dbl2;

						}
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["min"]==true &&
							this.m_dv.Table.Rows[intDvRow][intDvCol].ToString().Trim().Length > 0)
						{
							dbl = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_report_summary"]);
							dbl2 = Convert.ToDouble(this.m_dv.Table.Rows[intDvRow][intDvCol]);
							if (dbl2 < dbl || (bool)this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_init"] == false)
							{
								this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_report_summary"] = dbl2;
								if ((bool)this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_init"] == false)
									this.m_dsReportConfig.Tables["report_results"].Rows[x]["min_init"] = true;
							}
						}
					}
					if (this.m_dv.Table.Rows[intDvRow][intDvCol].ToString().Trim().Length > 0)
					{
						dbl = Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_report_summary_total"]);
						dbl += 1;
						this.m_dsReportConfig.Tables["report_results"].Rows[x]["count_report_summary_total"] = dbl;
					}
				}
			}

		}
		private void FinalizeCounts()
		{
			string strField="";
			string strCurField="";
			int intRptCfgRow=0;
			int intFirstFieldRecord=0;
			int intRptResultsRow=0;
		

			/*******************************************************************
			 **loop thru the report config data set and figure
			 ** accumalated sums, counts, max, and min from group summary
			 ** and save the values to report summary. 
			 *******************************************************************/
			for (intRptCfgRow=0;intRptCfgRow <= this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;intRptCfgRow++)
			{
				//loop through the results table
				for (intRptResultsRow=1;intRptResultsRow<=this.m_dsReportConfig.Tables["report_results"].Rows.Count-1;intRptResultsRow++)
				{
					/*****************************************
					 **initialize variables for a new field
					 *****************************************/
					strField = this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultsRow]["field_name"].ToString().Trim();
					/***************************************************************************
					 **if the report results field name equals the report config field name
					 **then accumulate and analyze group counts,sums,max, and min
					 ***************************************************************************/
					if (this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["field_name"].ToString().Trim().ToUpper() ==
						this.m_dsReportConfig.Tables["report_results"].Rows[intRptResultsRow]["field_name"].ToString().Trim().ToUpper())
					
					{
						intFirstFieldRecord = intRptResultsRow;
						strCurField = this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["field_name"].ToString().Trim();
						
						/********************************************************************
						 **okay, all the group summary rows for the current field name have
						 **  been summed and counted so lets now update the 
						 **  report summary fields.
						 ********************************************************************/
						if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["sum"]==true || 
							(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[intRptCfgRow]["avg_sum"]==true) 
						{
							
							this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["avg_sum_report_summary_total"] = 
								Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["sum_report_summary_total"]) 
								/
								Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["count_report_summary_total"]);
						}
						
						this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["pct_count_report_summary_total"] = 
							Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[intFirstFieldRecord]["count_report_summary_total"])
							/
							Convert.ToDouble(this.m_dsReportConfig.Tables["report_results"].Rows[0]["count_report_summary_record_total"]) * 100;

						//processed the column so lets break and get another column		    
						break;
					}
				}
			}

			this.m_dsReportConfig.WriteXml("c:\\temp\\fia_biosum_report.xml");

		}

		private void printReportGroupDoc_PrintPage(Object sender, PrintPageEventArgs e)
		{
			float titleHeight;     //title height
			float dateHeight;      //date height
			float pagenumHeight;   //page number height
			float linesPerPage;    
			float columnHeight;    //column name height
			float bodyHeight;      //body of the report
			float ypos;
			float xpos;
			float linecount=0;
			int x=0;
			int y=0;
			int z=0;
			string str="";
			int intWidth=0;
			int intWidth2 = 0;
			float floatWidth=0;
			float floatWidth2=0;
			bool bRecordPrinted=false;
			System.Drawing.SizeF sizeFWidth;
			System.Drawing.SizeF sizeFWidth2;

			try
			{
				switch (this.m_strPrintType)
				{
					case "REPORT":
						Pen pen = new Pen(System.Drawing.Color.Black);

						// Used to store the number of record detail lines printed so far on the
						// current page.
						linecount=0;
					

						// Determine the height of the elements in the report
						titleHeight = this.m_fontReportTitle.GetHeight(e.Graphics);
						pagenumHeight = this.m_fontPageNumber.GetHeight(e.Graphics) * 2;
						dateHeight = this.m_fontReportDate.GetHeight(e.Graphics);
						columnHeight = this.m_fontColumns.GetHeight(e.Graphics);
						bodyHeight = this.m_fontBody.GetHeight(e.Graphics);

						// Used to store the position at which the next body line
						// should be printed.
						ypos = e.MarginBounds.Top;
						xpos = e.MarginBounds.Left;

						//print a line 
						e.Graphics.DrawLine(pen, e.MarginBounds.Left, e.MarginBounds.Top,
							e.MarginBounds.Right, e.MarginBounds.Top);

						// Determine the number of lines of body text that can be printed per page, taking
						// into account the presence of the header and the size of the selected body font.
						//each column is printed on the same line
						if (this.m_intPageNumber == 0)
						{
							//1st page of report so print the report header
							linesPerPage = (e.MarginBounds.Height - titleHeight - pagenumHeight-dateHeight-columnHeight)/ bodyHeight;
							//print the report title
							e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_title"].Rows[0]["title"].ToString(), this.m_fontReportTitle, Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top, new StringFormat());
							ypos = ypos + titleHeight; // +  this.m_fontBody.GetHeight(e.Graphics);
							//print the date
							e.Graphics.DrawString(System.DateTime.Now.ToShortDateString(), this.m_fontReportDate, Brushes.Black, e.MarginBounds.Left, ypos, new StringFormat());
							ypos = ypos + pagenumHeight;
							//draw a line
							e.Graphics.DrawLine(pen, e.MarginBounds.Left, ypos,
								e.MarginBounds.Right, ypos);
						}
						else
						{
							linesPerPage = (e.MarginBounds.Height - pagenumHeight-columnHeight)/ this.m_fontBody.GetHeight(e.Graphics);
						}
						//print the column header
						for (x=0; x<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;x++)
						{
							//get the width of the columns longest value
							sizeFWidth = e.Graphics.MeasureString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["largeststring"].ToString().Trim() + "***", this.m_fontColumns);
							//get the width of the columns name
							sizeFWidth2 = e.Graphics.MeasureString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim() + "***", this.m_fontColumns);
							intWidth = (int)sizeFWidth.Width;
							intWidth2 = (int)sizeFWidth2.Width;
							floatWidth = (float)sizeFWidth.Width;
							floatWidth2 = (float)sizeFWidth2.Width;
							//see if column width is wider than the widest column value
							//if (intWidth2 > intWidth)
							//if the column name width is greater than the longest value than use the column name width
							if (floatWidth2 > floatWidth)
							{
								//if this column name will print beyond the margin 
								//than drop this column from the report
								//if (xpos + intWidth2 <= e.MarginBounds.Width)
								if (xpos + floatWidth2 <= e.MarginBounds.Width)
								{
									//this is within the margins so lets keep this column
									//assign the beginning position of the column
									this.m_dsReportConfig.Tables["report_fields"].Rows[x]["Xpos"] = xpos;
									//print the column name
									e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim(), this.m_fontColumns, Brushes.Black, xpos, ypos, new StringFormat());
									//xpos = xpos + intWidth2;
									xpos = xpos + floatWidth2;
							
								}
								else 
								{
									break;
								}
							}
							else
							{
								//if the width of the longest value in the column
								//is greater than the margin width than drop
								//this column from the report
								//if (xpos + intWidth <= e.MarginBounds.Width)
								if (xpos + floatWidth <= e.MarginBounds.Width)
								{
									//this is within the margins so lets keep this column
									//assign the beginning position of the column
									this.m_dsReportConfig.Tables["report_fields"].Rows[x]["Xpos"] = xpos;
									//print the column name
									e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["field_name"].ToString().Trim(), this.m_fontColumns, Brushes.Black, xpos, ypos, new StringFormat());
									//xpos = xpos + intWidth;
									xpos = xpos + floatWidth;
							
								}
								else 
								{
									break;
								}
							}
						
						}
						ypos = ypos + columnHeight + 2;
					
					
						//draw a line
						e.Graphics.DrawLine(pen, e.MarginBounds.Left, ypos,
							e.MarginBounds.Right, ypos);
					
						if ((bool)this.m_bPrintingReportSummary == false && (bool)this.m_bSummaryDoneButNewPage == false &&
						    (bool)m_bDetailLineNewPage==false &&
							((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["print_group_summary_YN"]=="Y" ||
							 (string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["print_detail_line_YN"] =="Y"))
						{
							//if this is a new page than print the group fields
							if (this.m_intPageNumber > 0)
							{
								for (x=0;x<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;x++)
								{
									if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[x]["group"]==true)
									{
										e.Graphics.DrawString(this.m_dsReportConfig.Tables["report_fields"].Rows[x]["value"].ToString().Trim(), this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[x]["xpos"],ypos, new StringFormat());                                         
									}
								}
								linecount++;
								ypos = ypos + bodyHeight;
							}
						}

						this.m_intPageNumber++;
                    
						
					
						//loop through each row in the dataview
						for (x=this.m_intCurrRow; x<=this.m_dv.Table.Rows.Count-1; x++)
						{
							
							str="";
							bRecordPrinted=false;
							//lets see if the group has changed
							//
							for (y=0;y<=this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;y++)
							{
								if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[y]["group"]==true)
								{
									for (z=0;z<=this.m_dv.Table.Columns.Count-1;z++)
									{
										if (this.m_dsReportConfig.Tables["report_fields"].Rows[y]["field_name"].ToString().Trim().ToUpper() ==
											this.m_dv.Table.Columns[z].ColumnName.ToString().Trim().ToUpper())
										{

                                            if (this.m_dv.Table.Rows[x][z].ToString().Trim().Length == 0)
                                            {
                                                str += "null";
                                            }
                                            else
                                            {
                                                str += this.m_dv.Table.Rows[x][z].ToString().Trim();
                                            }		  
											
										}
									}
								}
								else
								{
									break;
								}
					    
							}
							if (str != this.m_strCurrGroup)
							{
								if (str.Trim().Length > 0 && this.m_strCurrGroup.Trim().Length > 0 && this.m_intGroupSummaryLineTotal > 0
								    && m_bDetailLineNewPage==false)
								{
									if (this.m_bSummaryDoneButNewPage == false)
									{
										PrintGroupSummary_Style1(ref sender, ref e, ref ypos, bodyHeight, pagenumHeight,ref linecount, linesPerPage);
										if (this.m_bPrintingGroupSummary==true || this.m_bSummaryDoneButNewPage==true)
										{
											
											this.m_intCurrRow=x;
											break;
										}
									}
									else 
									{
										this.m_bSummaryDoneButNewPage=false;
									}
								
								}
								this.m_bNewGroup=true;
							
							}
						
							//loop thru the report config data set and only
							//print the columns from the dataview that are in
							//the report config dataset
							for (y=this.m_intCurrRow2;y <= this.m_dsReportConfig.Tables["report_fields"].Rows.Count-1;y++)
							{
								//loop through the dataview columns
								for (z=0;z<=this.m_dv.Table.Columns.Count-1;z++)
								{
									//if the dataview column name equals the report config column name
									//then print the column
									if (this.m_dsReportConfig.Tables["report_fields"].Rows[y]["field_name"].ToString().Trim().ToUpper() ==
										this.m_dv.Table.Columns[z].ColumnName.ToString().Trim().ToUpper())
									{
								
										//all columns on the same row
										if (this.m_dsReportConfig.Tables["report_fields"].Rows[y]["xpos"].ToString().Trim() != "-1")
										{
											if ((bool)this.m_dsReportConfig.Tables["report_fields"].Rows[y]["group"]==true)
											{
												if (this.m_bNewGroup==true)
												{
													if (linecount + 3 >= linesPerPage)
												  {
												  	m_bDetailLineNewPage=true;
												  	linecount= linecount + 3;
												  	this.m_intCurrRow2 = y;
												  	break;
												  }
												  else
												  {
												  	m_bDetailLineNewPage=false;
												  }
													//print the group
													this.m_dsReportConfig.Tables["report_fields"].Rows[y]["value"] = this.m_dv.Table.Rows[x][z].ToString().Trim();
													e.Graphics.DrawString(this.m_dv.Table.Rows[x][z].ToString().Trim(), this.m_fontColumns,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"],ypos, new StringFormat());                                         
													bRecordPrinted=true;
												}
											}
											else 
											{
												//print the column value
												if (this.m_bNewGroup==true)
												{
													this.m_bNewGroup=false;
													this.m_strCurrGroup=str;
													linecount++;
													ypos = ypos + bodyHeight;
												}
												if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["print_detail_line_YN"] == "Y")
												{
													bRecordPrinted=true;
													//added start
													if (this.m_dv.Table.Rows[x][z] != System.DBNull.Value)
													{
														//print the column value
														if ((string)this.m_dsReportConfig.Tables["report_fields"].Rows[y]["stringdatatype_YN"] == "N" ||
															(bool)this.m_dsReportConfig.Tables["report_fields"].Rows[y]["candomath"] == true)
														{
															//right justify for numeric items
															System.Drawing.SizeF sizeFWidthLargestString = e.Graphics.MeasureString(this.m_dsReportConfig.Tables["report_fields"].Rows[y]["largeststring"].ToString().Trim(), this.m_fontBody);
															float floatWidthLargestString = (float)sizeFWidthLargestString.Width;
															string strValue = String.Format("{0:f3}",Math.Round(Convert.ToDecimal(this.m_dv.Table.Rows[x][z]),3));
															System.Drawing.SizeF fltValueF = e.Graphics.MeasureString(strValue, this.m_fontBody);
															float fltValue  = (float)fltValueF.Width;
															if (z==0)  //first column
															{
													
																e.Graphics.DrawString(strValue, this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"] + (float)((float)floatWidthLargestString - (float)fltValue),ypos, new StringFormat());                                         
													
															}
															else
															{
																e.Graphics.DrawString(strValue, this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"] + (float)((float)floatWidthLargestString - (float)fltValue),ypos, new StringFormat());                                         
																//e.Graphics.DrawString(strValue, this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"] + (float)floatWidth - fltValue,ypos, new StringFormat());                                         
															}
														}
														else
														{
															//left justify
															e.Graphics.DrawString(this.m_dv.Table.Rows[x][z].ToString().Trim(), this.m_fontBody,Brushes.Black, (float)this.m_dsReportConfig.Tables["Report_Fields"].Rows[y]["xpos"],ypos, new StringFormat());                                         
														}
													}
													
												}
											}
										}
									}
								}
								if (m_bDetailLineNewPage==true) break;
							}
						
							if (this.m_bNewGroup==true && m_bDetailLineNewPage==false)
							{
								
								this.m_bNewGroup=false;
								this.m_strCurrGroup=str;
							}
							if (bRecordPrinted==true)
							{
								linecount++;
								ypos = ypos + bodyHeight;
							}
						
						
							if (linecount >= linesPerPage)
							{
								//print the page in the middle of the page at the bottom
								sizeFWidth = e.Graphics.MeasureString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber);
						
								xpos = (e.PageBounds.Width/2)- (float)(sizeFWidth.Width * .50) ;
								ypos = e.PageBounds.Height - (pagenumHeight * 2);
								e.Graphics.DrawString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber, Brushes.Black, xpos, ypos, new StringFormat());
								//increment by one to print the new row on the next page
								this.m_intCurrRow=x+1;
								e.HasMorePages = true;
								break;
							}
							else
							{
								this.m_intCurrRow2=0;
								e.HasMorePages = false;
							}
						}
					
						//see about printing the report summary lines
						if (x > this.m_dv.Table.Rows.Count-1)
						{
							this.m_intCurrRow = x;
							if (this.m_intGroupSummaryLineTotal > 0 && this.m_bPrintingReportSummary == false)
							{
								PrintGroupSummary_Style1(ref sender, ref e, ref ypos, bodyHeight, pagenumHeight,ref linecount, linesPerPage);
							}
							if ((string)this.m_dsReportConfig.Tables["record_layout"].Rows[0]["print_report_totals_YN"] == "Y"  && this.m_bPrintingGroupSummary==false)
							{
								PrintSummary_Style1(ref sender, ref e, ref ypos, bodyHeight, pagenumHeight,ref linecount, linesPerPage);
							}
						}

						//this takes care of the last page printed
						if (e.HasMorePages == false)
						{
							ypos = ypos + (bodyHeight * 2);
							e.Graphics.DrawString("**End Of Report**", this.m_fontEndOfReport,Brushes.Black, e.MarginBounds.Left,ypos, new StringFormat());                                         
	
							sizeFWidth = e.Graphics.MeasureString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber);
							xpos = (e.PageBounds.Width/2)- (float)(sizeFWidth.Width * .50) ;
							ypos = e.PageBounds.Height - (pagenumHeight * 2);
							e.Graphics.DrawString("Page " + this.m_intPageNumber.ToString().Trim(), this.m_fontPageNumber, Brushes.Black, xpos, ypos, new StringFormat());
							this.m_intPageNumber=0;
							this.m_intCurrCol = 0;
							this.m_intCurrRow=0;
							this.m_intCurrRow2=0;
							this.m_strCurrGroup = "";
							str="";
						}
						break;
				}
			}
			catch (Exception errormsg)
			{
				
				MessageBox.Show(errormsg.ToString());
			}
		}
		private void ThermCancel(object sender, System.EventArgs e)
		{
			string strMsg = "Do you wish to cancel the report (Y/N)?";
			DialogResult result = MessageBox.Show(strMsg,"Cancel Report", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result) 
			{
				case DialogResult.Yes:
					this.m_frmTherm.AbortProcess = true;
					this.m_frmTherm.Hide();
					return;
				case DialogResult.No:
					return;
			}                
		}
			

		
		

	}

}
