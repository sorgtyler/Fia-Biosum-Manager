using System;
using System.Collections;
using System.ComponentModel.Design;
using System.ComponentModel;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public class macrosubst
	{
		private System.Data.DataSet _dsVarSub;
		private bool _bUseRandomData=true;
		private bool m_bInit=true;
		private string m_strNewString="";
		private string m_strVariableNames="";
		private SQLMacroSubstitutionVariable_Collection _oSQLMacroSubstitutionVariable_Collection;
		private SQLMacroSubstitutionVariableItem _oSQLMacroSubstitutionVariableItem;
        private GeneralMacroSubstitutionVariable_Collection _oGeneralMacroSubstitutionVariable_Collection;
        private GeneralMacroSubstitutionVariableItem _oGeneralMacroSubstitutionVariableItem;

		public macrosubst()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		public string MacroSubst(string str, System.Data.DataRow p_row)
		{
			if (this._bUseRandomData==true) return str;
			if (str.IndexOf("##",0) < 0 && str.IndexOf("!!",0) < 0) return str;
			return ExecuteMacroSubst(str,p_row);
		}
		public string MacroSubst(string str, int p_intRowNum, System.Data.DataRow[] p_rows)
		{
			if (this._bUseRandomData==true) return str;
			if (str.IndexOf("##",0) < 0 && str.IndexOf("!!",0) < 0) return str;
			return ExecuteMacroSubst(str, p_rows[p_intRowNum]);
		}
		public string MacroSubst(string str, int p_intRowNum, System.Data.DataTable p_dt)
		{
			if (this._bUseRandomData==true) return str;
			if (str.IndexOf("##",0) < 0 && str.IndexOf("!!",0) < 0) return str;
			return ExecuteMacroSubst(str, p_dt.Rows[p_intRowNum]); 
		}
		/// <summary>
		/// Get the value from a column in a ado datarow. The column can
		/// either be delimited by ## or !!.  If delimited by the value ## then
		/// the column value is returned.  If the column is delimited by the value
		/// of !! then the column value is translated. for example, if a 
		/// statecd had a value of !!41!! then the returned value would be Oregon.
		/// The translation of the !! value can also be dependent upon the value
		/// of multiple columns in the data row.  For example, each state has a 
		/// county code of 1. Therefore, the translation of the county code is 
		/// dependent upon the value that is in the statecd column.
		/// </summary>
		/// <param name="str"></param>
		/// <param name="p_row"></param>
		/// <returns></returns>
		private string ExecuteMacroSubst(string str,System.Data.DataRow p_row)
		{
			//check for no macros
			if (this._bUseRandomData==true) return str;
			if (str.IndexOf("##",0) < 0 && str.IndexOf("!!",0) < 0) return str;
			string strNewString="";
			string strVariable;
			try
			{
				
				for (int x=0; x<= str.Trim().Length-1;x++)
				{
					if (x==str.Trim().Length-1) 
					{
						strNewString=strNewString + str.Substring(x,1);
						break;
					}
					if (str.Substring(x,2)=="##" || str.Substring(x,2)=="!!") 
					{
						strVariable="";
						for (int y = x+2;y<=str.Trim().Length-1;y++)
						{
							if (str.Substring(y,2)=="##" || str.Substring(y,2)=="!!")
							{
								for (int z=0;z<=p_row.Table.Columns.Count-1;z++)
								{
									if (p_row.Table.Columns[z].ColumnName.ToString().Trim().ToUpper()==strVariable.Trim().ToUpper())
									{
										if (str.Substring(y,2)=="##")
										{
											strNewString=strNewString + p_row[strVariable].ToString().Trim();
											break;
										}
										else
										{
											//find the correct variable substitution field
											for (int xx=0;xx<=this._dsVarSub.Tables.Count-1;xx++)
											{
												
												if (_dsVarSub.Tables[xx].TableName.Trim().ToUpper()==
													strVariable.Trim().ToUpper())
												{
													//locate the current value and what it's substitute is
													for (int yy=0;yy<=_dsVarSub.Tables[xx].Rows.Count-1;yy++)
													{
														bool bFound=true;
														/******************************************************************
														 **compare the value in each column
														 ******************************************************************/ 
														for (int zzz=0;zzz<=_dsVarSub.Tables[xx].Columns.Count-1;zzz++)
														{
															//check the column value if it is not equal to the variable substituition column
															if (_dsVarSub.Tables[xx].Columns[zzz].ColumnName.Trim().ToUpper() !="VARSUB")
															{
																if (_dsVarSub.Tables[xx].Rows[yy][zzz].ToString().Trim() != 
																	p_row[_dsVarSub.Tables[xx].Columns[zzz].ColumnName.Trim()].ToString().Trim())
																{
																	bFound=false;
																	break;
																}
																
															}
														}
														if (bFound==true)
														{
															//	if (p_dt.Rows[p_intRowNum][strVariable].ToString().Trim()==
															//		_dsVarSub.Tables[xx].Rows[yy]["value"].ToString().Trim())
															//	{
															strNewString=strNewString + _dsVarSub.Tables[xx].Rows[yy]["varsub"].ToString().Trim();
															break;
															//	}
														}
													}
												}
											}
										}
									}
								}
								x=y+1;
								break;
							}
							else
							{
								strVariable=strVariable+str.Substring(y,1);
							}

						}
					}
					else
					{
						strNewString=strNewString + str.Substring(x,1);
				
					}
				
				}
			}
			catch
			{
				return strNewString;
			}
			return strNewString;
		}
        public System.Data.DataTable LoadVariableSubstitutionTextFile(string p_strFile)
        {
            System.Data.DataTable oDt;

            oDt = new System.Data.DataTable();

            //check if variable substitution file exists
            if (System.IO.File.Exists(p_strFile.Trim()))
            {
                //Open the file in a stream reader.
                System.IO.StreamReader s = new System.IO.StreamReader(
                    p_strFile.Trim());

                string strDelimiter = ",";

                //Split the first line into the columns       
                string[] columns = s.ReadLine().Split(strDelimiter.ToCharArray());


                //Cycle the colums, adding those that don't exist yet 
                //and sequencing the one that do.
                foreach (string col in columns)
                {
                    bool added = false;
                    string next = "";
                    int i = 0;
                    while (!added)
                    {
                        //Build the column name and remove any unwanted characters.
                        string columnname = col + next;
                        columnname = columnname.Replace("#", "");
                        columnname = columnname.Replace("'", "");
                        columnname = columnname.Replace("&", "");
                        columnname = columnname.Replace("\"", "");

                        //See if the column already exists
                        if (!oDt.Columns.Contains(columnname))
                        {
                            //if it doesn't then we add it here and mark it as added
                            oDt.Columns.Add(columnname);
                            added = true;
                        }
                        else
                        {
                            //if it did exist then we increment the sequencer and try again.
                            i++;
                            next = "_" + i.ToString();
                        }
                    }
                }


                //Read the rest of the data in the file.        
                string AllData = s.ReadToEnd();

                //Split off each row at the Carriage Return/Line Feed
                //Default line ending in most <A class=iAs style="FONT-WEIGHT: normal; FONT-SIZE: 100%; PADDING-BOTTOM: 1px; COLOR: darkgreen; BORDER-BOTTOM: darkgreen 0.07em solid; BACKGROUND-COLOR: transparent; TEXT-DECORATION: underline" href="#" target=_blank itxtdid="2592535">windows</A> exports.  
                //You may have to edit this to match your particular file.
                //This will work for Excel, Access, etc. default exports.
                string[] rows = AllData.Split("\r\n".ToCharArray());


                //Now add each row to the DataSet        
                foreach (string r in rows)
                {
                    if (r != null && r.Trim().Length > 0)
                    {
                        //Split the row at the delimiter.
                        string[] items = r.Split(strDelimiter.ToCharArray());

                        //Add the item
                        oDt.Rows.Add(items);
                    }
                }
            }
            return oDt;
        }

		/// <summary>
		/// Assign dataset that contains variable substitution data
		/// </summary>
		public System.Data.DataSet ReferenceVariableSubstitutionDataSet
		{
			set {_dsVarSub=value;}
		}
		public bool UseRandomData
		{
			set {_bUseRandomData=value;}
		}
		public  SQLMacroSubstitutionVariable_Collection  ReferenceSQLMacroSubstitutionVariableCollection
		{
			get {return _oSQLMacroSubstitutionVariable_Collection;}
			set {_oSQLMacroSubstitutionVariable_Collection=value;}
		}
		public SQLMacroSubstitutionVariableItem ReferenceSQLMacroSubstitutionVariableItem
		{
			get {return _oSQLMacroSubstitutionVariableItem;}
			set {_oSQLMacroSubstitutionVariableItem=value;}
		}
        public GeneralMacroSubstitutionVariable_Collection ReferenceGeneralMacroSubstitutionVariableCollection
        {
            get { return _oGeneralMacroSubstitutionVariable_Collection; }
            set { _oGeneralMacroSubstitutionVariable_Collection = value; }
        }
        public GeneralMacroSubstitutionVariableItem ReferenceGeneralMacroSubstitutionVariableItem
        {
            get { return _oGeneralMacroSubstitutionVariableItem; }
            set { _oGeneralMacroSubstitutionVariableItem = value; }
        }

		/// <summary>
		/// Translate SQL variable substitution variables to there value string
		/// </summary>
		/// <param name="str"></param>
		public string SQLTranslateVariableSubstitution(string str)
		{
			if (str.IndexOf("@@",0) < 0) return str;
			if (this.m_bInit) m_strNewString="";
			string strVariable;
			try
			{
				
				for (int x=0; x<= str.Trim().Length-1;x++)
				{
					if (x==str.Trim().Length-1) 
					{
						m_strNewString=m_strNewString + str.Substring(x,1);
						break;
					}
					if (str.Substring(x,2)=="@@") 
					{
						strVariable="";
						for (int y = x+2;y<=str.Trim().Length-1;y++)
						{
							if (str.Substring(y,2)=="@@")
							{
								//find the substitution variable in the collection
								for (int z=0;z<=ReferenceSQLMacroSubstitutionVariableCollection.Count-1;z++)
								{
									if (ReferenceSQLMacroSubstitutionVariableCollection.Item(z).VariableName.ToString().Trim().ToUpper()==strVariable.Trim().ToUpper())
									{
										if (str.Substring(y,2)=="@@")
										{
											if (ReferenceSQLMacroSubstitutionVariableCollection.Item(z).SQLVariableSubstitutionString.IndexOf("@@",0) >=0)
											{
												m_bInit=false;
												m_strNewString=SQLTranslateVariableSubstitution(ReferenceSQLMacroSubstitutionVariableCollection.Item(z).SQLVariableSubstitutionString);
												m_bInit=true;
											}
											else
											{
												m_strNewString=m_strNewString  + ReferenceSQLMacroSubstitutionVariableCollection.Item(z).SQLVariableSubstitutionString;
											}
											break;
										}
									}
								}
								x=y+1;
								break;
							}
							else
							{
								strVariable=strVariable+str.Substring(y,1);
							}

						}
					}
					else
					{
						m_strNewString=m_strNewString + str.Substring(x,1);
				
					}
				
				}
			}
			catch
			{
				return m_strNewString;
			}
			return m_strNewString;
		}

        public string GeneralTranslateVariableSubstitution(string str)
        {
            if (str.IndexOf("@@", 0) < 0) return str;
            if (this.m_bInit) m_strNewString = "";
            string strVariable;
            try
            {

                for (int x = 0; x <= str.Trim().Length - 1; x++)
                {
                    if (x == str.Trim().Length - 1)
                    {
                        m_strNewString = m_strNewString + str.Substring(x, 1);
                        break;
                    }
                    if (str.Substring(x, 2) == "@@")
                    {
                        strVariable = "";
                        for (int y = x + 2; y <= str.Trim().Length - 1; y++)
                        {
                            if (str.Substring(y, 2) == "@@")
                            {
                                //find the substitution variable in the collection
                                for (int z = 0; z <= ReferenceGeneralMacroSubstitutionVariableCollection.Count - 1; z++)
                                {
                                    if (ReferenceGeneralMacroSubstitutionVariableCollection.Item(z).VariableName.ToString().Trim().ToUpper() == strVariable.Trim().ToUpper())
                                    {
                                        if (str.Substring(y, 2) == "@@")
                                        {
                                            if (ReferenceGeneralMacroSubstitutionVariableCollection.Item(z).VariableSubstitutionString.IndexOf("@@", 0) >= 0)
                                            {
                                                m_bInit = false;
                                                m_strNewString = GeneralTranslateVariableSubstitution(ReferenceGeneralMacroSubstitutionVariableCollection.Item(z).VariableSubstitutionString);
                                                m_bInit = true;
                                            }
                                            else
                                            {
                                                m_strNewString = m_strNewString + ReferenceGeneralMacroSubstitutionVariableCollection.Item(z).VariableSubstitutionString;
                                            }
                                            break;
                                        }
                                    }
                                }
                                x = y + 1;
                                break;
                            }
                            else
                            {
                                strVariable = strVariable + str.Substring(y, 1);
                            }

                        }
                    }
                    else
                    {
                        m_strNewString = m_strNewString + str.Substring(x, 1);

                    }

                }
            }
            catch
            {
                return m_strNewString;
            }
            return m_strNewString;
        }

		/// <summary>
		/// get the column names that are embedded between the ## or !! tags. Return a comma delimited list
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public string GetColumnNames(string str)
		{
			bool bFoundEnd=false;
			//check for no macros
			if (str.IndexOf("##",0) < 0 && str.IndexOf("!!",0) < 0) return ""; //str;
			string strNewString="";
			string strVariable;
			try
			{
				
				for (int x=0; x<= str.Trim().Length-1;x++)
				{
					if (x==str.Trim().Length-1) 
					{
					}
					if (str.Substring(x,2)=="##" || str.Substring(x,2)=="!!") 
					{
						strVariable="";
						bFoundEnd=false;
						for (int y = x+2;y<=str.Trim().Length-1;y++)
						{
							if (str.Substring(y,2)=="##" || str.Substring(y,2)=="!!")
							{
								bFoundEnd=true;
								x=y+1;
								break;
							}
							else
							{
								strVariable=strVariable+str.Substring(y,1);
							}

						}
						if (bFoundEnd==true) strNewString+=strVariable + ",";

					}
					
				
				}
			}
			catch
			{
				return strNewString;
			}
			return strNewString;
		}

		/// <summary>
		/// get the Sql Variable names that are embedded between the @@ tags. Return a comma-delimited list.
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public string GetSQLVariableNames2(string str)
		{
			bool bFoundEnd=false;
			//check for no macros
			if (str.IndexOf("@@",0) < 0) return "";
			string strNewString="";
			string strVariable;
			try
			{
				
				for (int x=0; x<= str.Trim().Length-1;x++)
				{
					if (x==str.Trim().Length-1) 
					{
					}
					if (str.Substring(x,2)=="@@") 
					{
						strVariable="";
						bFoundEnd=false;
						for (int y = x+2;y<=str.Trim().Length-1;y++)
						{
							if (str.Substring(y,2)=="@@")
							{
								bFoundEnd=true;
								x=y+1;
								break;
							}
							else
							{
								strVariable=strVariable+str.Substring(y,1);
							}

						}
						if (bFoundEnd==true) strNewString+=strVariable + ",";

					}
					
				
				}
			}
			catch
			{
				return strNewString;
			}
			//remove the comma from the end of the string
			if (strNewString.Length > 0 && strNewString.Substring(strNewString.Length-1,1)==",") strNewString=strNewString.Substring(0,strNewString.Length-1);
			return strNewString;
		}
		/// <summary>
		/// get the Sql Variable names that are embedded between the @@ tags. Return a comma-delimited list.
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public string GetSQLVariableNames(string str)
		{
			bool bFoundEnd=false;
			if (str.IndexOf("@@",0) < 0) return "";
			if (this.m_bInit) 
			{
					m_strNewString="";
				    this.m_strVariableNames="";
			}
			string strVariable;
			try
			{
				
				for (int x=0; x<= str.Trim().Length-1;x++)
				{
					if (x==str.Trim().Length-1) 
					{
						m_strNewString=m_strNewString + str.Substring(x,1);
						break;
					}
					if (str.Substring(x,2)=="@@") 
					{
						strVariable="";
						bFoundEnd=false;
						for (int y = x+2;y<=str.Trim().Length-1;y++)
						{
							if (str.Substring(y,2)=="@@")
							{
								bFoundEnd=true;
								//find the substitution variable in the collection
								for (int z=0;z<=ReferenceSQLMacroSubstitutionVariableCollection.Count-1;z++)
								{
									if (ReferenceSQLMacroSubstitutionVariableCollection.Item(z).VariableName.ToString().Trim().ToUpper()==strVariable.Trim().ToUpper())
									{
										if (str.Substring(y,2)=="@@")
										{
											if (ReferenceSQLMacroSubstitutionVariableCollection.Item(z).SQLVariableSubstitutionString.IndexOf("@@",0) >=0)
											{
												m_bInit=false;
												this.m_strVariableNames=this.GetSQLVariableNames(ReferenceSQLMacroSubstitutionVariableCollection.Item(z).SQLVariableSubstitutionString);
												m_bInit=true;
											}
											else
											{
												m_strNewString=m_strNewString  + ReferenceSQLMacroSubstitutionVariableCollection.Item(z).SQLVariableSubstitutionString;
											}
											break;
										}
									}
								}
								x=y+1;
								break;
							}
							else
							{
								strVariable=strVariable+str.Substring(y,1);
							}

						}
						if (bFoundEnd==true) m_strVariableNames+=strVariable + ",";
						
					}
					else
					{
						m_strNewString=m_strNewString + str.Substring(x,1);
					}
				
				}
			}
			catch
			{
				return m_strVariableNames;
			}
			//remove the comma from the end of the string
			//if (m_strVariableNames.Length > 0 && m_strVariableNames.Substring(m_strVariableNames.Length-1,1)==",") m_strVariableNames=m_strVariableNames.Substring(0,m_strVariableNames.Length-1);
			return m_strVariableNames;
		}
	}
	/*********************************************************************************************************
	 **SQL MACRO SUBSTITUTION VARIABLES                           
     *********************************************************************************************************/
	public class SQLMacroSubstitutionVariableItem
	{
		private int _intIndex;
		[CategoryAttribute("General"),ReadOnly(true),DescriptionAttribute("SQL Variable Name")]
		public int Index
		{
			get {return _intIndex;}
			set {_intIndex = value;}
		}

		private string _strVariableName="";
		[CategoryAttribute("General"),DescriptionAttribute("SQL Variable Substitution Name")]
		public string VariableName
		{
			get {return _strVariableName;}
			set {_strVariableName=value;}
		}
		private string _strDesc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Description
		{
			get {return _strDesc;}
			set {_strDesc=value;}

		}
		private string _strSQLVariableSubstitutionString="";
		[CategoryAttribute("General"),DescriptionAttribute("SQL Variable Substition String")]
		public string SQLVariableSubstitutionString
		{
			get {return _strSQLVariableSubstitutionString;}
			set {_strSQLVariableSubstitutionString=value;}

		}
		private string _strPropertyGridName="";
		[CategoryAttribute("Estimation Engine And Excel"), BrowsableAttribute(false), DescriptionAttribute("row group index")]
		public string PropertyGridName
		{
			get {return _strPropertyGridName;}
			set {_strPropertyGridName=value;}
		}
		public void CopyProperties(FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem p_oVarSubItemSource,ref FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem  p_oVarSubItemDestination)
		{
		  p_oVarSubItemDestination.Index = p_oVarSubItemSource.Index;
			p_oVarSubItemDestination.Description = p_oVarSubItemSource.Description;
			p_oVarSubItemDestination.PropertyGridName=p_oVarSubItemSource.PropertyGridName;
			p_oVarSubItemDestination.SQLVariableSubstitutionString=p_oVarSubItemSource.SQLVariableSubstitutionString;
			p_oVarSubItemDestination.VariableName = p_oVarSubItemSource.VariableName;
		}

	}
	public class SQLMacroSubstitutionVariable_Collection : System.Collections.CollectionBase
	{
		public SQLMacroSubstitutionVariable_Collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem m_PropertiesSQLMacroSubstitutionVariableItem)
		{
			// vérify if object is not already in
			if (this.List.Contains(m_PropertiesSQLMacroSubstitutionVariableItem))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(m_PropertiesSQLMacroSubstitutionVariableItem);
 
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
		public FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem) List[Index];
		}

	}
    /*********************************************************************************************************
	 **GENERAL MACRO SUBSTITUTION VARIABLES                           
     *********************************************************************************************************/
    public class GeneralMacroSubstitutionVariableItem
    {
        private int _intIndex;
        [CategoryAttribute("General"), ReadOnly(true), DescriptionAttribute("Variable Name")]
        public int Index
        {
            get { return _intIndex; }
            set { _intIndex = value; }
        }

        private string _strVariableName = "";
        [CategoryAttribute("General"), DescriptionAttribute("Variable Substitution Name")]
        public string VariableName
        {
            get { return _strVariableName; }
            set { _strVariableName = value; }
        }
        private string _strDesc = "";
        [CategoryAttribute("General"), DescriptionAttribute("Description")]
        public string Description
        {
            get { return _strDesc; }
            set { _strDesc = value; }

        }
        private string _strVariableSubstitutionString = "";
        [CategoryAttribute("General"), DescriptionAttribute("Variable Substition String")]
        public string VariableSubstitutionString
        {
            get { return _strVariableSubstitutionString; }
            set { _strVariableSubstitutionString = value; }

        }
        private string _strPropertyGridName = "";
        [CategoryAttribute("Estimation Engine And Excel"), BrowsableAttribute(false), DescriptionAttribute("row group index")]
        public string PropertyGridName
        {
            get { return _strPropertyGridName; }
            set { _strPropertyGridName = value; }
        }
        public void CopyProperties(FIA_Biosum_Manager.GeneralMacroSubstitutionVariableItem p_oVarSubItemSource, ref FIA_Biosum_Manager.GeneralMacroSubstitutionVariableItem p_oVarSubItemDestination)
        {
            p_oVarSubItemDestination.Index = p_oVarSubItemSource.Index;
            p_oVarSubItemDestination.Description = p_oVarSubItemSource.Description;
            p_oVarSubItemDestination.PropertyGridName = p_oVarSubItemSource.PropertyGridName;
            p_oVarSubItemDestination.VariableSubstitutionString = p_oVarSubItemSource.VariableSubstitutionString;
            p_oVarSubItemDestination.VariableName = p_oVarSubItemSource.VariableName;
        }

    }
    public class GeneralMacroSubstitutionVariable_Collection : System.Collections.CollectionBase
    {
        public GeneralMacroSubstitutionVariable_Collection()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public void Add(FIA_Biosum_Manager.GeneralMacroSubstitutionVariableItem m_PropertiesGeneralMacroSubstitutionVariableItem)
        {
            // vérify if object is not already in
            if (this.List.Contains(m_PropertiesGeneralMacroSubstitutionVariableItem))
                throw new InvalidOperationException();

            // adding it
            this.List.Add(m_PropertiesGeneralMacroSubstitutionVariableItem);

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
        public FIA_Biosum_Manager.GeneralMacroSubstitutionVariableItem Item(int Index)
        {
            // The appropriate item is retrieved from the List object and
            // explicitly cast to the Widget type, then returned to the 
            // caller.
            return (FIA_Biosum_Manager.GeneralMacroSubstitutionVariableItem)List[Index];
        }

    }
}
