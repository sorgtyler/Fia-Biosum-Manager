using System;
using System.Data;
using System.IO;

namespace FIA_Biosum_Manager
{
  public class DelimitedTextToDataSet
  { 
 
    public DelimitedTextToDataSet()
    {  }
  
   /// <summary>
    /// Converts a given delimited file into a dataset. 
    /// Assumes that the first line    
  /// of the text file contains the column names.
  /// </summary>
    /// <param name="File">The name of the file to open</param>    
  /// <param name="TableName">The name of the 
  /// Table to be made within the DataSet returned</param>
    /// <param name="delimiter">The string to delimit by</param>
   /// <returns></returns>  
  public static DataSet Convert(string File, 
   string TableName, string delimiter)
    {   
    //The DataSet to Return
    DataSet result = new DataSet();
    
    //Open the file in a stream reader.
      StreamReader s = new StreamReader(File);
        
    //Split the first line into the columns       
     string[] columns = s.ReadLine().Split(delimiter.ToCharArray());
  
    //Add the new DataTable to the RecordSet
      result.Tables.Add(TableName);
    
    //Cycle the colums, adding those that don't exist yet 
    //and sequencing the one that do.
      foreach(string col in columns)
      {
        bool added = false;
        string next = "";
      int i = 0;
      while(!added)        
      {
        //Build the column name and remove any unwanted characters.
        string columnname = col + next;
         columnname = columnname.Replace("#","");
          columnname = columnname.Replace("'","");
          columnname = columnname.Replace("&","");
        
        //See if the column already exists
        if(!result.Tables[TableName].Columns.Contains(columnname))
          {
          //if it doesn't then we add it here and mark it as added
            result.Tables[TableName].Columns.Add(columnname);
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
    foreach(string r in rows)
      {
      //Split the row at the delimiter.
        string[] items = r.Split(delimiter.ToCharArray());
      
      //Add the item
       result.Tables[TableName].Rows.Add(items);  
      }
    
    //Return the imported data.        
    return result;
    }
  }
 }