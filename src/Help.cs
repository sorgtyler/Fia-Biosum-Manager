using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Xps.Packaging;
using System.Windows;
using System.Data;

namespace FIA_Biosum_Manager
{
    public class Help
    {
        private System.Data.DataSet m_dsHelp;
        private System.Data.DataTable m_dtHelp;
        XPSDocumentViewer xpsDocumentViewer = null;

        private int m_intCurrentPageNumber=-1;

        public const int COL_DEFAULT_HELP_PARENT = 0;
        public const int COL_DEFAULT_HELP_CHILD = 1;
        private string _strXPSFile = null;
        private env _env = null;

        static public string DefaultDatabaseXPSFile { get { return "DATABASE_Help.xps"; } }
        
        //call default constructor after initializing variables
        public Help(string p_strXPSFile, env p_env)
        {
            _strXPSFile = p_strXPSFile;
            _env = p_env;
            _strXSD = p_env.strAppDir + "\\Help\\Help.xsd";
            _strXML = p_env.strAppDir + "\\Help\\Help.xml";

            //default values
            _strPrimaryKeyColumns = new string[2];
            _strPrimaryKeyColumns[COL_DEFAULT_HELP_PARENT] = "PARENT_ITEM";
            _strPrimaryKeyColumns[COL_DEFAULT_HELP_CHILD] = "CHILD_ITEM";
            _strPrimaryKeyValues = new string[2];
            _strPrimaryKeyValues[COL_DEFAULT_HELP_PARENT] = "MAINWINDOW";
            _strPrimaryKeyValues[COL_DEFAULT_HELP_CHILD] = "MAINWINDOW";
            LoadXml();
 
        }

        public void LoadXml(string p_strXSDFile, string p_strXMLFile,string p_strTableName,string[] p_strPrimaryKeyColumns)
        {
            _strXSD = p_strXSDFile;
            _strXML = p_strXMLFile;
            _strTable = p_strTableName;
            _strPrimaryKeyColumns = p_strPrimaryKeyColumns;
            LoadXml();
           
        }
        public void LoadXml()
        {
            this.m_dsHelp = new System.Data.DataSet();
            this.m_dsHelp.ReadXmlSchema(XSDFile);
            this.m_dsHelp.ReadXml(XMLFile, System.Data.XmlReadMode.InferSchema);
            m_dtHelp = m_dsHelp.Tables[TableName];

            //define the primary key for the table
            if (PrimaryKeyColumns != null)
            {
                DataColumn[] colPk = new DataColumn[PrimaryKeyColumns.Length];
                for (int x = 0; x <= PrimaryKeyColumns.Length - 1; x++)
                {
                    colPk[x] = m_dtHelp.Columns[PrimaryKeyColumns[x]];
                }
                m_dtHelp.PrimaryKey = colPk;
            }

        }
        
        public void ShowHelp(string[] p_strPrimaryKeyValues)
        {

            System.Object[] oSearch = new Object[PrimaryKeyColumns.Length];
            for (int x = 0; x <= PrimaryKeyColumns.Length - 1; x++)
            {
                oSearch[x] = p_strPrimaryKeyValues[x];
            }
            //search for the record
            System.Data.DataRow oRow = m_dtHelp.Rows.Find(oSearch);
            if (oRow != null)
            {
                PageNumber = Convert.ToInt32(oRow["PAGENUMBER"]);
            }
            ShowHelp();

        }
        public void ShowHelp()
        {
            if (xpsDocumentViewer != null)
            {
                xpsDocumentViewer.Close();
                xpsDocumentViewer = null;
            }
           
            xpsDocumentViewer = new XPSDocumentViewer();
            System.Windows.Xps.Packaging.XpsDocument xpsDoc = new System.Windows.Xps.Packaging.XpsDocument(_env.strAppDir + "\\Help\\" + _strXPSFile, System.IO.FileAccess.Read);
            xpsDocumentViewer.xpsViewer1.Document = xpsDoc.GetFixedDocumentSequence();

            
            
            
            m_intCurrentPageNumber = PageNumber;
            System.Windows.Documents.DocumentPage oPage = xpsDocumentViewer.xpsViewer1.Document.DocumentPaginator.GetPage(PageNumber);
            FrameworkElement fe = oPage.Visual as FrameworkElement;
            fe.BringIntoView();

            xpsDocumentViewer.WindowState = WindowState.Normal;
            xpsDocumentViewer.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            xpsDocumentViewer.Topmost = true;
            xpsDocumentViewer.IsEnabled = true;
            xpsDocumentViewer.Show();
           
        }

        private int _intPageNumber=0;
        public int PageNumber
        {
            get { return _intPageNumber; }
            set { _intPageNumber = value; }
        }
        private string _strXSD = null;
        public string XSDFile
        {
            set {_strXSD=value;}
            get {return _strXSD;}
        }
        private string _strXML = null;
        public string XMLFile
        {
            set { _strXML = value; }
            get { return _strXML; }
        }
        private string _strTable = "BIOSUM_HELP";
        public string TableName
        {
            set { _strTable = value; }
            get { return _strTable; }
        }
        private string[] _strPrimaryKeyColumns = null;
        public string[] PrimaryKeyColumns
        {
            get {return _strPrimaryKeyColumns;}
            set {_strPrimaryKeyColumns=value;}
        }
        private string[] _strPrimaryKeyValues = null;
        public string[] PrimaryKeyValues
        {
            get { return _strPrimaryKeyValues; }
            set { _strPrimaryKeyValues = value; }
        }
        public string XPSFile
        {
            get { return _strXPSFile; }
            set { _strXPSFile = value; }
        }

        private string _strApplicationDirectory;
        public string ApplicationDirectory
        {
            get { return _strApplicationDirectory; }
            set { _strApplicationDirectory = value; }
        }
    }
}
