using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Printing;

namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Interaction logic for XPSDocumentViewer.xaml
    /// </summary>
    public partial class XPSDocumentViewer : System.Windows.Navigation.NavigationWindow
    {
        public XPSDocumentViewer()
        {
            InitializeComponent();

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
           
            if (ReferenceHelp != null) ReferenceHelp.xpsDocumentViewer = null;
           
        }

        private Help _ReferenceHelp = null;
        public Help ReferenceHelp
        {
            get { return _ReferenceHelp; }
            set { _ReferenceHelp = value; }
        }

        void PrintButtonClick(object sender, RoutedEventArgs e)
        {
            PrintDialog dlg = new PrintDialog();
            dlg.UserPageRangeEnabled = true;
           
            dlg.PrintQueue = LocalPrintServer.GetDefaultPrintQueue();
            dlg.PageRangeSelection = PageRangeSelection.UserPages;
            
           

            PageRange oRange = new PageRange(ReferenceHelp.xpsDocumentViewer.xpsViewer1.MasterPageNumber, ReferenceHelp.xpsDocumentViewer.xpsViewer1.MasterPageNumber);
            dlg.PageRange = oRange;
          
            
            if (dlg.ShowDialog().Value == true)
            {
                DocumentPaginator paginator = _ReferenceHelp.xpsDocumentViewer.xpsViewer1.Document.DocumentPaginator;


                if (dlg.PageRangeSelection == PageRangeSelection.UserPages)
                {
                    paginator = new PageRangeDocumentPaginator(
                                     _ReferenceHelp.xpsDocumentViewer.xpsViewer1.Document.DocumentPaginator,
                                     dlg.PageRange);
                }

                
                
                dlg.PrintDocument(paginator, "FIA BIOSUM Help");
              
            }
        }


        

    }
    /// <summary>
    /// Encapsulates a DocumentPaginator and allows
    /// to paginate just some specific pages (a "PageRange")
    /// of the encapsulated DocumentPaginator
    ///  (c) Thomas Claudius Huber 2010 
    ///      http://www.thomasclaudiushuber.com
    /// </summary>
    public class PageRangeDocumentPaginator : DocumentPaginator
    {
        private int _startIndex;
        private int _endIndex;
        private DocumentPaginator _paginator;
        public PageRangeDocumentPaginator(
          DocumentPaginator paginator,
          PageRange pageRange)
        {
            _startIndex = pageRange.PageFrom - 1;
            _endIndex = pageRange.PageTo - 1;
            _paginator = paginator;

            // Adjust the _endIndex
            _endIndex = Math.Min(_endIndex, _paginator.PageCount - 1);
        }
        public override DocumentPage GetPage(int pageNumber)
        {
            // Just return the page from the original
            // paginator by using the "startIndex"
            return _paginator.GetPage(pageNumber + _startIndex);
        }

        public override bool IsPageCountValid
        {
            get { return true; }
        }

        public override int PageCount
        {
            get
            {
                if (_startIndex > _paginator.PageCount - 1)
                    return 0;
                if (_startIndex > _endIndex)
                    return 0;

                return _endIndex - _startIndex + 1;
            }
        }

        public override Size PageSize
        {
            get { return _paginator.PageSize; }
            set { _paginator.PageSize = value; }
        }

        public override IDocumentPaginatorSource Source
        {
            get { return _paginator.Source; }
        }
    }

}
