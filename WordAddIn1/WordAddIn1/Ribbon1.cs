using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private int m_maraks = 0;
        private const string m_markName = @"WTDMarks";
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_AddBookMark_Click(object sender, RibbonControlEventArgs e)
        {
            
            Document doc = getDocument();
            Debug.WriteLine(doc.Application.Selection.Start + " doc end :" + doc.Application.Selection.End);
            string strName = m_markName + (++m_maraks);
            doc.Bookmarks.Add(strName, doc.Application.Selection.Range);

            Debug.WriteLine(doc.Application.Selection.Start + " doc end :" + doc.Application.Selection.End);
        }

        private Document getDocument()
        {
            return Globals.ThisAddIn.Application.ActiveDocument;
        }
        private void btn_Location_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = getDocument();
            var list = doc.Bookmarks.Cast<Bookmark>().Where(x => x.Name.Contains(m_markName)).ToList();            
            
            foreach (var bookmark in list)
            {
                doc.Application.Selection.GoTo(WdGoToItem.wdGoToBookmark, Name: bookmark.Name);
                MessageBox.Show(bookmark.Range.Start.ToString() + "  end is  " + bookmark.Range.End.ToString());
            }
        }

        private void btn_Ident_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = getDocument();
            var list = doc.Paragraphs.Cast<Paragraph>().Reverse().ToList();
            foreach (Paragraph par in list)
            {
                string strText = par.Range.Text;
                if (strText.StartsWith("\t"))
                {
                    Range range = doc.Range(par.Range.Start, (par.Range.Start + 1));
                    Debug.WriteLine(range.Text + " start " + par.Range.Start + " end : " + par.Range.End);
                    range.Text = range.Text.Replace("\t", "    ");
                    Debug.WriteLine(range.Text + " start " + par.Range.Start + " end : " + par.Range.End);
                }
            }
        }
    }
}