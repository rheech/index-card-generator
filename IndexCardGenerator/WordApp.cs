using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace IndexCardGenerator
{
    class WordApp
    {
        private Application pApp;
        private Document pDoc;
        private Window pWnd;
        private Object vbNull = System.Reflection.Missing.Value;

        public WordApp(string template)
        {
            object oTemplate = template;

            pApp = new Microsoft.Office.Interop.Word.Application();
            pDoc = pApp.Documents.Add(ref oTemplate, ref vbNull, ref vbNull, ref vbNull);
            pWnd = pDoc.ActiveWindow;
        }

        public void TypeText(string value)
        {
            pWnd.Selection.TypeText(value);
        }

        public void GoToEnd()
        {
            object oBookMark = WdGoToItem.wdGoToBookmark;
            object oEndOfDoc = "\\EndOfDoc";

            pWnd.Selection.GoTo(ref oBookMark, ref vbNull, ref vbNull,
                        ref oEndOfDoc);
        }

        public void InsertPageBreak()
        {
            object oPageBreak = WdBreakType.wdPageBreak;

            pWnd.Selection.InsertBreak(ref oPageBreak);
        }

        public void SetAlignment(WdParagraphAlignment alignment)
        {
            pWnd.Selection.ParagraphFormat.Alignment = alignment;
        }

        public bool Visible
        {
            get
            {
                return pWnd.Visible;
            }
            set
            {
                pWnd.Visible = value;
            }
        }
    }
}
