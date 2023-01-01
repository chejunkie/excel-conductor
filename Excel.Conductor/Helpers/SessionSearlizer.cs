using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ChEJunkie.Office.Excel
{
    /// <summary>
    /// Creates string representations of Excel objects, in a JSON-like format.
    /// </summary>
    public static class SessionSearlizer
    {
        #region Implementation

        private const string TAB = "   ";

        private static string[] SplitLines(string str) =>
            str.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

        private static string Indent(string str)
        {
            var sb = new StringBuilder();
            foreach (var line in SplitLines(str))
            {
                sb.AppendLine(TAB + line);
            }
            return sb.ToString();
        }

        #endregion Implementation

        public static string SerializeSession(Session session)
        {
            if (session == null)
            {
                throw new ArgumentNullException(nameof(session));
            }

            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine(TAB + "SessionID: " + session.SessionId);
            sb.Append(Indent(SerializeApps(session.Applications.Cast<Application>())));

            int? primaryId = null;
            try
            {
                primaryId = session.PrimaryInstance?.AsProcess().Id;
            }
            catch { }

            int? topMostId = null;
            try
            {
                topMostId = session.TopMost?.AsProcess().Id;
            }
            catch { }

            sb.AppendLine(TAB + "PrimaryProcessID: " + primaryId);
            sb.AppendLine(TAB + "TopMostProcessID: " + topMostId);
            sb.Append("}");
            return sb.ToString();
        }

        public static string SerializeApps(IEnumerable<Application> apps)
        {
            if (apps == null) throw new ArgumentNullException(nameof(apps));

            var sb = new StringBuilder();
            sb.AppendLine("Apps: [");
            foreach (Application app in apps)
            {
                sb.Append(Indent(SerializeApp(app)));
            }
            sb.AppendLine("]");
            return sb.ToString();
        }

        public static string SerializeApp(Application app)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine(TAB + "ProcessID: " + app.AsProcess().Id);
            sb.Append(Indent(SerializeBooks(app.Workbooks)));
            sb.AppendLine("}");
            return sb.ToString();
        }

        public static string SerializeBooks(Workbooks books)
        {
            if (books == null) throw new ArgumentNullException(nameof(books));

            var sb = new StringBuilder();
            sb.AppendLine("Books: [");
            foreach (Workbook book in books)
            {
                sb.Append(Indent(SerializeBook(book)));
            }
            sb.AppendLine("]");
            return sb.ToString();
        }

        public static string SerializeBook(Workbook book)
        {
            if (book == null) throw new ArgumentNullException(nameof(book));

            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine(TAB + "Name: " + book.Name);
            sb.Append(Indent(SerializeSheets(book.Sheets)));
            sb.AppendLine("}");
            return sb.ToString();
        }

        public static string SerializeSheets(Sheets sheets)
        {
            if (sheets == null) throw new ArgumentNullException(nameof(sheets));

            var sb = new StringBuilder();
            sb.AppendLine("Sheets: [");
            foreach (dynamic sheet in sheets)
            {
                sb.AppendLine(TAB + SerializeSheet(sheet));
            }
            sb.AppendLine("]");
            return sb.ToString();
        }

        public static string SerializeSheet(dynamic sheet)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }
            return "{ Name: " + sheet.Name + " }";
        }
    }
}