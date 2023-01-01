using ChEJunkie.Office.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace ExcelSessionTests
{
    [TestClass]
    public class ExcelConductorTests
    {
        [TestMethod]
        public void Session_TracksNewWorkbook_Success()
        {
            // Arrange
            // Count existing Excel instances if any are open.
            Session session = Session.Current;
            int countStart = session.Applications.OfType<Excel.Application>().Count();
            int countUpdate;

            // Act
            // Open a new Excel instance and record.
            Excel.Application? app = null;
            try
            {
                app = new();
                app.Visible = true;
                Excel.Workbook wb = app.Workbooks.Add(app.Workbooks.Count + 1);
                Excel.Worksheet ws = wb.ActiveSheet;
                countUpdate = session.Applications.OfType<Excel.Application>().Count();
                wb.Close();
            }
            finally
            {
                app?.Quit();
            }

            // Assert
            // Check that Session succesfully tracked the new Excel instance.
            // Check that the starting state is returned.
            Assert.IsTrue(countUpdate > countStart);
            Assert.IsTrue(countStart == session.Applications.OfType<Excel.Application>().Count());
        }
    }
}