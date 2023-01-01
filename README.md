# Excel Conductor
.NET library to access any running Excel Application (single or multiple instances). Based on [CodeProject | Automate multiple Excel instances](https://www.codeproject.com/Articles/1157395/Automate-multiple-Excel-instances)

# Description
All running Excel instances are reliably returned, even workbooks opened from OneDrive. It extends the Excel type hierarchy by adding a top-level `Session` that contains `Applications` that contains `Workbooks` that contains `Sheets` etc.

# Example
`Applications` is the main access point. It returns an untyped sequence because `Excel.Application` is an embedded interop type. Elements can be safely cast to `Excel.Application` by the consuming project via the main `Session` class.

````csharp
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
````
Some extras, possibly useful, are also included. Extension methods for `Appication` and `Process` conversion.
