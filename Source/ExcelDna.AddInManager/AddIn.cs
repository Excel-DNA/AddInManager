using ExcelDna.Integration;

namespace ExcelDna.AddInManager;

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        Controller.AutoOpen();
    }

    public void AutoClose()
    {
    }
}

