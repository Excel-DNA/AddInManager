using ExcelDna.Integration;

namespace ExcelDna.AddInManager;

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        controller = new Controller();
    }

    public void AutoClose()
    {
    }

    internal static Controller? GetController()
    {
        return controller;
    }

    private static Controller? controller;
}

