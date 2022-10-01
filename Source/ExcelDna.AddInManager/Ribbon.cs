using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace ExcelDna.AddInManager;

[ComVisible(true)]
public class Ribbon : ExcelRibbon
{
    public override string GetCustomUI(string RibbonID)
    {
        return RibbonResources.Ribbon;
    }

    public override object? LoadImage(string imageId)
    {
        // This will return the image resource with the name specified in the image='xxxx' tag
        return RibbonResources.ResourceManager.GetObject(imageId);
    }

    public void OnButtonInstallPressed(IRibbonControl control)
    {
        string path = @"MyAddIn1-AddIn64-packed.xll";
        Controller.Install(path);
    }

    public void OnButtonUninstallPressed(IRibbonControl control)
    {
        string path = @"MyAddIn1-AddIn64-packed.xll";
        Controller.Uninstall(path);
    }
}
