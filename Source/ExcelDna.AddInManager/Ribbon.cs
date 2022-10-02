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
        try
        {
            AddIn.GetController()?.OnInstall();
        }
        catch (Exception e)
        {
            ExceptionHandler.ShowException(e);
        }
    }

    public void OnButtonManagePressed(IRibbonControl control)
    {
        try
        {
            AddIn.GetController()?.OnManage();
        }
        catch (Exception e)
        {
            ExceptionHandler.ShowException(e);
        }
    }

    public void OnButtonOptionsPressed(IRibbonControl control)
    {
        try
        {
            AddIn.GetController()?.OnOptions();
        }
        catch (Exception e)
        {
            ExceptionHandler.ShowException(e);
        }
    }
}
