namespace ZTools.Standalone;

internal static class ProjectDialog
{
    internal static DialogResult Show(CommonDialog dialog)
    {
        // The workbench lives in a separate browser process. A dialog opened
        // on the STA worker otherwise has no owner and can hide behind it.
        using var owner=new Form
        {
            Text="凯乐提交文件工具",
            ShowInTaskbar=false,
            FormBorderStyle=FormBorderStyle.None,
            StartPosition=FormStartPosition.CenterScreen,
            Size=new Size(1,1),
            Opacity=0,
            TopMost=true
        };
        owner.Show();
        owner.Activate();
        return dialog.ShowDialog(owner);
    }
}
