using Microsoft.Office.Tools.Ribbon;

namespace OFFICE_LOCAL_AUTOSAVE_WORD
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e) { }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // Open settings window
            SettingsForm form = new SettingsForm();
            form.ShowDialog(); // Opens a dialogue box popup which you must interact with
        }
    }
}
