using System;
using System.Windows.Forms;

namespace OFFICE_LOCAL_AUTOSAVE_WORD
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            // Load the saved value in the box when opening
            numInterval.Value = Properties.Settings.Default.IntervalSeconds;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Save user's number to memory
            Properties.Settings.Default.IntervalSeconds = (int)numInterval.Value;
            Properties.Settings.Default.Save();

            // Update the timer immediately
            Globals.ThisAddIn.UpdateTimerInterval();

            // Close this dialogue box
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e) { }
    }
}
