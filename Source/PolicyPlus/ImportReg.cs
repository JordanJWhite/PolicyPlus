using System;
using System.Windows.Forms;

namespace PolicyPlus
{
    public partial class ImportReg
    {
        private IPolicySource PolicySource;

        public ImportReg()
        {
            InitializeComponent();
        }
        public DialogResult PresentDialog(IPolicySource Target)
        {
            TextReg.Text = "";
            TextRoot.Text = "";
            PolicySource = Target;
            return ShowDialog();
        }
        private void ButtonBrowse_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Registry scripts|*.reg";
                if (ofd.ShowDialog() != DialogResult.OK)
                    return;
                TextReg.Text = ofd.FileName;
                if (string.IsNullOrEmpty(TextRoot.Text))
                {
                    try
                    {
                        var reg = RegFile.Load(ofd.FileName, "");
                        TextRoot.Text = reg.GuessPrefix();
                        if (reg.HasDefaultValues())
                            MessageBox.Show("This REG file contains data for default values, which cannot be applied to all policy sources.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("An error occurred while trying to guess the prefix.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
        }
        private void ImportReg_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                DialogResult = DialogResult.Cancel;
        }
        private void ButtonImport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(TextReg.Text))
            {
                MessageBox.Show("Please specify a REG file to import.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (string.IsNullOrEmpty(TextRoot.Text))
            {
                MessageBox.Show("Please specify the prefix used to fully qualify paths in the REG file.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                var reg = RegFile.Load(TextReg.Text, TextRoot.Text);
                reg.Apply(PolicySource);
                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to import the REG file.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}