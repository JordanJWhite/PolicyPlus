using System;
using System.Windows.Forms;

namespace PolicyPlus
{
    public partial class ImportSpol
    {
        public SpolFile Spol;

        public ImportSpol()
        {
            InitializeComponent();
        }
        private void ButtonOpenFile_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Semantic Policy files|*.spol|All files|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    TextSpol.Text = System.IO.File.ReadAllText(ofd.FileName);
                }
            }
        }
        private void ButtonVerify_Click(object sender, EventArgs e)
        {
            try
            {
                var spol = SpolFile.FromText(TextSpol.Text);
                MessageBox.Show($"Validation successful, {spol.Policies.Count} policy settings found.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"SPOL validation failed: {ex.Message}", "Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void ButtonApply_Click(object sender, EventArgs e)
        {
            try
            {
                Spol = SpolFile.FromText(TextSpol.Text); // Tell the main form that the SPOL is ready to be committed
                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"The SPOL text is invalid: {ex.Message}", "Invalid SPOL", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void ImportSpol_Shown(object sender, EventArgs e)
        {
            Spol = null;
        }
        private void ImportSpol_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape & !(TextSpol.Focused & TextSpol.SelectionLength > 0))
                DialogResult = DialogResult.Cancel;
        }
        private void TextSpol_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.A & e.Control)
                TextSpol.SelectAll();
        }
        private void ButtonReset_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Are you sure you want to reset the text box?", "Reset Text", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                TextSpol.Text = "Policy Plus Semantic Policy" + System.Environment.NewLine + System.Environment.NewLine;
            }
        }
    }
}