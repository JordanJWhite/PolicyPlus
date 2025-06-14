﻿using System;
using System.Windows.Forms;

namespace PolicyPlus
{
    public partial class EditPolDelete
    {
        public EditPolDelete()
        {
            InitializeComponent();
        }
        public DialogResult PresentDialog(string ContainerKey)
        {
            OptClearFirst.Checked = false;
            OptDeleteOne.Checked = false;
            OptPurge.Checked = false;
            TextKey.Text = ContainerKey;
            TextValueName.Text = "";
            return ShowDialog();
        }
        private void ButtonOK_Click(object sender, EventArgs e)
        {
            if (OptClearFirst.Checked | OptPurge.Checked)
                DialogResult = DialogResult.OK;
            if (OptDeleteOne.Checked)
            {
                if (string.IsNullOrEmpty(TextValueName.Text))
                {
                    MessageBox.Show("You must enter a value name.", "Missing Value Name", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                DialogResult = DialogResult.OK;
            }
        }
        private void EditPolDelete_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                DialogResult = DialogResult.Cancel;
        }
        public void ChoiceChanged(object sender, EventArgs e)
        {
            TextValueName.Enabled = OptDeleteOne.Checked;
        }
    }
}