using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PolicyPlus
{
    public partial class EditPol
    {
        public PolFile EditingPol;
        private bool EditingUserSource;

        public EditPol()
        {
            InitializeComponent();
        }
        public void UpdateTree()
        {
            // Repopulate the main list view, keeping the scroll position in the same place
            var topItemIndex = default(int?);
            if (LsvPol.TopItem is not null)
                topItemIndex = LsvPol.TopItem.Index;
            LsvPol.BeginUpdate();
            LsvPol.Items.Clear();
            
            // Use local function instead of lambda for recursive calls
            void AddKey(string Prefix, int Depth)
            {
                var subkeys = EditingPol.GetKeyNames(Prefix);
                subkeys.Sort(StringComparer.InvariantCultureIgnoreCase);
                foreach (var subkey in subkeys)
                {
                    string keypath = string.IsNullOrEmpty(Prefix) ? subkey : Prefix + @"\" + subkey;
                    var lsvi = LsvPol.Items.Add(subkey);
                    lsvi.IndentCount = Depth;
                    lsvi.ImageIndex = 0; // Folder
                    lsvi.Tag = keypath;
                    AddKey(keypath, Depth + 1);
                }
                var values = EditingPol.GetValueNames(Prefix, false);
                values.Sort(StringComparer.InvariantCultureIgnoreCase);
                var iconIndex = default(int);
                foreach (var value in values)
                {
                    if (string.IsNullOrEmpty(value))
                        continue;
                    var data = EditingPol.GetValue(Prefix, value);
                    var kind = EditingPol.GetValueKind(Prefix, value);
                    ListViewItem addToLsv(string ItemText, int Icon, bool Deletion)
                    {
                        var lsvItem = LsvPol.Items.Add(ItemText, Icon);
                        lsvItem.IndentCount = Depth;
                        var tag = new PolValueInfo() { Name = value, Key = Prefix };
                        if (Deletion)
                        {
                            tag.IsDeleter = true;
                        }
                        else
                        {
                            tag.Kind = kind;
                            tag.Data = data;
                        }
                        lsvItem.Tag = tag;
                        return lsvItem;
                    };
                    if (value.Equals("**deletevalues", StringComparison.InvariantCultureIgnoreCase))
                    {
                        addToLsv("Delete values", 8, true).SubItems.Add(data.ToString());
                    }
                    else if (value.StartsWith("**del.", StringComparison.InvariantCultureIgnoreCase))
                    {
                        addToLsv("Delete value", 8, true).SubItems.Add(value.Substring(6));
                    }
                    else if (value.StartsWith("**delvals", StringComparison.InvariantCultureIgnoreCase))
                    {
                        addToLsv("Delete all values", 8, true);
                    }
                    else
                    {
                        string text = "";
                        if (data is string[])
                        {
                            text = string.Join(" ", (string[])data);
                            iconIndex = 39; // Two pages
                        }
                        else if (data is string)
                        {
                            text = data.ToString();
                            iconIndex = kind == RegistryValueKind.ExpandString ? 42 : 40; // One page with arrow, or without
                        }
                        else if (data is uint)
                        {
                            text = data.ToString();
                            iconIndex = 15; // Calculator
                        }
                        else if (data is ulong)
                        {
                            text = data.ToString();
                            iconIndex = 41; // Calculator+
                        }
                        else if (data is byte[])
                        {
                            text = BitConverter.ToString((byte[])data).Replace("-", " ");
                            iconIndex = 13; // Gear
                        }
                        addToLsv(value, iconIndex, false).SubItems.Add(text);
                    }
                }
            }
            
            AddKey("", 0);
            LsvPol.EndUpdate();
            if (topItemIndex.HasValue && LsvPol.Items.Count > topItemIndex.Value)
                LsvPol.TopItem = LsvPol.Items[topItemIndex.Value];
        }
        public void PresentDialog(ImageList Images, PolFile Pol, bool IsUser)
        {
            LsvPol.SmallImageList = Images;
            EditingPol = Pol;
            EditingUserSource = IsUser;
            UpdateTree();
            ChItem.Width = LsvPol.ClientSize.Width - ChValue.Width - SystemInformation.VerticalScrollBarWidth;
            LsvPol_SelectedIndexChanged(null, null);
            ShowDialog();
        }
        private void ButtonSave_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }
        private void EditPol_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                DialogResult = DialogResult.Cancel;
        }
        public void SelectKey(string KeyPath)
        {
            var lsvi = LsvPol.Items.OfType<ListViewItem>().FirstOrDefault(i => i.Tag is string && KeyPath.Equals(i.Tag.ToString(), StringComparison.InvariantCultureIgnoreCase));
            if (lsvi is null)
                return;
            lsvi.Selected = true;
            lsvi.EnsureVisible();
        }
        public void SelectValue(string KeyPath, string ValueName)
        {
            var lsvi = LsvPol.Items.OfType<ListViewItem>().FirstOrDefault((Item) =>
            {
                if (!(Item.Tag is PolValueInfo))
                    return false;
                PolValueInfo pvi = (PolValueInfo)Item.Tag;
                return pvi.Key.Equals(KeyPath, StringComparison.InvariantCultureIgnoreCase) & pvi.Name.Equals(ValueName, StringComparison.InvariantCultureIgnoreCase);
            });
            if (lsvi is null)
                return;
            lsvi.Selected = true;
            lsvi.EnsureVisible();
        }
        public bool IsKeyNameValid(string Name)
        {
            return !Name.Contains(@"\");
        }
        public bool IsKeyNameAvailable(string ContainerPath, string Name)
        {
            return !EditingPol.GetKeyNames(ContainerPath).Any(k => k.Equals(Name, StringComparison.InvariantCultureIgnoreCase));
        }
        private void ButtonAddKey_Click(object sender, EventArgs e)
        {
            var editPolKeyForm = new EditPolKey();
            string keyName = editPolKeyForm.PresentDialog("");
            if (string.IsNullOrEmpty(keyName))
                return;
            if (!IsKeyNameValid(keyName))
            {
                MessageBox.Show("The key name is not valid.", "Invalid Key Name", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            string containerKey = LsvPol.SelectedItems.Count > 0 ? LsvPol.SelectedItems[0].Tag as string : "";
            if (!IsKeyNameAvailable(containerKey, keyName))
            {
                MessageBox.Show("The key name is already taken.", "Key Name Exists", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            string newPath = string.IsNullOrEmpty(containerKey) ? keyName : containerKey + @"\" + keyName;
            EditingPol.SetValue(newPath, "", Array.Empty<byte>(), RegistryValueKind.None);
            UpdateTree();
            SelectKey(newPath);
        }
        public object PromptForNewValueData(string ValueName, object CurrentData, RegistryValueKind Kind)
        {
            if (Kind == RegistryValueKind.String | Kind == RegistryValueKind.ExpandString)
            {
                var editPolStringDataForm = new EditPolStringData();
                if (editPolStringDataForm.PresentDialog(ValueName, CurrentData?.ToString()) == DialogResult.OK)
                {
                    return editPolStringDataForm.TextData.Text;
                }
                else
                {
                    return null;
                }
            }
            else if (Kind == RegistryValueKind.DWord | Kind == RegistryValueKind.QWord)
            {
                var editPolNumericDataForm = new EditPolNumericData();
                if (editPolNumericDataForm.PresentDialog(ValueName, Convert.ToUInt64(CurrentData), Kind == RegistryValueKind.QWord) == DialogResult.OK)
                {
                    return editPolNumericDataForm.NumData.Value;
                }
                else
                {
                    return null;
                }
            }
            else if (Kind == RegistryValueKind.MultiString)
            {
                var editPolMultiStringDataForm = new EditPolMultiStringData();
                if (editPolMultiStringDataForm.PresentDialog(ValueName, (string[])CurrentData) == DialogResult.OK)
                {
                    return editPolMultiStringDataForm.TextData.Lines;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                MessageBox.Show("This value kind is not supported.", "Unsupported Value Kind", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }
        }
        private void ButtonAddValue_Click(object sender, EventArgs e)
        {
            string keyPath = LsvPol.SelectedItems[0].Tag.ToString();
            var editPolValueForm = new EditPolValue();
            if (editPolValueForm.PresentDialog() != DialogResult.OK)
                return;
            string value = editPolValueForm.ChosenName;
            var kind = editPolValueForm.SelectedKind;
            object defaultData;
            if (kind == RegistryValueKind.String | kind == RegistryValueKind.ExpandString)
            {
                defaultData = "";
            }
            else if (kind == RegistryValueKind.DWord | kind == RegistryValueKind.QWord)
            {
                defaultData = 0;
            }
            else // Multi-string
            {
                defaultData = Array.CreateInstance(typeof(string), 0);
            }
            var newData = PromptForNewValueData(value, defaultData, kind);
            if (newData is not null)
            {
                EditingPol.SetValue(keyPath, value, newData, kind);
                UpdateTree();
                SelectValue(keyPath, value);
            }
        }
        private void ButtonDeleteValue_Click(object sender, EventArgs e)
        {
            var tag = LsvPol.SelectedItems[0].Tag;
            if (tag is string)
            {
                var editPolDeleteForm = new EditPolDelete();
                if (editPolDeleteForm.PresentDialog(tag.ToString().Split('\\').Last()) != DialogResult.OK)
                    return;
                if (editPolDeleteForm.OptPurge.Checked)
                {
                    EditingPol.ClearKey(tag.ToString()); // Delete everything
                }
                else if (editPolDeleteForm.OptClearFirst.Checked)
                {
                    EditingPol.ForgetKeyClearance(tag.ToString()); // So the clearance is before the values in the POL
                    EditingPol.ClearKey(tag.ToString());
                    // Add the existing values back
                    int index = LsvPol.SelectedIndices[0] + 1;
                    do
                    {
                        if (index >= LsvPol.Items.Count)
                            break;
                        var subItem = LsvPol.Items[index];
                        if (subItem.IndentCount <= LsvPol.SelectedItems[0].IndentCount)
                            break;
                        if (subItem.IndentCount == LsvPol.SelectedItems[0].IndentCount + 1 & subItem.Tag is PolValueInfo)
                        {
                            PolValueInfo valueInfo = (PolValueInfo)subItem.Tag;
                            if (!valueInfo.IsDeleter)
                                EditingPol.SetValue(valueInfo.Key, valueInfo.Name, valueInfo.Data, valueInfo.Kind);
                        }
                        index += 1;
                    }
                    while (true);
                }
                else
                {
                    // Delete only the specified value
                    EditingPol.DeleteValue(tag.ToString(), editPolDeleteForm.TextValueName.Text);
                }
                UpdateTree();
                SelectKey(tag.ToString());
            }
            else
            {
                PolValueInfo info = (PolValueInfo)tag;
                EditingPol.DeleteValue(info.Key, info.Name);
                UpdateTree();
                SelectValue(info.Key, "**del." + info.Name);
            }
        }
        private void ButtonForget_Click(object sender, EventArgs e)
        {
            string containerKey = "";
            var tag = LsvPol.SelectedItems[0].Tag;
            if (tag is string)
            {
                var result = MessageBox.Show("Are you sure you want to remove this key and all its contents?", "Remove Key", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (result == DialogResult.No)
                    return;
                string keyPath = tag as string;
                if (keyPath.Contains("\\"))
                    containerKey = keyPath.Remove(keyPath.LastIndexOf('\\'));
                // Use local function instead of lambda for recursive calls
                Action<string> RemoveKey = null;
                RemoveKey = (Key) =>
                {
                    foreach (var subkey in EditingPol.GetKeyNames(Key))
                        RemoveKey(Key + "\\" + subkey);
                    EditingPol.ClearKey(Key);
                    EditingPol.ForgetKeyClearance(Key);
                };
                RemoveKey(keyPath);
            }
            else
            {
                PolValueInfo info = (PolValueInfo)tag;
                containerKey = info.Key;
                EditingPol.ForgetValue(info.Key, info.Name);
            }
            UpdateTree();
            if (!string.IsNullOrEmpty(containerKey))
            {
                string[] pathParts = containerKey.Split(new[] { '\\' }, StringSplitOptions.None);
                for (int n = 1; n <= pathParts.Length; n++)
                    SelectKey(string.Join("\\", pathParts.Take(n)));
            }
            else
            {
                LsvPol_SelectedIndexChanged(null, null);
            } // Make sure the buttons don't stay enabled
        }
        private void ButtonEdit_Click(object sender, EventArgs e)
        {
            PolValueInfo info = (PolValueInfo)LsvPol.SelectedItems[0].Tag;
            var newData = PromptForNewValueData(info.Name, info.Data, info.Kind);
            if (newData is not null)
            {
                EditingPol.SetValue(info.Key, info.Name, newData, info.Kind);
                UpdateTree();
                SelectValue(info.Key, info.Name);
            }
        }
        private void LsvPol_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Update the enabled status of all the buttons
            if (LsvPol.SelectedItems.Count == 0)
            {
                ButtonAddKey.Enabled = true;
                ButtonAddValue.Enabled = false;
                ButtonDeleteValue.Enabled = false;
                ButtonEdit.Enabled = false;
                ButtonForget.Enabled = false;
                ButtonExport.Enabled = true;
            }
            else
            {
                var tag = LsvPol.SelectedItems[0].Tag;
                ButtonForget.Enabled = true;
                if (tag is string) // It's a key
                {
                    ButtonAddKey.Enabled = true;
                    ButtonAddValue.Enabled = true;
                    ButtonEdit.Enabled = false;
                    ButtonDeleteValue.Enabled = true;
                    ButtonExport.Enabled = true;
                }
                else // It's a value
                {
                    ButtonAddKey.Enabled = false;
                    ButtonAddValue.Enabled = false;
                    bool delete = ((PolValueInfo)tag).IsDeleter;
                    ButtonEdit.Enabled = !delete;
                    ButtonDeleteValue.Enabled = !delete;
                    ButtonExport.Enabled = false;
                }
            }
        }
        private void ButtonImport_Click(object sender, EventArgs e)
        {
            var importRegForm = new ImportReg();
            if (importRegForm.PresentDialog(EditingPol) == DialogResult.OK)
                UpdateTree();
        }
        private void ButtonExport_Click(object sender, EventArgs e)
        {
            string branch = "";
            if (LsvPol.SelectedItems.Count > 0)
                branch = LsvPol.SelectedItems[0].Tag.ToString();
            var exportRegForm = new ExportReg();
            exportRegForm.PresentDialog(branch, EditingPol, EditingUserSource);
        }
        private class PolValueInfo
        {
            public string Key;
            public string Name;
            public RegistryValueKind Kind;
            public object Data;
            public bool IsDeleter;
        }
    }
}