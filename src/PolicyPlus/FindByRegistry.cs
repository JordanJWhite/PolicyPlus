using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PolicyPlus
{
    public partial class FindByRegistry
    {
        public Func<PolicyPlusPolicy, bool> Searcher;

        public FindByRegistry()
        {
            InitializeComponent();
        }
        private void FindByRegistry_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                DialogResult = DialogResult.Cancel;
        }
        private void SearchButton_Click(object sender, EventArgs e)
        {
            string keyName = KeyTextbox.Text.ToLowerInvariant();
            string valName = ValueTextbox.Text.ToLowerInvariant();
            if (string.IsNullOrEmpty(keyName) & string.IsNullOrEmpty(valName))
            {
                MessageBox.Show("Please enter search terms.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (new[] { @"HKLM\", @"HKCU\", @"HKEY_LOCAL_MACHINE\", @"HKEY_CURRENT_USER\" }.Any(bad => keyName.StartsWith(bad, StringComparison.InvariantCultureIgnoreCase)))
            {
                MessageBox.Show("Policies' root keys are determined only by their section. Remove the root key from the search terms and try again.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            Searcher = new Func<PolicyPlusPolicy, bool>((Policy) =>
                {
                    var affected = PolicyProcessing.GetReferencedRegistryValues(Policy);
                    
                    // Helper method to replace LikeOperator.LikeString
                    bool IsLike(string input, string pattern)
                    {
                        // Convert wildcard pattern to regex pattern
                        string regexPattern = "^" + Regex.Escape(pattern)
                            .Replace("\\*", ".*")
                            .Replace("\\?", ".") + "$";
                        return Regex.IsMatch(input, regexPattern);
                    }
                    
                    foreach (var rkvp in affected)
                    {
                        if (!string.IsNullOrEmpty(valName))
                        {
                            if (!IsLike(rkvp.Value.ToLowerInvariant(), valName))
                                continue;
                        }
                        if (!string.IsNullOrEmpty(keyName))
                        {
                            if (keyName.Contains("*") | keyName.Contains("?")) // Wildcard path
                            {
                                if (!IsLike(rkvp.Key.ToLowerInvariant(), keyName))
                                    continue;
                            }
                            else if (keyName.Contains(@"\")) // Path root
                            {
                                if (!rkvp.Key.StartsWith(keyName, StringComparison.InvariantCultureIgnoreCase))
                                    continue;
                            }
                            else if (!rkvp.Key.Split('\\').Any(part => part.Equals(keyName, StringComparison.InvariantCultureIgnoreCase))) // One path component
                                continue;
                        }
                        return true;
                    }
                    return false;
                });
            DialogResult = DialogResult.OK;
        }
    }
}