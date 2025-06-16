using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PolicyPlus
{
    public partial class Main
    {
        private ConfigurationStorage Configuration;
        private AdmxBundle AdmxWorkspace = new AdmxBundle();
        private IPolicySource UserPolicySource, CompPolicySource;
        private PolicyLoader UserPolicyLoader, CompPolicyLoader;
        private Dictionary<string, string> UserComments, CompComments;
        private PolicyPlusCategory CurrentCategory;
        private PolicyPlusPolicy CurrentSetting;
        private FilterConfiguration CurrentFilter = new FilterConfiguration();
        private PolicyPlusCategory HighlightCategory;
        private Dictionary<PolicyPlusCategory, TreeNode> CategoryNodes = new Dictionary<PolicyPlusCategory, TreeNode>();
        private bool ViewEmptyCategories = false;
        private AdmxPolicySection ViewPolicyTypes = AdmxPolicySection.Both;
        private bool ViewFilteredOnly = false;
        private bool PoliciesChanged = false;
        private FindResults findResultsDialog; // Added for FindResults instance

        public Main()
        {
            InitializeComponent();
        }

        // Helper method to get or create the FindResults dialog instance
        private FindResults GetFindResultsDialog()
        {
            if (findResultsDialog == null || findResultsDialog.IsDisposed)
            {
                findResultsDialog = new FindResults();
            }
            return findResultsDialog;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            // Create the configuration manager (for the Registry)
            Configuration = new ConfigurationStorage(RegistryHive.CurrentUser, @"Software\Policy Plus");
            // Restore the last ADMX source and policy loaders
            OpenLastAdmxSource();
            PolicyLoaderSource compLoaderType = (PolicyLoaderSource)Convert.ToInt32(Configuration.GetValue("CompSourceType", 0));
            var compLoaderData = Configuration.GetValue("CompSourceData", "");
            PolicyLoaderSource userLoaderType = (PolicyLoaderSource)Convert.ToInt32(Configuration.GetValue("UserSourceType", 0));
            var userLoaderData = Configuration.GetValue("UserSourceData", "");
            try
            {
                OpenPolicyLoaders(new PolicyLoader(userLoaderType, Convert.ToString(userLoaderData), true), new PolicyLoader(compLoaderType, Convert.ToString(compLoaderData), false), true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("The previous policy sources are not accessible. The defaults will be loaded.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Configuration.SetValue("CompSourceType", (int)PolicyLoaderSource.LocalGpo);
                Configuration.SetValue("UserSourceType", (int)PolicyLoaderSource.LocalGpo);
                OpenPolicyLoaders(new PolicyLoader(PolicyLoaderSource.LocalGpo, "", true), new PolicyLoader(PolicyLoaderSource.LocalGpo, "", false), true);
            }
            // Set up the UI
            ComboAppliesTo.Text = Convert.ToString(ComboAppliesTo.Items[0]);
            CategoriesTree.Height -= InfoStrip.ClientSize.Height;
            SettingInfoPanel.Height -= InfoStrip.ClientSize.Height;
            PoliciesList.Height -= InfoStrip.ClientSize.Height;
            InfoStrip.Items.Insert(2, new ToolStripSeparator());
            PopulateAdmxUi();
        }

        private void Main_Shown(object sender, EventArgs e)
        {
            // Check whether ADMX files probably need to be downloaded
            if (Convert.ToInt32(Configuration.GetValue("CheckedPolicyDefinitions", 0)) == 0)
            {
                Configuration.SetValue("CheckedPolicyDefinitions", 1);
                if (!SystemInfo.HasGroupPolicyInfrastructure() && AdmxWorkspace.Categories.Values.Where(c => IsOrphanCategory(c) & !IsEmptyCategory(c)).Count() > 2)
                {
                    if (MessageBox.Show($"Welcome to Policy Plus!{Environment.NewLine}{Environment.NewLine}Home editions do not come with the full set of policy definitions. Would you like to download them now? " + "This can also be done later with Help | Acquire ADMX Files.", "Policy Plus", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        AcquireADMXFilesToolStripMenuItem_Click(null, null);
                    }
                }
            }
        }

        public void OpenLastAdmxSource()
        {
            string defaultAdmxSource = Environment.ExpandEnvironmentVariables(@"%windir%\PolicyDefinitions");
            string admxSource = Convert.ToString(Configuration.GetValue("AdmxSource", defaultAdmxSource));
            try
            {
                var fails = AdmxWorkspace.LoadFolder(admxSource, GetPreferredLanguageCode());
                if (DisplayAdmxLoadErrorReport(fails, true) == DialogResult.No)
                    throw new Exception("You decided to not use the problematic ADMX bundle.");
            }
            catch (Exception ex)
            {
                AdmxWorkspace = new AdmxBundle();
                string loadFailReason = "";
                if ((admxSource ?? "") != (defaultAdmxSource ?? ""))
                {
                    if (MessageBox.Show("Policy definitions could not be loaded from \"" + admxSource + "\": " + ex.Message + Environment.NewLine + Environment.NewLine + "Load from the default location?", "Policy Plus", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        try
                        {
                            Configuration.SetValue("AdmxSource", defaultAdmxSource);
                            AdmxWorkspace = new AdmxBundle();
                            DisplayAdmxLoadErrorReport(AdmxWorkspace.LoadFolder(defaultAdmxSource, GetPreferredLanguageCode()));
                        }
                        catch (Exception ex2)
                        {
                            loadFailReason = ex2.Message;
                        }
                    }
                }
                else
                {
                    loadFailReason = ex.Message;
                }
                if (!string.IsNullOrEmpty(loadFailReason))
                    MessageBox.Show("Policy definitions could not be loaded: " + loadFailReason, "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public void PopulateAdmxUi()
        {
            // Populate the left categories tree
            CategoriesTree.Nodes.Clear();
            CategoryNodes.Clear();
            Action<IEnumerable<PolicyPlusCategory>, TreeNodeCollection> addCategory = null; // Initialize the variable
            addCategory = new Action<IEnumerable<PolicyPlusCategory>, TreeNodeCollection>((CategoryList, ParentNode) =>
            {
                foreach (var category in CategoryList.Where(ShouldShowCategory))
                {
                    var newNode = ParentNode.Add(category.UniqueID, category.DisplayName, GetImageIndexForCategory(category));
                    newNode.SelectedImageIndex = 3;
                    newNode.Tag = category;
                    CategoryNodes.Add(category, newNode);
                    addCategory(category.Children, newNode.Nodes);
                }
            }); // "Go" arrow
            addCategory(AdmxWorkspace.Categories.Values, CategoriesTree.Nodes);
            CategoriesTree.Sort();
            CurrentCategory = null;
            UpdateCategoryListing();
            ClearSelections();
            UpdatePolicyInfo();
        }

        public void UpdateCategoryListing()
        {
            // Update the right pane to include the current category's children
            var topItemIndex = default(int?);
            if (PoliciesList.TopItem is not null)
                topItemIndex = PoliciesList.TopItem.Index;
            bool inSameCategory = false;
            PoliciesList.Items.Clear();
            if (CurrentCategory is not null)
            {
                if (CurrentSetting is not null && ReferenceEquals(CurrentSetting.Category, CurrentCategory))
                    inSameCategory = true;
                if (CurrentCategory.Parent is not null) // Add the parent
                {
                    var listItem = PoliciesList.Items.Add("Up: " + CurrentCategory.Parent.DisplayName);
                    listItem.Tag = CurrentCategory.Parent;
                    listItem.ImageIndex = 6; // Up arrow
                    listItem.SubItems.Add("Parent");
                }
                foreach (var category in CurrentCategory.Children.Where(ShouldShowCategory).OrderBy(c => c.DisplayName)) // Add subcategories
                {
                    var listItem = PoliciesList.Items.Add(category.DisplayName);
                    listItem.Tag = category;
                    listItem.ImageIndex = GetImageIndexForCategory(category);
                }
                foreach (var policy in CurrentCategory.Policies.Where(ShouldShowPolicy).OrderBy(p => p.DisplayName)) // Add policies
                {
                    var listItem = PoliciesList.Items.Add(policy.DisplayName);
                    listItem.Tag = policy;
                    listItem.ImageIndex = GetImageIndexForSetting(policy);
                    listItem.SubItems.Add(GetPolicyState(policy));
                    listItem.SubItems.Add(GetPolicyCommentText(policy));
                    if (ReferenceEquals(policy, CurrentSetting)) // Keep the current policy selected
                    {
                        listItem.Selected = true;
                        listItem.Focused = true;
                        listItem.EnsureVisible();
                    }
                }
                if (topItemIndex.HasValue & inSameCategory) // Minimize the list view's jumping around when refreshing
                {
                    if (PoliciesList.Items.Count > topItemIndex.Value)
                        PoliciesList.TopItem = PoliciesList.Items[topItemIndex.Value];
                }
                if (CategoriesTree.SelectedNode is null || !ReferenceEquals(CategoriesTree.SelectedNode.Tag, CurrentCategory)) // Update the tree view
                {
                    CategoriesTree.SelectedNode = CategoryNodes[CurrentCategory];
                }
            }
        }

        public void UpdatePolicyInfo()
        {
            // Update the middle pane with the selected object's information
            bool hasCurrentSetting = CurrentSetting is not null | HighlightCategory is not null | CurrentCategory is not null;
            PolicyTitleLabel.Visible = hasCurrentSetting;
            PolicySupportedLabel.Visible = hasCurrentSetting;
            if (CurrentSetting is not null)
            {
                PolicyTitleLabel.Text = CurrentSetting.DisplayName;
                if (CurrentSetting.SupportedOn is null)
                {
                    PolicySupportedLabel.Text = "(no requirements information)";
                }
                else
                {
                    PolicySupportedLabel.Text = "Requirements:" + Environment.NewLine + CurrentSetting.SupportedOn.DisplayName;
                }
                PolicyDescLabel.Text = PrettifyDescription(CurrentSetting.DisplayExplanation);
                PolicyIsPrefTable.Visible = IsPreference(CurrentSetting);
            }
            else if (HighlightCategory is not null | CurrentCategory is not null)
            {
                var shownCategory = HighlightCategory ?? CurrentCategory;
                PolicyTitleLabel.Text = shownCategory.DisplayName;
                PolicySupportedLabel.Text = (HighlightCategory is null ? "This" : "The selected") + " category contains " + shownCategory.Policies.Count + " policies and " + shownCategory.Children.Count + " subcategories.";
                PolicyDescLabel.Text = PrettifyDescription(shownCategory.DisplayExplanation);
                PolicyIsPrefTable.Visible = false;
            }
            else
            {
                PolicyDescLabel.Text = "Select an item to see its description.";
                PolicyIsPrefTable.Visible = false;
            }
            SettingInfoPanel_ClientSizeChanged(null, null);
        }

        public bool IsOrphanCategory(PolicyPlusCategory Category)
        {
            return Category.Parent is null & !string.IsNullOrEmpty(Category.RawCategory.ParentID);
        }

        public bool IsEmptyCategory(PolicyPlusCategory Category)
        {
            return Category.Children.Count == 0 & Category.Policies.Count == 0;
        }

        public int GetImageIndexForCategory(PolicyPlusCategory Category)
        {
            if (IsOrphanCategory(Category))
            {
                return 1; // Orphaned
            }
            else if (IsEmptyCategory(Category))
            {
                return 2; // Empty
            }
            else
            {
                return 0;
            } // Normal
        }

        public int GetImageIndexForSetting(PolicyPlusPolicy Setting)
        {
            if (IsPreference(Setting))
            {
                return 7; // Preference, not policy (exclamation mark)
            }
            else if (Setting.RawPolicy.Elements is null || Setting.RawPolicy.Elements.Count == 0)
            {
                return 4; // Normal
            }
            else
            {
                return 5;
            } // Extra configuration
        }

        public bool ShouldShowCategory(PolicyPlusCategory Category)
        {
            // Should this category be shown considering the current filter?
            if (ViewEmptyCategories)
            {
                return true;
            }
            else // Only if it has visible children
            {
                return Category.Policies.Any(ShouldShowPolicy) || Category.Children.Any(ShouldShowCategory);
            }
        }

        public bool ShouldShowPolicy(PolicyPlusPolicy Policy)
        {
            // Should this policy be shown considering the current filter and active sections?
            if (!PolicyVisibleInSection(Policy, ViewPolicyTypes))
                return false;
            if (ViewFilteredOnly)
            {
                bool visibleAfterFilter = false;
                if ((int)(ViewPolicyTypes & AdmxPolicySection.Machine) > 0 & PolicyVisibleInSection(Policy, AdmxPolicySection.Machine))
                {
                    if (IsPolicyVisibleAfterFilter(Policy, false))
                        visibleAfterFilter = true;
                }
                else if ((int)(ViewPolicyTypes & AdmxPolicySection.User) > 0 & PolicyVisibleInSection(Policy, AdmxPolicySection.User))
                {
                    if (IsPolicyVisibleAfterFilter(Policy, true))
                        visibleAfterFilter = true;
                }
                if (!visibleAfterFilter)
                    return false;
            }
            return true;
        }

        public void MoveToVisibleCategoryAndReload()
        {
            // Move up in the categories tree until a visible one is found
            var newFocusCategory = CurrentCategory;
            var newFocusPolicy = CurrentSetting;
            while (!(newFocusCategory is null) && !ShouldShowCategory(newFocusCategory))
            {
                newFocusCategory = newFocusCategory.Parent;
                newFocusPolicy = null;
            }
            if (newFocusPolicy is not null && !ShouldShowPolicy(newFocusPolicy))
                newFocusPolicy = null;
            PopulateAdmxUi();
            CurrentCategory = newFocusCategory;
            UpdateCategoryListing();
            CurrentSetting = newFocusPolicy;
            UpdatePolicyInfo();
        }

        public string GetPolicyState(PolicyPlusPolicy Policy)
        {
            // Get a human-readable string describing the status of a policy, considering all active sections
            if (ViewPolicyTypes == AdmxPolicySection.Both)
            {
                string userState = GetPolicyState(Policy, AdmxPolicySection.User);
                string machState = GetPolicyState(Policy, AdmxPolicySection.Machine);
                var section = Policy.RawPolicy.Section;
                if (section == AdmxPolicySection.Both)
                {
                    if ((userState ?? "") == (machState ?? ""))
                    {
                        return userState + " (2)";
                    }
                    else if (userState == "Not Configured")
                    {
                        return machState + " (C)";
                    }
                    else if (machState == "Not Configured")
                    {
                        return userState + " (U)";
                    }
                    else
                    {
                        return "Mixed";
                    }
                }
                else if (section == AdmxPolicySection.Machine)
                    return machState + " (C)";
                else
                    return userState + " (U)";
            }
            else
            {
                return GetPolicyState(Policy, ViewPolicyTypes);
            }
        }

        public string GetPolicyState(PolicyPlusPolicy Policy, AdmxPolicySection Section)
        {
            // Get the human-readable status of a policy considering only one section
            switch (PolicyProcessing.GetPolicyState(Section == AdmxPolicySection.Machine ? CompPolicySource : UserPolicySource, Policy))
            {
                case PolicyState.Disabled:
                    {
                        return "Disabled";
                    }
                case PolicyState.Enabled:
                    {
                        return "Enabled";
                    }
                case PolicyState.NotConfigured:
                    {
                        return "Not Configured";
                    }

                default:
                    {
                        return "Unknown";
                    }
            }
        }

        public string GetPolicyCommentText(PolicyPlusPolicy Policy)
        {
            // Get the comment text to show in the Comment column, considering all active sections
            if (ViewPolicyTypes == AdmxPolicySection.Both)
            {
                string userComment = GetPolicyComment(Policy, AdmxPolicySection.User);
                string compComment = GetPolicyComment(Policy, AdmxPolicySection.Machine);
                if (string.IsNullOrEmpty(userComment) & string.IsNullOrEmpty(compComment))
                {
                    return "";
                }
                else if (!string.IsNullOrEmpty(userComment) & !string.IsNullOrEmpty(compComment))
                {
                    return "(multiple)";
                }
                else if (!string.IsNullOrEmpty(userComment))
                {
                    return userComment;
                }
                else
                {
                    return compComment;
                }
            }
            else
            {
                return GetPolicyComment(Policy, ViewPolicyTypes);
            }
        }

        public string GetPolicyComment(PolicyPlusPolicy Policy, AdmxPolicySection Section)
        {
            // Get a policy's comment in one section
            var commentSource = Section == AdmxPolicySection.Machine ? CompComments : UserComments;
            if (commentSource is null)
            {
                return "";
            }
            else if (commentSource.ContainsKey(Policy.UniqueID))
                return commentSource[Policy.UniqueID];
            else
                return "";
        }

        public bool IsPreference(PolicyPlusPolicy Policy)
        {
            return !string.IsNullOrEmpty(Policy.RawPolicy.RegistryKey) & !RegistryPolicyProxy.IsPolicyKey(Policy.RawPolicy.RegistryKey);
        }

        public void ShowSettingEditor(PolicyPlusPolicy Policy, AdmxPolicySection Section)
        {
            using (var editSettingForm = new EditSetting())
            {
                if (editSettingForm.PresentDialog(Policy, Section, AdmxWorkspace, CompPolicySource, UserPolicySource, CompPolicyLoader, UserPolicyLoader, CompComments, UserComments) == DialogResult.OK)
                {
                    PoliciesChanged = true;
                    // Keep the selection where it is if possible
                    if (CurrentCategory is null || ShouldShowCategory(CurrentCategory))
                        UpdateCategoryListing();
                    else
                        MoveToVisibleCategoryAndReload();
                }
            }
        }

        public void ClearSelections()
        {
            CurrentSetting = null;
            HighlightCategory = null;
        }

        public void OpenPolicyLoaders(PolicyLoader User, PolicyLoader Computer, bool Quiet)
        {
            // Create policy sources from the given loaders
            if (CompPolicyLoader is not null | UserPolicyLoader is not null)
                ClosePolicySources();
            UserPolicyLoader = User;
            UserPolicySource = User.OpenSource();
            CompPolicyLoader = Computer;
            CompPolicySource = Computer.OpenSource();
            bool allOk = true;
            string policyStatus(PolicyLoader Loader)
            {
                switch (Loader.GetWritability())
                {
                    case PolicySourceWritability.Writable:
                        {
                            return "is fully writable";
                        }
                    case PolicySourceWritability.NoCommit:
                        {
                            allOk = false;
                            return "cannot be saved";
                        }
                    default:
                        {
                            allOk = false;
                            return "cannot be modified";
                        }
                }
            }; // No writing
            Dictionary<string, string> loadComments(PolicyLoader Loader)
            {
                string cmtxPath = Loader.GetCmtxPath();
                if (string.IsNullOrEmpty(cmtxPath))
                {
                    return null;
                }
                else
                {
                    try
                    {
                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(cmtxPath));
                        if (System.IO.File.Exists(cmtxPath))
                        {
                            return CmtxFile.Load(cmtxPath).ToCommentTable();
                        }
                        else
                        {
                            return new Dictionary<string, string>();
                        }
                    }
                    catch (Exception ex)
                    {
                        return null;
                    }
                }
            };
            string userStatus = policyStatus(User);
            string compStatus = policyStatus(Computer);
            UserComments = loadComments(User);
            CompComments = loadComments(Computer);
            UserSourceLabel.Text = UserPolicyLoader.GetDisplayInfo();
            ComputerSourceLabel.Text = CompPolicyLoader.GetDisplayInfo();
            if (allOk)
            {
                if (!Quiet)
                {
                    MessageBox.Show("Both the user and computer policy sources are loaded and writable.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                string msgText = "Not all policy sources are fully writable." + Environment.NewLine + Environment.NewLine + "The user source " + userStatus + "." + Environment.NewLine + Environment.NewLine + "The computer source " + compStatus + ".";
                MessageBox.Show(msgText, "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public void ClosePolicySources()
        {
            // Clean up the policy sources
            bool allOk = true;
            if (UserPolicyLoader is not null)
            {
                if (!UserPolicyLoader.Close())
                    allOk = false;
            }
            if (CompPolicyLoader is not null)
            {
                if (!CompPolicyLoader.Close())
                    allOk = false;
            }
            if (!allOk)
            {
                MessageBox.Show("Cleanup did not complete fully because the loaded resources are open in other programs.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private bool ConfirmSaveChanges()
        {
            if (!PoliciesChanged)
                return true;
            var result = MessageBox.Show("Do you want to save changes to your policies?", "Policy Plus", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Cancel)
                return false;
            if (result == DialogResult.Yes)
            {
                SavePoliciesToolStripMenuItem_Click(null, null);
            }
            PoliciesChanged = false;
            return true;
        }

        public void ShowSearchDialog(Func<PolicyPlusPolicy, bool> Searcher)
        {
            DialogResult result;
            PolicyPlusPolicy selPol = null;

            var currentFindResultsDialog = GetFindResultsDialog(); // Use the instance
            if (Searcher is null)
            {
                result = currentFindResultsDialog.PresentDialog();
            }
            else
            {
                result = currentFindResultsDialog.PresentDialogStartSearch(AdmxWorkspace, Searcher);
            }
            if (result == DialogResult.OK)
            {
                selPol = currentFindResultsDialog.SelectedPolicy;
            }

            if (selPol != null)
            {
                ShowSettingEditor(selPol, ViewPolicyTypes);
                FocusPolicy(selPol);
            }
        }

        public void ClearAdmxWorkspace()
        {
            AdmxWorkspace = new AdmxBundle();
            GetFindResultsDialog().ClearSearch(); // Use the instance
        }

        public void FocusPolicy(PolicyPlusPolicy Policy)
        {
            // Try to automatically select a policy in the list view
            if (Policy.Category != null && CategoryNodes.ContainsKey(Policy.Category)) // Added null check for Policy.Category
            {
                CurrentCategory = Policy.Category;
                UpdateCategoryListing();
                foreach (ListViewItem entry in PoliciesList.Items)
                {
                    if (ReferenceEquals(entry.Tag, Policy))
                    {
                        entry.Selected = true;
                        entry.Focused = true;
                        entry.EnsureVisible();
                        break;
                    }
                }
            }
        }

        public bool IsPolicyVisibleAfterFilter(PolicyPlusPolicy Policy, bool IsUser)
        {
            // Calculate whether a policy is visible with the current filter
            if (CurrentFilter.ManagedPolicy.HasValue)
            {
                if (IsPreference(Policy) == CurrentFilter.ManagedPolicy.Value)
                    return false;
            }
            if (CurrentFilter.PolicyState.HasValue)
            {
                var policyState = PolicyProcessing.GetPolicyState(IsUser ? UserPolicySource : CompPolicySource, Policy);
                switch (CurrentFilter.PolicyState.Value)
                {
                    case FilterPolicyState.Configured:
                        {
                            if (policyState == PolicyState.NotConfigured)
                                return false;
                            break;
                        }
                    case FilterPolicyState.NotConfigured:
                        {
                            if (policyState != PolicyState.NotConfigured)
                                return false;
                            break;
                        }
                    case FilterPolicyState.Disabled:
                        {
                            if (policyState != PolicyState.Disabled)
                                return false;
                            break;
                        }
                    case FilterPolicyState.Enabled:
                        {
                            if (policyState != PolicyState.Enabled)
                                return false;
                            break;
                        }
                }
            }
            if (CurrentFilter.Commented.HasValue)
            {
                var commentDict = IsUser ? UserComments : CompComments;
                if ((commentDict.ContainsKey(Policy.UniqueID) && !string.IsNullOrEmpty(commentDict[Policy.UniqueID])) != CurrentFilter.Commented.Value)
                    return false;
            }
            if (CurrentFilter.AllowedProducts is not null)
            {
                if (!PolicyProcessing.IsPolicySupported(Policy, CurrentFilter.AllowedProducts, CurrentFilter.AlwaysMatchAny, CurrentFilter.MatchBlankSupport))
                    return false;
            }
            return true;
        }

        public bool PolicyVisibleInSection(PolicyPlusPolicy Policy, AdmxPolicySection Section)
        {
            // Does this policy apply to the given section?
            return (int)(Policy.RawPolicy.Section & Section) > 0;
        }

        public PolFile GetOrCreatePolFromPolicySource(IPolicySource Source)
        {
            if (Source is PolFile)
            {
                // If it's already a POL, just save it
                return (PolFile)Source;
            }
            else if (Source is RegistryPolicyProxy)
            {
                // Recurse through the Registry branch and create a POL
                var regRoot = ((RegistryPolicyProxy)Source).EncapsulatedRegistry;
                var pol = new PolFile();
                Action<string, RegistryKey> addSubtree = null; // Initialize the variable
                addSubtree = new Action<string, RegistryKey>((PathRoot, Key) =>
                {
                    foreach (var valName in Key.GetValueNames())
                    {
                        var valData = Key.GetValue(valName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                        pol.SetValue(PathRoot, valName, valData, Key.GetValueKind(valName));
                    }
                    foreach (var subkeyName in Key.GetSubKeyNames())
                    {
                        using (var subkey = Key.OpenSubKey(subkeyName, false))
                        {
                            addSubtree(PathRoot + @"\" + subkeyName, subkey);
                        }
                    }
                });
                foreach (var policyPath in RegistryPolicyProxy.PolicyKeys)
                {
                    using (var policyKey = regRoot.OpenSubKey(policyPath, false))
                    {
                        if (policyKey != null) addSubtree(policyPath, policyKey); // Added null check
                    }
                }
                return pol;
            }
            else
            {
                throw new InvalidOperationException("Policy source type not supported");
            }
        }

        public DialogResult DisplayAdmxLoadErrorReport(IEnumerable<AdmxLoadFailure> Failures, bool AskContinue = false)
        {
            if (!Failures.Any())
                return DialogResult.OK;
            var boxButtons = AskContinue ? MessageBoxButtons.YesNo : MessageBoxButtons.OK;
            string header = "Errors were encountered while adding administrative templates to the workspace.";
            return MessageBox.Show(header + (AskContinue ? " Continue trying to use this workspace?" : "") + Environment.NewLine + Environment.NewLine + string.Join(Environment.NewLine + Environment.NewLine, Failures.Select(f => f.ToString())), "Policy Plus", boxButtons, MessageBoxIcon.Exclamation);
        }

        public string GetPreferredLanguageCode()
        {
            return Convert.ToString(Configuration.GetValue("LanguageCode", System.Globalization.CultureInfo.CurrentCulture.Name));
        }

        private void CategoriesTree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            // When the user selects a new category in the left pane
            CurrentCategory = (PolicyPlusCategory)e.Node.Tag;
            UpdateCategoryListing();
            ClearSelections();
            UpdatePolicyInfo();
        }

        private void ResizePolicyNameColumn(object sender, EventArgs e)
        {
            // Fit the policy name column to the window size
            if (IsHandleCreated)
                BeginInvoke(new Action(() => PoliciesList.Columns[0].Width = PoliciesList.ClientSize.Width - (PoliciesList.Columns[1].Width + PoliciesList.Columns[2].Width)));
        }

        private void PoliciesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            // When the user highlights an item in the right pane
            if (PoliciesList.SelectedItems.Count > 0)
            {
                var selObject = PoliciesList.SelectedItems[0].Tag;
                if (selObject is PolicyPlusPolicy)
                {
                    CurrentSetting = (PolicyPlusPolicy)selObject;
                    HighlightCategory = null;
                }
                else if (selObject is PolicyPlusCategory)
                {
                    HighlightCategory = (PolicyPlusCategory)selObject;
                    CurrentSetting = null;
                }
            }
            else
            {
                ClearSelections();
            }
            UpdatePolicyInfo();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void OpenADMXFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var openAdmxFolderForm = new OpenAdmxFolder())
            {
                if (openAdmxFolderForm.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (openAdmxFolderForm.ClearWorkspace)
                            ClearAdmxWorkspace();
                        DisplayAdmxLoadErrorReport(AdmxWorkspace.LoadFolder(openAdmxFolderForm.SelectedFolder, GetPreferredLanguageCode()));
                        if (openAdmxFolderForm.ClearWorkspace)
                            Configuration.SetValue("AdmxSource", openAdmxFolderForm.SelectedFolder);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("The folder could not be fully added to the workspace. " + ex.Message, "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    PopulateAdmxUi();
                }
            }
        }

        private void OpenADMXFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Policy definitions files|*.admx";
                ofd.Title = "Open ADMX file";
                if (ofd.ShowDialog() != DialogResult.OK)
                    return;
                try
                {
                    DisplayAdmxLoadErrorReport(AdmxWorkspace.LoadFile(ofd.FileName, GetPreferredLanguageCode()));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("The ADMX file could not be added to the workspace. " + ex.Message, "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                PopulateAdmxUi();
            }
        }

        private void CloseADMXWorkspaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Close all policy definitions and clear the workspace
            ClearAdmxWorkspace();
            PopulateAdmxUi();
        }

        private void EmptyCategoriesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Toggle whether empty categories are visible
            ViewEmptyCategories = !ViewEmptyCategories;
            EmptyCategoriesToolStripMenuItem.Checked = ViewEmptyCategories;
            MoveToVisibleCategoryAndReload();
        }

        private void ComboAppliesTo_SelectedIndexChanged(object sender, EventArgs e)
        {
            // When the user chooses a different section to work with
            switch (ComboAppliesTo.Text ?? "")
            {
                case "User":
                    {
                        ViewPolicyTypes = AdmxPolicySection.User;
                        break;
                    }
                case "Computer":
                    {
                        ViewPolicyTypes = AdmxPolicySection.Machine;
                        break;
                    }

                default:
                    {
                        ViewPolicyTypes = AdmxPolicySection.Both;
                        break;
                    }
            }
            MoveToVisibleCategoryAndReload();
        }

        private void PoliciesList_DoubleClick(object sender, EventArgs e)
        {
            // When the user opens a policy object in the right pane
            if (PoliciesList.SelectedItems.Count == 0)
                return;
            var policyItem = PoliciesList.SelectedItems[0].Tag;
            if (policyItem is PolicyPlusCategory)
            {
                CurrentCategory = (PolicyPlusCategory)policyItem;
                UpdateCategoryListing();
            }
            else if (policyItem is PolicyPlusPolicy)
            {
                ShowSettingEditor((PolicyPlusPolicy)policyItem, ViewPolicyTypes);
            }
        }

        private void DeduplicatePoliciesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Make otherwise-identical pairs of user and computer policies into single dual-section policies
            ClearSelections();
            int deduped = PolicyProcessing.DeduplicatePolicies(AdmxWorkspace);
            MessageBox.Show("Deduplicated " + deduped + " policies.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Information);
            UpdateCategoryListing();
            UpdatePolicyInfo();
        }

        private void FindByIDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var findByIdForm = new FindById())
            {
                findByIdForm.AdmxWorkspace = AdmxWorkspace;
                if (findByIdForm.ShowDialog() == DialogResult.OK)
                {
                    var selCat = findByIdForm.SelectedCategory;
                    var selPol = findByIdForm.SelectedPolicy;
                    var selPro = findByIdForm.SelectedProduct;
                    var selSup = findByIdForm.SelectedSupport;
                    var selectedSection = findByIdForm.SelectedSection;

                    if (selCat is not null)
                    {
                        if (CategoryNodes.ContainsKey(selCat))
                        {
                            CurrentCategory = selCat;
                            UpdateCategoryListing();
                        }
                        else
                        {
                            MessageBox.Show("The category is not currently visible. Change the view settings and try again.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    else if (selPol is not null)
                    {
                        ShowSettingEditor(selPol, (AdmxPolicySection)Math.Min((int)ViewPolicyTypes, (int)selectedSection));
                        FocusPolicy(selPol);
                    }
                    else if (selPro is not null)
                    {
                        using (var detailProductForm = new DetailProduct()) detailProductForm.PresentDialog(selPro);
                    }
                    else if (selSup is not null)
                    {
                        using (var detailSupportForm = new DetailSupport()) detailSupportForm.PresentDialog(selSup);
                    }
                    else
                    {
                        MessageBox.Show("That object could not be found.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
        }

        private void OpenPolicyResourcesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!ConfirmSaveChanges())
                return;
            using (var openPolForm = new OpenPol())
            {
                if (UserPolicyLoader != null && CompPolicyLoader != null)
                {
                    openPolForm.SetLastSources(
                        CompPolicyLoader.Source,
                        CompPolicyLoader.LoaderData ?? "",
                        UserPolicyLoader.Source,
                        UserPolicyLoader.LoaderData ?? ""
                    );
                }

                if (openPolForm.ShowDialog() == DialogResult.OK)
                {
                    OpenPolicyLoaders(openPolForm.SelectedUser, openPolForm.SelectedComputer, false);
                    MoveToVisibleCategoryAndReload();
                }
            }
        }

        private void SavePoliciesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Save policy state and comments to disk
            void saveComments(Dictionary<string, string> Comments, PolicyLoader Loader)
            {
                try
                {
                    if (Comments is not null && Loader != null) // Added null check for Loader
                        CmtxFile.FromCommentTable(Comments).Save(Loader.GetCmtxPath());
                }
                catch (Exception ex)
                {
                    // Log or handle exception appropriately
                }
            };
            saveComments(UserComments, UserPolicyLoader);
            saveComments(CompComments, CompPolicyLoader);
            try
            {
                string compStatus = "not writable";
                string userStatus = "not writable";
                if (CompPolicyLoader != null && CompPolicyLoader.GetWritability() == PolicySourceWritability.Writable) // Added null check
                    compStatus = CompPolicyLoader.Save();
                if (UserPolicyLoader != null && UserPolicyLoader.GetWritability() == PolicySourceWritability.Writable) // Added null check
                    userStatus = UserPolicyLoader.Save();
                if (CompPolicyLoader != null)
                { // Added null check
                    Configuration.SetValue("CompSourceType", (int)CompPolicyLoader.Source);
                    Configuration.SetValue("CompSourceData", CompPolicyLoader.LoaderData ?? "");
                }
                if (UserPolicyLoader != null)
                { // Added null check
                    Configuration.SetValue("UserSourceType", (int)UserPolicyLoader.Source);
                    Configuration.SetValue("UserSourceData", UserPolicyLoader.LoaderData ?? "");
                }
                MessageBox.Show("Success." + Environment.NewLine + Environment.NewLine + "User policies: " + userStatus + "." + Environment.NewLine + Environment.NewLine + "Computer policies: " + compStatus + ".", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                PoliciesChanged = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Saving failed!" + Environment.NewLine + Environment.NewLine + ex.Message, "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string about = $"Policy Plus by Jordan White.{Environment.NewLine}Originally created by Ben Nordick.{Environment.NewLine}{Environment.NewLine}Creative Commons Attribution 4.0 International License.{Environment.NewLine}{Environment.NewLine}Available on GitHub: https://github.com/JordanJWhite/PolicyPlus";
            if (!string.IsNullOrEmpty(VersionData.Version.Trim()))
                about += $"{Environment.NewLine}{Environment.NewLine}Version: {VersionData.Version.Trim()}.";
            MessageBox.Show(about, "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ByTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var findByTextForm = new FindByText())
            {
                if (findByTextForm.PresentDialog(UserComments, CompComments) == DialogResult.OK)
                {
                    ShowSearchDialog(findByTextForm.Searcher);
                }
            }
        }

        private void SearchResultsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Show the search results window but don't start a search
            ShowSearchDialog(null);
        }

        private void FindNextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            do
            {
                var nextPol = GetFindResultsDialog().NextPolicy(); // Use the instance
                if (nextPol is null)
                {
                    MessageBox.Show("There are no more results that match the filter.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                }
                else if (ShouldShowPolicy(nextPol))
                {
                    FocusPolicy(nextPol);
                    break;
                }
            }
            while (true);
        }

        private void ByRegistryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var findByRegistryForm = new FindByRegistry())
            {
                if (findByRegistryForm.ShowDialog() == DialogResult.OK)
                    ShowSearchDialog(findByRegistryForm.Searcher);
            }
        }

        private void SettingInfoPanel_ClientSizeChanged(object sender, EventArgs e)
        {
            // Finagle the middle pane's UI elements
            SettingInfoPanel.AutoScrollMinSize = SettingInfoPanel.Size;
            PolicyTitleLabel.MaximumSize = new Size(PolicyInfoTable.Width, 0);
            PolicySupportedLabel.MaximumSize = new Size(PolicyInfoTable.Width, 0);
            PolicyDescLabel.MaximumSize = new Size(PolicyInfoTable.Width, 0);
            PolicyIsPrefLabel.MaximumSize = new Size(PolicyInfoTable.Width - 22, 0); // Leave room for the exclamation icon
            PolicyInfoTable.MaximumSize = new Size(SettingInfoPanel.Width - (SettingInfoPanel.VerticalScroll.Visible ? SystemInformation.VerticalScrollBarWidth : 0), 0);
            PolicyInfoTable.Width = PolicyInfoTable.MaximumSize.Width;
            if (PolicyInfoTable.ColumnCount > 0)
                PolicyInfoTable.ColumnStyles[0].Width = PolicyInfoTable.ClientSize.Width; // Only once everything is initialized
            PolicyInfoTable.PerformLayout(); // Force the table to take up its full desired size
            PInvoke.ShowScrollBar(SettingInfoPanel.Handle, 0, false); // 0 means horizontal
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!ConfirmSaveChanges())
                e.Cancel = true;

            // Dispose the shared FindResults dialog if it exists
            if (findResultsDialog != null && !findResultsDialog.IsDisposed)
            {
                findResultsDialog.Dispose();
            }
        }

        private void Main_Closed(object sender, EventArgs e)
        {
            ClosePolicySources(); // Make sure everything is cleaned up before quitting
        }

        private void PoliciesList_KeyDown(object sender, KeyEventArgs e)
        {
            // Activate a right pane item if the user presses Enter on it
            if (e.KeyCode == Keys.Enter & PoliciesList.SelectedItems.Count > 0)
                PoliciesList_DoubleClick(sender, e);
        }

        private void FilterOptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var filterOptionsForm = new FilterOptions())
            {
                if (filterOptionsForm.PresentDialog(CurrentFilter, AdmxWorkspace) == DialogResult.OK)
                {
                    CurrentFilter = filterOptionsForm.CurrentFilter;
                    ViewFilteredOnly = true;
                    OnlyFilteredObjectsToolStripMenuItem.Checked = true;
                    MoveToVisibleCategoryAndReload();
                }
            }
        }

        private void OnlyFilteredObjectsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Toggle whether the filter is used
            ViewFilteredOnly = !ViewFilteredOnly;
            OnlyFilteredObjectsToolStripMenuItem.Checked = ViewFilteredOnly;
            MoveToVisibleCategoryAndReload();
        }

        private void ImportSemanticPolicyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var importSpolForm = new ImportSpol())
            {
                if (importSpolForm.ShowDialog() == DialogResult.OK)
                {
                    var spol = importSpolForm.Spol;
                    int fails = spol.ApplyAll(AdmxWorkspace, UserPolicySource, CompPolicySource, UserComments, CompComments);
                    PoliciesChanged = true;
                    MoveToVisibleCategoryAndReload();
                    if (fails == 0)
                    {
                        MessageBox.Show("Semantic Policy successfully applied.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(fails + " out of " + spol.Policies.Count + " could not be applied.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
        }

        private void ImportPOLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "POL files|*.pol";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    PolFile pol = null;
                    try
                    {
                        pol = PolFile.Load(ofd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("The POL file could not be loaded. " + ex.Message, "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); // Added ex.Message
                        return;
                    }
                    using (var openSectionForm = new OpenSection())
                    {
                        if (openSectionForm.PresentDialog(true, true) == DialogResult.OK)
                        {
                            var section = openSectionForm.SelectedSection == AdmxPolicySection.User ? UserPolicySource : CompPolicySource;
                            pol.Apply(section);
                            PoliciesChanged = true;
                            MoveToVisibleCategoryAndReload();
                            MessageBox.Show("POL import successful.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
        }

        private void ExportPOLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "POL files|*.pol";
                using (var openSectionForm = new OpenSection())
                {
                    if (sfd.ShowDialog() == DialogResult.OK && openSectionForm.PresentDialog(true, true) == DialogResult.OK)
                    {
                        var section = openSectionForm.SelectedSection == AdmxPolicySection.Machine ? CompPolicySource : UserPolicySource;
                        try
                        {
                            GetOrCreatePolFromPolicySource(section).Save(sfd.FileName);
                            MessageBox.Show("POL exported successfully.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("The POL file could not be saved. " + ex.Message, "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); // Added ex.Message
                        }
                    }
                }
            }
        }

        private void AcquireADMXFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var downloadAdmxForm = new DownloadAdmx())
            {
                if (downloadAdmxForm.ShowDialog() == DialogResult.OK)
                {
                    if (!string.IsNullOrEmpty(downloadAdmxForm.NewPolicySourceFolder))
                    {
                        ClearAdmxWorkspace();
                        DisplayAdmxLoadErrorReport(AdmxWorkspace.LoadFolder(downloadAdmxForm.NewPolicySourceFolder, GetPreferredLanguageCode()));
                        Configuration.SetValue("AdmxSource", downloadAdmxForm.NewPolicySourceFolder);
                        PopulateAdmxUi();
                    }
                }
            }
        }

        private void LoadedADMXFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var loadedAdmxForm = new LoadedAdmx()) loadedAdmxForm.PresentDialog(AdmxWorkspace);
        }

        private void AllSupportDefinitionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var loadedSupportForm = new LoadedSupportDefinitions()) loadedSupportForm.PresentDialog(AdmxWorkspace);
        }

        private void AllProductsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var loadedProductsForm = new LoadedProducts()) loadedProductsForm.PresentDialog(AdmxWorkspace);
        }

        private void EditRawPOLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool userIsPol = UserPolicySource is PolFile;
            bool compIsPol = CompPolicySource is PolFile;
            if (!(userIsPol | compIsPol))
            {
                MessageBox.Show("Neither loaded source is backed by a POL file.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (Convert.ToInt32(Configuration.GetValue("EditPolDangerAcknowledged", 0)) == 0)
            {
                if (MessageBox.Show("Caution! This tool is for very advanced users. Improper modifications may result in inconsistencies in policies' states." + Environment.NewLine + Environment.NewLine + "Changes operate directly on the policy source, though they will not be committed to disk until you save. Are you sure you want to continue?", "Policy Plus", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                    return;
                Configuration.SetValue("EditPolDangerAcknowledged", 1);
            }
            using (var openSectionForm = new OpenSection())
            {
                if (openSectionForm.PresentDialog(userIsPol, compIsPol) == DialogResult.OK)
                {
                    using (var editPolForm = new EditPol())
                    {
                        editPolForm.PresentDialog(PolicyIcons, (PolFile)(openSectionForm.SelectedSection == AdmxPolicySection.Machine ? CompPolicySource : UserPolicySource), openSectionForm.SelectedSection == AdmxPolicySection.User);
                    }
                    PoliciesChanged = true;
                }
            }
            MoveToVisibleCategoryAndReload();
        }

        private void ExportREGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var openSectionForm = new OpenSection())
            {
                if (openSectionForm.PresentDialog(true, true) == DialogResult.OK)
                {
                    var source = openSectionForm.SelectedSection == AdmxPolicySection.Machine ? CompPolicySource : UserPolicySource;
                    using (var exportRegForm = new ExportReg())
                    {
                        exportRegForm.PresentDialog("", GetOrCreatePolFromPolicySource(source), openSectionForm.SelectedSection == AdmxPolicySection.User);
                    }
                }
            }
        }

        private void ImportREGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var openSectionForm = new OpenSection())
            {
                if (openSectionForm.PresentDialog(true, true) == DialogResult.OK)
                {
                    var source = openSectionForm.SelectedSection == AdmxPolicySection.Machine ? CompPolicySource : UserPolicySource;
                    using (var importRegForm = new ImportReg())
                    {
                        if (importRegForm.PresentDialog(source) == DialogResult.OK)
                        {
                            PoliciesChanged = true;
                            MoveToVisibleCategoryAndReload();
                        }
                    }
                }
            }
        }

        private void SetADMLLanguageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var langOptsForm = new LanguageOptions())
            {
                if (langOptsForm.PresentDialog(GetPreferredLanguageCode()) == DialogResult.OK)
                {
                    Configuration.SetValue("LanguageCode", langOptsForm.NewLanguage);
                    if (MessageBox.Show("Language changes will take effect when ADML files are next loaded. Would you like to reload the workspace now?", "Policy Plus", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        ClearAdmxWorkspace();
                        OpenLastAdmxSource();
                        PopulateAdmxUi();
                    }
                }
            }
        }

        private void RunGpupdateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var proc = Process.Start("gpupdate", "/force");
                proc.WaitForExit();
                if (proc.ExitCode == 0)
                {
                    MessageBox.Show("Group Policy update completed successfully.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Group Policy update finished with exit code {proc.ExitCode}.", "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not run Group Policy update: " + ex.Message, "Policy Plus", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PolicyObjectContext_Opening(object sender, CancelEventArgs e)
        {
            // When the right-click menu is opened
            bool showingForCategory;
            if (ReferenceEquals(PolicyObjectContext.SourceControl, CategoriesTree))
            {
                showingForCategory = true;
                PolicyObjectContext.Tag = CategoriesTree.SelectedNode.Tag;
            }
            else if (PoliciesList.SelectedItems.Count > 0) // Shown from the main view
            {
                var selEntryTag = PoliciesList.SelectedItems[0].Tag;
                showingForCategory = selEntryTag is PolicyPlusCategory;
                PolicyObjectContext.Tag = selEntryTag;
            }
            else
            {
                e.Cancel = true;
                return;
            }
            // Items are tagged in the designer for the objects they apply to
            foreach (var item in PolicyObjectContext.Items.OfType<ToolStripMenuItem>())
            {
                bool ok = true;
                if (Convert.ToString(item.Tag) == "P" & showingForCategory)
                    ok = false;
                if (Convert.ToString(item.Tag) == "C" & !showingForCategory)
                    ok = false;
                item.Visible = ok;
            }
        }

        private void PolicyObjectContext_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            // When the user clicks an item in the right-click menu
            var polObject = PolicyObjectContext.Tag; // The current policy object is in the Tag field
            if (ReferenceEquals(e.ClickedItem, CmeCatOpen))
            {
                CurrentCategory = (PolicyPlusCategory)polObject;
                UpdateCategoryListing();
            }
            else if (ReferenceEquals(e.ClickedItem, CmePolEdit))
            {
                ShowSettingEditor((PolicyPlusPolicy)polObject, ViewPolicyTypes);
            }
            else if (ReferenceEquals(e.ClickedItem, CmeAllDetails))
            {
                if (polObject is PolicyPlusCategory category)
                {
                    using (var detailCategoryForm = new DetailCategory()) detailCategoryForm.PresentDialog(category);
                }
                else if (polObject is PolicyPlusPolicy policy)
                {
                    using (var detailPolicyForm = new DetailPolicy()) detailPolicyForm.PresentDialog(policy);
                }
            }
            else if (ReferenceEquals(e.ClickedItem, CmePolInspectElements))
            {
                using (var inspectElementsForm = new InspectPolicyElements()) inspectElementsForm.PresentDialog((PolicyPlusPolicy)polObject, PolicyIcons, AdmxWorkspace);
            }
            else if (ReferenceEquals(e.ClickedItem, CmePolSpolFragment))
            {
                using (var inspectSpolForm = new InspectSpolFragment()) inspectSpolForm.PresentDialog((PolicyPlusPolicy)polObject, AdmxWorkspace, CompPolicySource, UserPolicySource, CompComments, UserComments);
            }
        }

        private void CategoriesTree_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            // Right-clicking doesn't actually select the node by default
            if (e.Button == MouseButtons.Right)
                CategoriesTree.SelectedNode = e.Node;
        }

        public static string PrettifyDescription(string Description)
        {
            if (string.IsNullOrEmpty(Description)) return string.Empty;
            var sb = new StringBuilder();
            foreach (var line in Description.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None)) // Fixed string split
                sb.AppendLine(line.Trim());
            return sb.ToString().TrimEnd();
        }
    }
}