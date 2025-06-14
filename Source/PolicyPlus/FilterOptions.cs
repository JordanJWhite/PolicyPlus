﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace PolicyPlus
{
    public partial class FilterOptions
    {
        public FilterConfiguration CurrentFilter;
        private Dictionary<PolicyPlusProduct, TreeNode> ProductNodes;

        public FilterOptions()
        {
            InitializeComponent();
        }
        public DialogResult PresentDialog(FilterConfiguration Configuration, AdmxBundle Workspace)
        {
            CurrentFilter = null;
            ProductNodes = new Dictionary<PolicyPlusProduct, TreeNode>();
            AllowedProductsTreeview.Nodes.Clear();
            
            // Use local function instead of lambda for recursive calls
            void AddProductsToNodeCollection(IEnumerable<PolicyPlusProduct> products, TreeNodeCollection nodes)
            {
                foreach (var product in products)
                {
                    var node = nodes.Add(product.DisplayName);
                    ProductNodes.Add(product, node);
                    node.Tag = product;
                    if (product.Children is not null)
                        AddProductsToNodeCollection(product.Children, node.Nodes);
                }
            }
            
            AddProductsToNodeCollection(Workspace.Products.Values, AllowedProductsTreeview.Nodes); // Recursively add products
            PrepareDialog(Configuration);
            return ShowDialog();
        }
        public void PrepareDialog(FilterConfiguration Configuration)
        {
            // Set the UI element state from the current filter
            if (Configuration.ManagedPolicy.HasValue)
            {
                PolicyTypeCombobox.SelectedIndex = Configuration.ManagedPolicy.Value ? 1 : 2;
            }
            else
            {
                PolicyTypeCombobox.SelectedIndex = 0;
            }
            if (Configuration.PolicyState.HasValue)
            {
                switch (Configuration.PolicyState.Value)
                {
                    case FilterPolicyState.NotConfigured:
                        {
                            PolicyStateCombobox.SelectedIndex = 1;
                            break;
                        }
                    case FilterPolicyState.Configured:
                        {
                            PolicyStateCombobox.SelectedIndex = 2;
                            break;
                        }
                    case FilterPolicyState.Enabled:
                        {
                            PolicyStateCombobox.SelectedIndex = 3;
                            break;
                        }
                    case FilterPolicyState.Disabled:
                        {
                            PolicyStateCombobox.SelectedIndex = 4;
                            break;
                        }
                }
            }
            else
            {
                PolicyStateCombobox.SelectedIndex = 0;
            }
            if (Configuration.Commented.HasValue)
            {
                CommentedCombobox.SelectedIndex = Configuration.Commented.Value ? 1 : 2;
            }
            else
            {
                CommentedCombobox.SelectedIndex = 0;
            }
            foreach (var node in ProductNodes.Values)
                node.Checked = false;
            if (Configuration.AllowedProducts is null)
            {
                SupportedCheckbox.Checked = false;
                AlwaysMatchAnyCheckbox.Checked = true;
                MatchBlankSupportCheckbox.Checked = true;
            }
            else
            {
                SupportedCheckbox.Checked = true;
                foreach (var product in Configuration.AllowedProducts)
                    ProductNodes[product].Checked = true;
                AlwaysMatchAnyCheckbox.Checked = Configuration.AlwaysMatchAny;
                MatchBlankSupportCheckbox.Checked = Configuration.MatchBlankSupport;
                
                // Use local function instead of lambda for recursive calls
                void ExpandIfNecessary(TreeNode node)
                {
                    foreach (TreeNode subnode in node.Nodes)
                    {
                        ExpandIfNecessary(subnode);
                        if (subnode.IsExpanded | subnode.Checked != node.Checked)
                        {
                            node.Expand();
                        }
                    }
                }
                
                foreach (TreeNode node in AllowedProductsTreeview.Nodes)
                    ExpandIfNecessary(node);
            }
        }
        private void SupportedCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            RequirementsBox.Enabled = SupportedCheckbox.Checked;
        }
        private void FilterOptions_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                DialogResult = DialogResult.Cancel;
        }
        private void ResetButton_Click(object sender, EventArgs e)
        {
            PrepareDialog(new FilterConfiguration());
        }
        private void OkButton_Click(object sender, EventArgs e)
        {
            // Create a filter configuration object from the user's settings
            var newConf = new FilterConfiguration();
            switch (PolicyTypeCombobox.SelectedIndex)
            {
                case 1:
                    {
                        newConf.ManagedPolicy = true;
                        break;
                    }
                case 2:
                    {
                        newConf.ManagedPolicy = false;
                        break;
                    }
            }
            switch (PolicyStateCombobox.SelectedIndex)
            {
                case 1:
                    {
                        newConf.PolicyState = FilterPolicyState.NotConfigured;
                        break;
                    }
                case 2:
                    {
                        newConf.PolicyState = FilterPolicyState.Configured;
                        break;
                    }
                case 3:
                    {
                        newConf.PolicyState = FilterPolicyState.Enabled;
                        break;
                    }
                case 4:
                    {
                        newConf.PolicyState = FilterPolicyState.Disabled;
                        break;
                    }
            }
            switch (CommentedCombobox.SelectedIndex)
            {
                case 1:
                    {
                        newConf.Commented = true;
                        break;
                    }
                case 2:
                    {
                        newConf.Commented = false;
                        break;
                    }
            }
            if (SupportedCheckbox.Checked)
            {
                newConf.AlwaysMatchAny = AlwaysMatchAnyCheckbox.Checked;
                newConf.MatchBlankSupport = MatchBlankSupportCheckbox.Checked;
                newConf.AllowedProducts = new List<PolicyPlusProduct>();
                foreach (var kv in ProductNodes)
                {
                    if (kv.Value.Checked)
                        newConf.AllowedProducts.Add(kv.Key);
                }
            }
            CurrentFilter = newConf;
            DialogResult = DialogResult.OK;
        }
        private void AllowedProductsTreeview_EnabledChanged(object sender, EventArgs e)
        {
            // Without this, the tree view looks really bad when disabled
            AllowedProductsTreeview.BackColor = AllowedProductsTreeview.Enabled ? SystemColors.Window : SystemColors.Control;
        }
        private void AllowedProductsTreeview_AfterCheck(object sender, TreeViewEventArgs e)
        {
            // Recursively set the check state on child products only in response to the user
            if (e.Action == TreeViewAction.Unknown)
                return;
            PropogateCheckStateDown(e.Node);
        }
        public void PropogateCheckStateDown(TreeNode Node)
        {
            foreach (TreeNode subnode in Node.Nodes)
            {
                subnode.Checked = Node.Checked;
                PropogateCheckStateDown(subnode);
            }
        }
    }

    public enum FilterPolicyState
    {
        Configured,
        NotConfigured,
        Enabled,
        Disabled
    }

    public class FilterConfiguration
    {
        public bool? ManagedPolicy;
        public FilterPolicyState? PolicyState;
        public bool? Commented;
        public List<PolicyPlusProduct> AllowedProducts;
        public bool AlwaysMatchAny;
        public bool MatchBlankSupport;
    }
    // The TreeView control has a bug: the displayed check state gets out of sync with the Checked property when the checkbox is double-clicked
    // Fix adapted from https://stackoverflow.com/a/3174824
    internal class DoubleClickIgnoringTreeView : TreeView
    {
        protected override void WndProc(ref Message m)
        {
            // Ignore WM_LBUTTONDBLCLK
            if (m.Msg != 0x203)
                base.WndProc(ref m);
        }
    }
}