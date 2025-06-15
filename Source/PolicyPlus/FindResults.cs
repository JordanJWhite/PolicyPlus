﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PolicyPlus
{
    public partial class FindResults
    {
        private AdmxBundle AdmxWorkspace;
        private Func<PolicyPlusPolicy, bool> SearchFunc;
        private bool CancelingSearch = false;
        private bool CancelDueToFormClose = false;
        private bool SearchPending = false;
        private bool HasSearched;
        private int LastSelectedIndex;
        public PolicyPlusPolicy SelectedPolicy;

        public FindResults()
        {
            InitializeComponent();
        }
        public DialogResult PresentDialogStartSearch(AdmxBundle Workspace, Func<PolicyPlusPolicy, bool> Searcher)
        {
            // Start running a search defined by one of the Find By windows
            AdmxWorkspace = Workspace;
            SearchFunc = Searcher;
            ResultsListview.Items.Clear();
            SearchProgress.Maximum = Workspace.Policies.Count;
            SearchProgress.Value = 0;
            StopButton.Enabled = true;
            CancelingSearch = false;
            CancelDueToFormClose = false;
            ProgressLabel.Text = "Starting search";
            SearchPending = true;
            HasSearched = true;
            LastSelectedIndex = -1;
            return ShowDialog();
        }
        public DialogResult PresentDialog()
        {
            // Open the dialog normally, like from the main form
            if (!HasSearched)
            {
                MessageBox.Show("No search has been run yet, so there are no results to display.", "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return DialogResult.Cancel;
            }
            CancelingSearch = false;
            CancelDueToFormClose = false;
            SearchPending = false;
            return ShowDialog();
        }
        public void ClearSearch()
        {
            HasSearched = false;
            ResultsListview.Items.Clear();
        }
        public PolicyPlusPolicy NextPolicy()
        {
            if (LastSelectedIndex >= ResultsListview.Items.Count - 1 | !HasSearched)
                return null;
            LastSelectedIndex += 1;
            return (PolicyPlusPolicy)ResultsListview.Items[LastSelectedIndex].Tag;
        }
        public void SearchJob(AdmxBundle Workspace, Func<PolicyPlusPolicy, bool> Searcher)
        {
            // The long-running task that searches all the policies
            int searchedSoFar = 0;
            int results = 0;
            bool stoppedByCancel = false;
            var pendingInsertions = new List<PolicyPlusPolicy>();
            void addPendingInsertions()
            {
                ResultsListview.BeginUpdate();
                foreach (var insert in pendingInsertions)
                {
                    var lsvi = ResultsListview.Items.Add(insert.DisplayName);
                    lsvi.Tag = insert;
                    lsvi.SubItems.Add(insert.Category.DisplayName);
                }
                ResultsListview.EndUpdate();
                pendingInsertions.Clear();
            };
            foreach (var policy in Workspace.Policies)
            {
                bool isCanceled = System.Threading.Volatile.Read(ref CancelingSearch);
                if (isCanceled)
                {
                    stoppedByCancel = true;
                    break;
                }
                searchedSoFar += 1;
                bool isHit = Searcher(policy.Value); // The potentially expensive check
                if (isHit)
                {
                    results += 1;
                    pendingInsertions.Add(policy.Value);
                }
                if (searchedSoFar % 20 == 0) // UI updating is costly
                {
                    Invoke(new Action(() =>
                        {
                            addPendingInsertions();
                            SearchProgress.Value = searchedSoFar;
                            ProgressLabel.Text = "Searching: checked " + searchedSoFar + " policies so far, found " + results + " hits";
                        }));
                }
            }
            
            bool isFormClosing = System.Threading.Volatile.Read(ref CancelDueToFormClose);
            if (stoppedByCancel && isFormClosing)
                return; // Avoid accessing a disposed object
                
            Invoke(new Action(() =>
                {
                    addPendingInsertions();
                    string status = stoppedByCancel ? "Search canceled" : "Finished searching";
                    ProgressLabel.Text = status + ": checked " + searchedSoFar + " policies, found " + results + " hits";
                    SearchProgress.Value = SearchProgress.Maximum;
                    StopButton.Enabled = false;
                }));
        }
        public void StopSearch(bool Force)
        {
            System.Threading.Volatile.Write(ref CancelingSearch, true);
            System.Threading.Volatile.Write(ref CancelDueToFormClose, Force);
        }
        private void FindResults_Shown(object sender, EventArgs e)
        {
            if (SearchPending)
            {
                Task.Factory.StartNew(() => SearchJob(AdmxWorkspace, SearchFunc));
            }
            else if (LastSelectedIndex >= 0 & LastSelectedIndex < ResultsListview.Items.Count)
            {
                // Restore the last selection
                var lastSelected = ResultsListview.Items[LastSelectedIndex];
                lastSelected.Selected = true;
                lastSelected.Focused = true;
                lastSelected.EnsureVisible();
            }
        }
        private void StopButton_Click(object sender, EventArgs e)
        {
            StopSearch(false);
        }
        private void ResultsListview_SizeChanged(object sender, EventArgs e)
        {
            ChTitle.Width = ResultsListview.ClientSize.Width - ChCategory.Width;
        }
        private void FindResults_Closing(object sender, CancelEventArgs e)
        {
            StopSearch(true);
            if (SearchProgress.Value != SearchProgress.Maximum)
            {
                ProgressLabel.Text = "Search canceled";
                SearchProgress.Maximum = 100;
                SearchProgress.Value = SearchProgress.Maximum;
            }
        }
        private void GoClicked(object sender, EventArgs e)
        {
            if (ResultsListview.SelectedItems.Count == 0)
                return;
            SelectedPolicy = (PolicyPlusPolicy)ResultsListview.SelectedItems[0].Tag;
            LastSelectedIndex = ResultsListview.SelectedIndices[0]; // Remember which item is selected
            DialogResult = DialogResult.OK;
        }
        private void FindResults_Load(object sender, EventArgs e)
        {
            // Enable double-buffering for the results view
            var doubleBufferProp = typeof(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            doubleBufferProp.SetValue(ResultsListview, true);
        }
    }
}