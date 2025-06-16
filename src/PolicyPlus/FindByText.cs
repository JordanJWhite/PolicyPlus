using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PolicyPlus
{
    public partial class FindByText
    {
        private Dictionary<string, string>[] CommentSources;
        public Func<PolicyPlusPolicy, bool> Searcher;

        public FindByText()
        {
            InitializeComponent();
        }
        public DialogResult PresentDialog(params Dictionary<string, string>[] CommentDicts)
        {
            CommentSources = CommentDicts.Where(d => d is not null).ToArray();
            return ShowDialog();
        }
        private void FindByText_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                DialogResult = DialogResult.Cancel;
        }
        private void SearchButton_Click(object sender, EventArgs e)
        {
            string text = StringTextbox.Text;
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("Please enter search terms.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            bool checkTitle = TitleCheckbox.Checked;
            bool checkDesc = DescriptionCheckbox.Checked;
            bool checkComment = CommentCheckbox.Checked;
            if (!(checkTitle | checkDesc | checkComment))
            {
                MessageBox.Show("At least one attribute must be searched. Check one of the boxes and try again.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            Searcher = new Func<PolicyPlusPolicy, bool>((Policy) =>
                {
                    string cleanupStr(string RawText) => new string(RawText.Trim().ToLowerInvariant().Where(c => !".,'\";/!(){}[]".Contains(c)).ToArray());
                    // Parse the query string for wildcards or quoted strings
                    string[] rawSplitted = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    var simpleWords = new List<string>();
                    var wildcards = new List<string>();
                    var quotedStrings = new List<string>();
                    string partialQuotedString = "";
                    for (int n = 0; n < rawSplitted.Length; n++)
                    {
                        string curString = rawSplitted[n];
                        if (!string.IsNullOrEmpty(partialQuotedString))
                        {
                            partialQuotedString += curString + " ";
                            if (curString.EndsWith("\""))
                            {
                                quotedStrings.Add(cleanupStr(partialQuotedString));
                                partialQuotedString = "";
                            }
                        }
                        else if (curString.StartsWith("\""))
                        {
                            partialQuotedString = curString + " ";
                        }
                        else if (curString.Contains("*") | curString.Contains("?"))
                        {
                            wildcards.Add(cleanupStr(curString));
                        }
                        else
                        {
                            simpleWords.Add(cleanupStr(curString));
                        }
                    }
                    // Do the searching
                    bool isStringAHit(string SearchedText)
                    {
                        string cleanText = cleanupStr(SearchedText);
                        string[] wordsInText = cleanText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        
                        // Helper method to replace LikeOperator.LikeString
                        bool IsLike(string input, string pattern)
                        {
                            // Convert wildcard pattern to regex pattern
                            string regexPattern = "^" + Regex.Escape(pattern)
                                .Replace("\\*", ".*")
                                .Replace("\\?", ".") + "$";
                            return Regex.IsMatch(input, regexPattern);
                        }
                        
                        return simpleWords.All(w => wordsInText.Contains(w)) && 
                               wildcards.All(w => wordsInText.Any(wit => IsLike(wit, w))) && 
                               quotedStrings.All(w => cleanText.Contains(" " + w + " ") || 
                                                    cleanText.StartsWith(w + " ") || 
                                                    cleanText.EndsWith(" " + w) || 
                                                    cleanText == w);
                    };
                    
                    if (checkTitle)
                    {
                        if (isStringAHit(Policy.DisplayName))
                            return true;
                    }
                    if (checkDesc)
                    {
                        if (isStringAHit(Policy.DisplayExplanation))
                            return true;
                    }
                    if (checkComment)
                    {
                        if (CommentSources.Any((Source) => Source.ContainsKey(Policy.UniqueID) && isStringAHit(Source[Policy.UniqueID])))
                            return true;
                    }
                    return false;
                });
            DialogResult = DialogResult.OK;
        }
    }
}