using System;
using System.Collections;
using System.Linq;
using System.Windows.Forms;

namespace PolicyPlus
{
    public partial class DetailAdmx
    {
        public DetailAdmx()
        {
            InitializeComponent();
        }
        public void PresentDialog(AdmxFile Admx, AdmxBundle Workspace)
        {
            TextPath.Text = Admx.SourceFile;
            TextNamespace.Text = Admx.AdmxNamespace;
            TextSupersededAdm.Text = Admx.SupersededAdm;

            // Changed to use a more generic approach with proper casting
            void fillListview(ListView Control, IEnumerable Collection, Func<object, string> IdSelector, Func<object, string> NameSelector)
            {
                Control.Items.Clear();
                foreach (var item in Collection)
                {
                    var lsvi = Control.Items.Add(IdSelector(item));
                    lsvi.Tag = item;
                    lsvi.SubItems.Add(NameSelector(item));
                }
                Control.Columns[1].Width = Control.ClientRectangle.Width - Control.Columns[0].Width - SystemInformation.VerticalScrollBarWidth;
            };

            // Fix type conversion issues by using object parameter and casting inside lambdas
            fillListview(LsvPolicies, 
                Workspace.Policies.Values.Where(p => ReferenceEquals(p.RawPolicy.DefinedIn, Admx)), 
                obj => ((PolicyPlusPolicy)obj).RawPolicy.ID,
                obj => ((PolicyPlusPolicy)obj).DisplayName);
            
            fillListview(LsvCategories, 
                Workspace.FlatCategories.Values.Where(c => ReferenceEquals(c.RawCategory.DefinedIn, Admx)), 
                obj => ((PolicyPlusCategory)obj).RawCategory.ID,
                obj => ((PolicyPlusCategory)obj).DisplayName);
            
            fillListview(LsvProducts, 
                Workspace.FlatProducts.Values.Where(p => ReferenceEquals(p.RawProduct.DefinedIn, Admx)), 
                obj => ((PolicyPlusProduct)obj).RawProduct.ID,
                obj => ((PolicyPlusProduct)obj).DisplayName);
            
            fillListview(LsvSupportDefinitions, 
                Workspace.SupportDefinitions.Values.Where(s => ReferenceEquals(s.RawSupport.DefinedIn, Admx)), 
                obj => ((PolicyPlusSupport)obj).RawSupport.ID,
                obj => ((PolicyPlusSupport)obj).DisplayName);
            
            ShowDialog();
        }
        private void LsvPolicies_DoubleClick(object sender, EventArgs e)
        {
            var detailPolicyForm = new DetailPolicy();
            detailPolicyForm.PresentDialog((PolicyPlusPolicy)LsvPolicies.SelectedItems[0].Tag);
        }
        private void LsvCategories_DoubleClick(object sender, EventArgs e)
        {
            var detailCategoryForm = new DetailCategory();
            detailCategoryForm.PresentDialog((PolicyPlusCategory)LsvCategories.SelectedItems[0].Tag);
        }
        private void LsvProducts_DoubleClick(object sender, EventArgs e)
        {
            var detailProductForm = new DetailProduct();
            detailProductForm.PresentDialog((PolicyPlusProduct)LsvProducts.SelectedItems[0].Tag);
        }
        private void LsvSupportDefinitions_DoubleClick(object sender, EventArgs e)
        {
            var detailSupportForm = new DetailSupport();
            detailSupportForm.PresentDialog((PolicyPlusSupport)LsvSupportDefinitions.SelectedItems[0].Tag);
        }
    }
}