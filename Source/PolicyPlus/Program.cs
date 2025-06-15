using System;
using System.Windows.Forms;

namespace PolicyPlus
{
    static class Program
    {
        private static Main mainForm;
        private static DetailPolicy detailPolicyForm;
        private static DetailProduct detailProductForm;
        private static DetailAdmx detailAdmxForm;
        private static OpenUserGpo openUserGpoForm;
        private static OpenUserRegistry openUserRegistryForm;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            mainForm = new Main();
            Application.Run(mainForm);
        }

        public static Main GetMainForm()
        {
            return mainForm;
        }

        public static DetailPolicy GetDetailPolicyForm()
        {
            if (detailPolicyForm == null || detailPolicyForm.IsDisposed)
                detailPolicyForm = new DetailPolicy();
            return detailPolicyForm;
        }

        public static DetailProduct GetDetailProductForm()
        {
            if (detailProductForm == null || detailProductForm.IsDisposed)
                detailProductForm = new DetailProduct();
            return detailProductForm;
        }

        public static DetailAdmx GetDetailAdmxForm()
        {
            if (detailAdmxForm == null || detailAdmxForm.IsDisposed)
                detailAdmxForm = new DetailAdmx();
            return detailAdmxForm;
        }

        public static OpenUserGpo GetOpenUserGpoForm()
        {
            if (openUserGpoForm == null || openUserGpoForm.IsDisposed)
                openUserGpoForm = new OpenUserGpo();
            return openUserGpoForm;
        }

        public static OpenUserRegistry GetOpenUserRegistryForm()
        {
            if (openUserRegistryForm == null || openUserRegistryForm.IsDisposed)
                openUserRegistryForm = new OpenUserRegistry();
            return openUserRegistryForm;
        }
    }
}