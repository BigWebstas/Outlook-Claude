// The VSTO wiring that Visual Studio normally generates for an Outlook add-in.
namespace ClaudeMailSorter
{
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public sealed partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
    {
        internal Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes;
        internal Microsoft.Office.Tools.SmartTagCollection VstoSmartTags;
        internal Microsoft.Office.Interop.Outlook.Application Application;

        public ThisAddIn(global::Microsoft.Office.Tools.Outlook.Factory factory, global::System.IServiceProvider serviceProvider)
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            Globals.Factory = factory;
        }

        protected override void Initialize()
        {
            base.Initialize();
            this.Application = this.GetHostItem<Microsoft.Office.Interop.Outlook.Application>(
                typeof(Microsoft.Office.Interop.Outlook.Application), "Application");
            Globals.ThisAddIn = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeControls();
        }

        protected override void FinishInitialization()
        {
            this.InternalStartup();
            this.OnStartup();
        }

        protected override void InitializeDataBindings()
        {
            this.BeginInitialization();
            this.EndInitialization();
        }

        private void BeginInitialization()
        {
            this.BeginInit();
            this.CustomTaskPanes.BeginInit();
            this.VstoSmartTags.BeginInit();
        }

        private void EndInitialization()
        {
            this.VstoSmartTags.EndInit();
            this.CustomTaskPanes.EndInit();
            this.EndInit();
        }

        private void InitializeControls()
        {
            this.CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
            this.VstoSmartTags = Globals.Factory.CreateSmartTagCollection(null, null, "VstoSmartTags", "VstoSmartTags", this);
        }

        protected override void OnShutdown()
        {
            this.VstoSmartTags.Dispose();
            this.CustomTaskPanes.Dispose();
            base.OnShutdown();
        }
    }

    internal sealed partial class Globals
    {
        private Globals() { }

        private static ThisAddIn thisAddIn;
        private static Microsoft.Office.Tools.Outlook.Factory factory;

        internal static ThisAddIn ThisAddIn
        {
            get { return thisAddIn; }
            set
            {
                if (thisAddIn == null) thisAddIn = value;
                else throw new global::System.NotSupportedException();
            }
        }

        internal static Microsoft.Office.Tools.Outlook.Factory Factory
        {
            get { return factory; }
            set
            {
                if (factory == null) factory = value;
                else throw new global::System.NotSupportedException();
            }
        }
    }
}
