using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Shapes;
using Xarial.XCad.Base.Attributes;
using Xarial.XCad.Documents;
using Xarial.XCad.SolidWorks;
using Xarial.XCad.UI.Commands;
using Xarial.XCad.UI.Commands.Attributes;
using Xarial.XCad.UI.Commands.Enums;
using Xarial.XCad.UI.Commands.Structures;
using static SW_BOM_Add_In.BOMTable;


namespace SW_BOM_Add_In
{
    [ComVisible(true)]
    [Guid("A246164E-0C98-403D-8149-A8AA80DBCBA6")]
    [Title("BOM Table Add-in")]
    public class BOMTable : SwAddInEx
    {
        ICommandManager iCmdMgr = null;
        ISldWorks iSwApp = null;
        int addinID = 0;
        ModelDoc2 swModel = default(ModelDoc2);
        ModelDocExtension swModelDocExtension = default(ModelDocExtension);

        public enum CommandA_e {
            [CommandItemInfo(WorkspaceTypes_e.Assembly)]
            removeToolBars,
        }
        public enum CommandB_e
        {
            [CommandItemInfo(WorkspaceTypes_e.Assembly)]
            addToolBars,
        }
        private void OnCommandsAButtonClick(CommandA_e cmd)
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelDocExtension = swModel.Extension;
            Debug.WriteLine("sketch visibility =" + swModel.GetToolbarVisibility((int)swToolbar_e.swSketchToolbar));
            
            

            Application.ShowMessageBox("disabling main bar");
            ////handle the button click
            //bool mainVisible = swModel.GetToolbarVisibility(1);
            swModel.SetToolbarVisibility(1, false);
            swApp.RunCommand((int)swCommands_e.swCommands_View_Showhide_Tb, "");//swCommands_Toolbar_Context, "");
            swApp.RunCommand((int)swCommands_e.swCommands_Sw_Animatorpane, ""); 
            //swModel.SetToolbarVisibility(6, false);
            //swModel.SetToolbarVisibility(7, false);
            //SendKeys.SendWait("{F12}");
            //bool treeVisible2 = swModel.FeatureManager.EnableFeatureTreeWindow;

            swModelDocExtension.HideFeatureManager(true);

            //swModel.FeatureManager.EnableFeatureTreeWindow = false;
            //swModel.FeatureManager.EnableFeatureTree = false;
            /*swModelView = (ModelView)swModel.ActiveView;
            swModelView.FrameState = (int)swWindowState_e.swWindowMinimized;
            */
            //swModelView = (ModelView)swModelView.GetNext();


        }
        private void OnCommandsBButtonClick(CommandB_e cmd)
        {
            Application.ShowMessageBox("Printing num Groups");
            //handle the button click 
            //IModelDoc2.getToolbarVisibility(1, false);
            swModelDocExtension.HideFeatureManager(false);
        }
        public override void OnConnect()
        {
            //ConnectToSW();
            //Application.ShowMessageBox("Hello world");
            this.CommandManager.AddCommandGroup<CommandA_e>().CommandClick += OnCommandsAButtonClick;
            this.CommandManager.AddCommandGroup<CommandB_e>().CommandClick += OnCommandsBButtonClick;
            
            swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
            addinID = this.AddInId;
            

        }

        /*public bool ConnectToSW(object ThisSW, int cookie)
        {
            Application.ShowMessageBox("Hello world");
            this.CommandManager.AddCommandGroup<CommandA_e>().CommandClick += OnCommandsAButtonClick;
            this.CommandManager.AddCommandGroup<CommandB_e>().CommandClick += OnCommandsBButtonClick;
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Set up callbacks
            iSwApp.SetAddinCallbackInfo2(0, this, addinID);

            #region Set up the CommandManager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            //AddCommandMgr();
            #endregion

            return true;
        }
        */
        public SldWorks swApp; //= Application.SldWorks;//null;
        private ModelView swModelView;

    }
}