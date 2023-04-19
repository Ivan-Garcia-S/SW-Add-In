using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Shapes;
using Xarial.XCad.Base.Attributes;
using Xarial.XCad.Documents;
using Xarial.XCad.SolidWorks;
using Xarial.XCad.UI.Commands;
using Xarial.XCad.UI.Commands.Attributes;
using Xarial.XCad.UI.Commands.Enums;
using Xarial.XCad.UI.Commands.Structures;
using static SW_BOM_Add_In.BOMTable;
using static System.Windows.Forms.AxHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace SW_BOM_Add_In
{
    [ComVisible(true)]
    [Guid("A246164E-0C98-403D-8149-A8AA80DBCBA6")]
    [Title("BOM Table Add-in")]
    public class BOMTable : SwAddInEx
    {
        //ICommandManager iCmdMgr = null;
        ISldWorks iSwApp = null;
        int addinID = 0;
        ModelDoc2 swModel = default(ModelDoc2);
        ModelDocExtension swModelDocExt = default(ModelDocExtension);
        AssemblyDoc swAssembly = default(AssemblyDoc);
        MateLoadReference swMateLoadRef = default(MateLoadReference);

        public enum Toolbar_Toggler {
            [CommandItemInfo(true,false, WorkspaceTypes_e.Assembly,false)]
            removeToolBars,
            
            [CommandItemInfo(true, false, WorkspaceTypes_e.Assembly, false)]
            addToolBars,
        }
        public enum New_Configuration
        {
            [CommandItemInfo(WorkspaceTypes_e.Assembly)]
            createNewConfig,
        }
        private void OnTogglerButtonClick(Toolbar_Toggler cmd)
        {
            //Setting SW model variables
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelDocExt = swModel.Extension;

            //Remove toolbars
            if (cmd == 0)
            { 
                Application.ShowMessageBox("disabling toolbars");

                //Hide/toggle all toolbars and workspace covering menus
                //swModel.SetToolbarVisibility(1, false);
                swApp.RunCommand((int)swCommands_e.swCommands_View_Showhide_Tb, "");
                swApp.RunCommand((int)swCommands_e.swCommands_Sw_Animatorpane, "");
                swModelDocExt.HideFeatureManager(true);
            }
            else if((int)cmd == 1)
            {
                Application.ShowMessageBox("enabling toolbars");
                //Show/toggle all toolbars and workspace covering menus
                swApp.RunCommand((int)swCommands_e.swCommands_View_Showhide_Tb, "");
                swApp.RunCommand((int)swCommands_e.swCommands_Sw_Animatorpane, "");
                swModelDocExt.HideFeatureManager(false);
            }


            //bool mainVisible = swModel.GetToolbarVisibility(1);
            //swModel.FeatureManager.EnableFeatureTreeWindow = false;
            //swModel.FeatureManager.EnableFeatureTree = false;
            /*swModelView = (ModelView)swModel.ActiveView;
            swModelView.FrameState = (int)swWindowState_e.swWindowMinimized;
            */
            //swModelView = (ModelView)swModelView.GetNext();


        }
        private void OnConfigButtonClick(New_Configuration cmd)
        {
            //Setting SW model variables
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelDocExt = swModel.Extension;
            swAssembly = (AssemblyDoc)swModel;

            SelectData selData;
            ConfigurationManager swConfigMgr = swModel.ConfigurationManager;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            //Lock main config
            Configuration originalConfig = swConfigMgr.ActiveConfiguration;
            swConfigMgr.ActiveConfiguration.Lock = true;
            Configuration newConfig = swConfigMgr.AddConfiguration2("Floating config", null, null, 1, null, "Present config", false);

            Debug.WriteLine("COunt = " + swAssembly.GetComponentCount(true).ToString());

            //Object[] parameters = new Object[newConfig.GetParameterCount()];
            //Object[] values = new Object[newConfig.GetParameterCount()];
            //Object parameters; //= new Object();
            // Object values;
            //newConfig.GetParameters(out parameters,out values);
            // Debug.WriteLine(parameters);
            Feature swFeat = default(Feature);
            Object objArray = swAssembly.GetComponents(true);
            IEnumerable<object> components = (IEnumerable<object>)objArray;
            int compCount = swAssembly.GetComponentCount(true);
            string[] names = new string[compCount];
            int ind = 0;
            selData = swSelMgr.CreateSelectData();
            Component2[] compArray = new Component2[compCount];

            Feature[] selObjs = new Feature[compCount];
            swFeat = (Feature)swModel.FirstFeature();
            string typeName = "";
            string typeName2 = "";
            string name = "";

            for (int i = 0; i <= 50; i++){
                if (swFeat != null){
                    typeName = swFeat.Name;
                    name = swFeat.GetNameForSelection(out typeName2);
                    Debug.WriteLine("Name of feat :" + name);
                    Debug.Write("Type1 is : " + swFeat.GetTypeName2() + swFeat.GetTypeName() + ".....");
                    typeName = typeName.ToUpper();
                    Debug.WriteLine("Type2 is: " + typeName);

                    if ("COMPONENT" == typeName){

                        selObjs[i] = swFeat;
                        Debug.WriteLine("component added:" + swFeat.Name);
                    }

                    swFeat = (Feature)swFeat.GetNextFeature();
                }

            }
            Debug.WriteLine("Finished loop");






            

            foreach(object comp in components)
            {
                Component2 compObj = (Component2)comp;
                names[ind] = compObj.Name2;
                compArray[ind++] = compObj;
            }

            for (int i = 0; i < swAssembly.GetComponentCount(true); i++)
            {
                Debug.WriteLine("Loop once");
                string featName = names[i];
                Debug.WriteLine("Featname = " + featName);
            
            /*
            //swFeat = (Feature)swSelMgr.GetSelectedObject6(i, -1);
            //swFeat.GetNameForSelection(out featType);
            bool selected = swModelDocExt.SelectByID2(featName, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            //swAssembly.UnfixComponent();
            Debug.WriteLine("selected = " + selected);
            */


                bool selected = swSelMgr.AddSelectionListObject((Object)compArray[0], null);
                Debug.WriteLine("selected = " + selected);

                swFeat = (Feature)swSelMgr.GetSelectedObject6(1, 0);
                Component2 swComp = (Component2)swFeat.GetSpecificFeature2();

                Object mateArray = swComp.GetMates();
                IEnumerable<object> mates = (IEnumerable<object>)mateArray;
                foreach(Mate2 mate in mates)
                {
                    swMateLoadRef = swAssembly.InsertLoadReference(mate);
                    swModelDocExt.SelectByID2(swMateLoadRef.Name, "MATE", 0, 0, 0, false, 0, null, 0);
                    bool retVal = swModel.EditSuppress2();
                    Debug.Assert(retVal);
                    
                }
                swAssembly.EditAssembly();


                /*if ( featType == "Mate"){
                    swModel.EditSuppress2();
                    Debug.WriteLine("Suppressed Mate");
                }
                else if(featType == "Part")
                {
                    swAssembly.UnfixComponent();
                    Debug.WriteLine("Unfixed Part");
                }
                else
                {
                    Debug.WriteLine(featType);
                }
                */
            }

        }    
        public override void OnConnect()
        {
            //Create add-in buttons to activate funcitonality in the SW toolbar
            this.CommandManager.AddCommandGroup<Toolbar_Toggler>().CommandClick += OnTogglerButtonClick;
            this.CommandManager.AddCommandGroup<New_Configuration>().CommandClick += OnConfigButtonClick;
            
            //Setting SW object and ID# variables
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
        public SldWorks swApp;
        //Setting SW model variables
        //private ModelView swModelView;

    }
}