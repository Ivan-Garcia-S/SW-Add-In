using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Xarial.XCad.Base.Attributes;
using Xarial.XCad.SolidWorks;
using Xarial.XCad.SolidWorks.Services;
using Xarial.XCad.UI.Commands;
using Xarial.XCad.UI.Commands.Attributes;
using Xarial.XCad.UI.Commands.Enums;



namespace FORGEX
{
    [ComVisible(true)]
    [Guid("A246164E-0C98-403D-8149-A8AA80DBCBA6")]
    [Title("FORGEX AddIn")]
    public class SW_FORGEX : SwAddInEx
    {
        //Initialize SW variables
        ModelDoc2 swModel = default(ModelDoc2);
        ModelDocExtension swModelDocExt = default(ModelDocExtension);
        AssemblyDoc swAssembly = default(AssemblyDoc);
        SelectData selData = default(SelectData);
        int addinID;
        int configID = 1;

        public enum FORGEX
        {
            [Title("Create New Configuration")]
            [Description("Creates a copy of the default configuration with all components unfixed and all mates supressed")]
            [CommandItemInfo(true, true, WorkspaceTypes_e.Assembly, true, RibbonTabTextDisplay_e.TextBelow)]
            createNewConfig,

            [Title("Show Toolbars")]
            [Description("Enables all necessary toolbars for work")]
            [CommandItemInfo(true, true, WorkspaceTypes_e.Assembly, true, RibbonTabTextDisplay_e.TextBelow)]
            addToolBars,

            [Title("Hide Toolbars")]
            [Description("Disables all toolbars that block the workspace")]
            [CommandItemInfo(true, true, WorkspaceTypes_e.Assembly, true, RibbonTabTextDisplay_e.TextBelow)]
            removeToolBars,
        }

        private void OnButtonClick(FORGEX cmd)
        {
            switch (cmd)
            {
                case FORGEX.createNewConfig:
                    OnConfigButtonClick();
                    break;

                case FORGEX.addToolBars:
                    AddToolbars();
                    break;

                case FORGEX.removeToolBars:
                    RemoveToolbars();
                    break;
                default: break;
            } 
        }

        private void AddToolbars()
        {
            //Setting SW model variables
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelDocExt = swModel.Extension;
            IModelViewManager swModelViewMgr = swModel.ModelViewManager;

            if (!swModelViewMgr.FeatureManagerTabsVisible)
            {
                //Show/toggle all toolbars and workspace covering menus
                swApp.RunCommand((int)swCommands_e.swCommands_View_Showhide_Tb, "");
                swApp.RunCommand((int)swCommands_e.swCommands_Sw_Animatorpane, "");
                swModelDocExt.HideFeatureManager(false);
            }
        }

        private void RemoveToolbars()
        {
            //Setting SW model variables
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelDocExt = swModel.Extension;
            IModelViewManager swModelViewMgr = swModel.ModelViewManager;

            if (swModelViewMgr.FeatureManagerTabsVisible)
            {
                //Hide/toggle all toolbars and workspace covering menus
                swApp.RunCommand((int)swCommands_e.swCommands_View_Showhide_Tb, "");
                swApp.RunCommand((int)swCommands_e.swCommands_Sw_Animatorpane, "");
                swModelDocExt.HideFeatureManager(true);
            }
        }
        private void OnConfigButtonClick()
        {
            //Setting SW model variables
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelDocExt = swModel.Extension;
            swAssembly = (AssemblyDoc)swModel;
            ConfigurationManager swConfigMgr = swModel.ConfigurationManager;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            
            //Lock all configs so their components are not affected
            Configuration originalConfig = swConfigMgr.ActiveConfiguration;
            swConfigMgr.ActiveConfiguration.Lock = true;

            //Get number of configurations and their names
            int numConfigs = swModel.GetConfigurationCount();
            string[] configNames = (string[])swModel.GetConfigurationNames();
            Configuration defaultConfig = default(Configuration);

            //Locks the configurations
            foreach (string name in configNames)
            {
                Configuration config = (Configuration)swModel.GetConfigurationByName(name);
                if (name == configNames[0]) defaultConfig = config;
                config.Lock = true;
            }

            /*
            Component2 root = originalConfig.GetRootComponent3(false);
            Debug.WriteLine("Root name= " + root.Name);

            swModelDocExt.SelectByID2("Default", "CONFIGURATIONS", 0, 0, 0, false, 0, null, 0);//defaultConfig.Select2(false, selData);

            swModel.EditCopy();
            //swModel.ClearSelection2(true);
            //swApp.RunCommand((int)swCommands_e.swCommands_Copy, "");
            //swModelDocExt.SelectByID2("piezo top.SLDASM", "COMPONENT", 0, 0, 0, false, 0, null, 0);//root.Select4(false, selData, false);

            swModel.Paste();
            //swApp.RunCommand((int)swCommands_e.swCommands_Paste, "");
            swModel.ClearSelection2(true);

            swModel.ShowConfiguration2("Copy of " + defaultConfig.Name);
            //Configuration newConfig = swConfigMgr.AddConfiguration2("Floating config #" + numConfigs, null, null, 1, null, "Present config", true);
            //bool success = swModel.AddConfiguration2("Floating config #" + numConfigs, "", "", true, false, false, true, 256);

            //swConfigMgr.AddConfiguration2("Floating config #" + numConfigs, "", "", 65, "Default", "", false);

            */

            swModel.AddConfiguration3("Floating config #" + numConfigs, "", "", 65);
            Configuration activeConfig2 = swConfigMgr.ActiveConfiguration;
            Debug.WriteLine("New Active config = " + activeConfig2.Name);

            Feature swFeat = default(Feature);
            Object objArray = swAssembly.GetComponents(true);
            IEnumerable<object> components = (IEnumerable<object>)objArray;
            int compCount = swAssembly.GetComponentCount(true);
            string[] names = new string[compCount];
            int ind = 0;
            selData = swSelMgr.CreateSelectData();
            Component2[] compArray = new Component2[compCount];
            Feature[] selObjs = new Feature[compCount];
            
            /*MathTransform[] compTransforms = new MathTransform[compCount];
            swAssembly.SelectComponentsBySize(100.0);

            for (int i = 0; i < compCount; i++)
            {
                Component2 swComp2 = (Component2)swSelMgr.GetSelectedObject6(1, -1);
                compTransforms[i] = swComp2.Transform2;
                swSelMgr.DeSelect2(1, -1);
                Debug.WriteLine(swComp2.Transform2.ToString());
            }
            */
            //Select all components and make them "floating"
            swAssembly.SelectComponentsBySize(100.0);
            swAssembly.UnfixComponent();

            swAssembly.SelectComponentsBySize(100.0);
            int numComps = swSelMgr.GetSelectedObjectCount2(-1);
           // Debug.WriteLine("Number of selected components = " + numComps);
            for(int i = 1; i <= numComps; i++)
            {
               // Debug.WriteLine("in suppress loop");
                Component2 swComp2 = (Component2)swSelMgr.GetSelectedObject6(1, -1);
                //Debug.WriteLine("Component name = " + swComp2.Name2);
                Object mateArrayObj = swComp2.GetMates();
                
                IEnumerable<object> mates = (IEnumerable<object>)mateArrayObj;

                if (mates is not null)
                {
                    
                    Mate2[] mateArray = new Mate2[mates.Count()];
                    int cnt = 0;
                    foreach (Mate2 mate in mates)
                    {
                        mateArray[cnt++] = mate;
                    }

                   //Debug.WriteLine("Mates to suppress = " + mates.Count());
                    //Debug.WriteLine("Number of selected components before suspended list= " + swSelMgr.GetSelectedObjectCount2(-1));
                    swSelMgr.SuspendSelectionList();
   
                    for (int j = 1; j <= mates.Count(); j++)
                    {
                        bool added = swSelMgr.AddSelectionListObject(mateArray[j-1], selData);
                        //Debug.WriteLine("Number of selected components after add= " + swSelMgr.GetSelectedObjectCount2(-1));
                        //Debug.WriteLine("Added to list= " + added);
                        swFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                        //Debug.WriteLine("NAME OF MATE IS " + swFeat.Name);
                        bool status = swModel.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, true, true, swFeat.Name);

                       // Debug.WriteLine("suppressed = " + status);
                       // Debug.WriteLine("Number of selected components after suppressed= " + swSelMgr.GetSelectedObjectCount2(-1));

                    }
                    swSelMgr.ResumeSelectionList2(false);

                }
                swSelMgr.DeSelect2(1, -1);

            }
            swAssembly.EditAssembly();   

        }    
        public override void OnConnect()
        {
            //Create add-in buttons to activate funcitonality in the SW toolbar
            //this.CommandManager.AddCommandGroup<Toolbar_Toggler>().CommandClick += OnTogglerButtonClick;
            this.CommandManager.AddCommandGroup<FORGEX>().CommandClick += OnButtonClick;

            //Setting SW object and ID# variables
            swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
            addinID = this.AddInId;
            
        }
        public SldWorks swApp;
        //Setting SW model variables

    }
}