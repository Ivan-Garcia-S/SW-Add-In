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

            //Create new configuration to be modified
            swModel.AddConfiguration3("Floating config #" + numConfigs, "", "", 65);
            
            //Get and store all components in the current assembly
            Feature swFeat = default(Feature);
            Object objArray = swAssembly.GetComponents(true);
            IEnumerable<object> components = (IEnumerable<object>)objArray;
            int compCount = swAssembly.GetComponentCount(true);
            
            //Create select data for selecting components
            selData = swSelMgr.CreateSelectData();

            //Select all components and make them "floating"
            swAssembly.SelectComponentsBySize(100.0);
            swAssembly.UnfixComponent();

            //Get all mates of selected components and suppress each one
            swAssembly.SelectComponentsBySize(100.0);
            int numComps = swSelMgr.GetSelectedObjectCount2(-1);
            for(int i = 1; i <= numComps; i++)
            {
                Component2 swComp2 = (Component2)swSelMgr.GetSelectedObject6(1, -1);
                Object mateArrayObj = swComp2.GetMates();
                IEnumerable<object> mates = (IEnumerable<object>)mateArrayObj;

                //Put all mates in array
                if (mates is not null)
                { 
                    Mate2[] mateArray = new Mate2[mates.Count()];
                    int cnt = 0;
                    foreach (Mate2 mate in mates)
                    {
                        mateArray[cnt++] = mate;
                    }

                    //Suspend the current selection list to revisit
                    swSelMgr.SuspendSelectionList();
   
                    for (int j = 1; j <= mates.Count(); j++)
                    {
                        bool added = swSelMgr.AddSelectionListObject(mateArray[j-1], selData);
                        swFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                        //Update supressed property to true
                        bool status = swModel.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, true, true, swFeat.Name);
                    }
                    //Resume the old selection list
                    swSelMgr.ResumeSelectionList2(false);

                }
                swSelMgr.DeSelect2(1, -1);
            }
            //Open the assembly for editing again
            swAssembly.EditAssembly();   

        }    
        public override void OnConnect()
        {
            //Create add-in buttons to activate funcitonality in the SW toolbar and CommandManager tab
            this.CommandManager.AddCommandGroup<FORGEX>().CommandClick += OnButtonClick;

            //Setting SW object and ID# variables
            swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
            addinID = this.AddInId;
            
        }
        public SldWorks swApp;
    }
}