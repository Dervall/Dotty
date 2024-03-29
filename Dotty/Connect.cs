using System;
using System.Windows;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Window = EnvDTE.Window;

namespace Dotty
{
    public class Connect : IDTExtensibility2, IDTCommandTarget
    {
        // Constants for command properties
        private const string CommandName = "ViewGraph";
        private const string CommandCaption = "View graph";
        private const string CommandTooltip = "View graph at cursor location";

        // Variables for IDE and add-in instances
        private DTE applicationObject;
        private AddIn addInInstance;

        // Buttons that will be created on built-in commandbars of Visual Studio
        // We must keep them at class level to remove them when the add-in is unloaded
        //     private CommandBarButton myStandardCommandBarButton;
        private CommandBarButton myCodeWindowCommandBarButton;

        // CommandBars that will be created by the add-in
        // We must keep them at class level to remove them when the add-in is unloaded
        private CommandBar myTemporaryToolbar;
        private Window myToolWindow;

        public void OnConnection(object application, ext_ConnectMode connectMode,
           object addInInst, ref Array custom)
        {
            try
            {
                applicationObject = (DTE)application;
                addInInstance = (AddIn)addInInst;

                switch (connectMode)
                {
                    case ext_ConnectMode.ext_cm_UISetup:

                        // Do nothing for this add-in with temporary user interface
                        break;

                    case ext_ConnectMode.ext_cm_Startup:

                        // The add-in was marked to load on startup
                        // Do nothing at this point because the IDE may not be fully initialized
                        // Visual Studio will call OnStartupComplete when fully initialized
                        break;

                    case ext_ConnectMode.ext_cm_AfterStartup:

                        // The add-in was loaded by hand after startup using the Add-In Manager
                        // Initialize it in the same way that when is loaded on startup
                        AddTemporaryUI();
                        break;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            AddTemporaryUI();
        }

        public void AddTemporaryUI()
        {
            const string vsCodeWindowCommandbarName = "Code Window";
            const string myTemporaryToolbarCaption = "Dotty";

            // The only command that will be created. We will create several buttons from it
            Command myCommand = null;

            // Built-in commandbars of Visual Studio

            // Buttons that will be created on a toolbars/commandbar popups created by the add-in
            // We don't need to keep them at class level to remove them when the add-in is unloaded 
            // because we will remove the whole toolbars/commandbar popups

            // The collection of Visual Studio commandbars

            var contextUIGuids = new object[] { };

            try
            {
                // ------------------------------------------------------------------------------------

                // Try to retrieve the command, just in case it was already created, ignoring the 
                // exception that would happen if the command was not created yet.
                try
                {
                    myCommand = applicationObject.Commands.Item(addInInstance.ProgID + "." + CommandName);
                }
                catch
                {
                }

                // Add the command if it does not exist
                if (myCommand == null)
                {
                    myCommand = applicationObject.Commands.AddNamedCommand(addInInstance,
                       CommandName, CommandCaption, CommandTooltip, true, 59, ref contextUIGuids,
                       (int)(vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled));
                }
                // ------------------------------------------------------------------------------------

                // Retrieve the collection of commandbars
                // Note:
                // - In VS.NET 2002/2003 (which uses the Office.dll reference) 
                //   DTE.CommandBars returns directly a CommandBars type, so a cast 
                //   to CommandBars is redundant
                // - In VS 2005 or higher (which uses the new Microsoft.VisualStudio.CommandBars.dll reference) 
                //   DTE.CommandBars returns an Object type, so we do need a cast to CommandBars
                var commandBars = (CommandBars)applicationObject.CommandBars;

                // Retrieve some built-in commandbars
                CommandBar codeCommandBar = commandBars[vsCodeWindowCommandbarName];


                // ------------------------------------------------------------------------------------
                // Button on the "Code Window" context menu
                // ------------------------------------------------------------------------------------

                // Add a button to the built-in "Code Window" context menu
                myCodeWindowCommandBarButton = (CommandBarButton)myCommand.AddControl(codeCommandBar,
                   codeCommandBar.Controls.Count + 1);

                // Change some button properties
                myCodeWindowCommandBarButton.Caption = CommandCaption;
                myCodeWindowCommandBarButton.BeginGroup = true; // Separator line above button

                // ------------------------------------------------------------------------------------
                // New toolbar
                // ------------------------------------------------------------------------------------

                // Add a new toolbar 
                myTemporaryToolbar = commandBars.Add(myTemporaryToolbarCaption,
                   MsoBarPosition.msoBarTop, Type.Missing, true);

                // Add a new button on that toolbar
                var myToolBarButton = (CommandBarButton)myCommand.AddControl(myTemporaryToolbar,
                                                                                          myTemporaryToolbar.Controls.Count + 1);

                // Change some button properties
                myToolBarButton.Caption = CommandCaption;
                myToolBarButton.Style = MsoButtonStyle.msoButtonIconAndCaption; // It could be also msoButtonIcon

                // Make visible the toolbar
                myTemporaryToolbar.Visible = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                switch (removeMode)
                {
                    case ext_DisconnectMode.ext_dm_HostShutdown:
                    case ext_DisconnectMode.ext_dm_UserClosed:
                        if ((myCodeWindowCommandBarButton != null))
                        {
                            myCodeWindowCommandBarButton.Delete(true);
                        }

                        if ((myTemporaryToolbar != null))
                        {
                            myTemporaryToolbar.Delete();
                        }
                        break;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void Exec(string cmdName, vsCommandExecOption executeOption, ref object varIn,
           ref object varOut, ref bool handled)
        {
            if ((executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault))
            {
                if (cmdName == addInInstance.ProgID + "." + CommandName)
                {
                    ShowToolWindow();
                    handled = true;
                }
            }
        }

        public void QueryStatus(string cmdName, vsCommandStatusTextWanted neededText,
           ref vsCommandStatus statusOption, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (cmdName == addInInstance.ProgID + "." + CommandName)
                {
                    statusOption = vsCommandStatus.vsCommandStatusEnabled |
                                   vsCommandStatus.vsCommandStatusSupported;
                }
                else
                {
                    statusOption = vsCommandStatus.vsCommandStatusUnsupported;
                }
            }
        }

        private void ShowToolWindow()
        {
            const string toolwindowGuid = "{54E3B802-AB4E-4EBF-999B-4DD19E980608}";

            try
            {
                if (myToolWindow == null) // First time, create it
                {
                    var windows2 = (Windows2)applicationObject.Windows;

                    string assembly = System.Reflection.Assembly.GetExecutingAssembly().Location;
                    object userControl = null;

                    myToolWindow = windows2.CreateToolWindow2(addInInstance, assembly,
                       typeof(ViewGraphControl).FullName, "Graph view", toolwindowGuid, ref userControl);
                    // TODO: Really ugly solution here
                    ViewGraphControl.Initialize(myToolWindow, applicationObject);
                }
                myToolWindow.Visible = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
    }
}
