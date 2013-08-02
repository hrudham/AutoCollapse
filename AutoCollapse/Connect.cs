using EnvDTE;
using EnvDTE80;
using Extensibility;
using System;
using System.Collections.Generic;
using System.Threading;

namespace AutoCollapse
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2
	{
        #region Fields

        private DTE2 _applicationObject;
        private AddIn _addInInstance;
        private WindowEvents _windowEvents;
        private SolutionEvents _solutionEvents;
        private HashSet<string> _targets = new HashSet<string>();

        #endregion

        #region Constructor

        /// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
        public Connect()
        {
        }

        #endregion

        #region Methods

        /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
        /// <param term='application'>Root object of the host application.</param>
        /// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
        /// <param term='addInInst'>Object representing this Add-in.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (DTE2)application;
            _addInInstance = (AddIn)addInInst;

            _windowEvents = _applicationObject.Events.WindowEvents;
            _windowEvents.WindowCreated += OnWindowCreated;

            _solutionEvents = _applicationObject.Events.SolutionEvents;
            _solutionEvents.Opened += OnSolutionOpened;
        }

        /// <summary>
        /// Called when [disconnection].
        /// </summary>
        /// <param name="disconnectMode">The disconnect mode.</param>
        /// <param name="custom">The custom.</param>
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
        }

        /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />		
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>
        /// Called when window created.
        /// </summary>
        /// <param name="window">The window.</param>
        private void OnWindowCreated(Window window)
        {
            // Only collapse outlines if we are not in debug mode, and the
            // window that opened this window was the solution explorer.
            if (this._applicationObject.Debugger.CurrentMode != dbgDebugMode.dbgBreakMode &&
                this._applicationObject.ActiveWindow.ObjectKind == EnvDTE.Constants.vsWindowKindSolutionExplorer)
            {
                CollapseToDefinitions(window);
            }
        }

        /// <summary>
        /// Called when [solution opened].
        /// </summary>
        private void OnSolutionOpened()
        {
            foreach (Window window in this._applicationObject.Windows)
            {
                CollapseToDefinitions(window);
            }
        }
        
        /// <summary>
        /// Collapses to definitions.
        /// </summary>
        private void CollapseToDefinitions(Window window)
        {
            if (window.ObjectKind == EnvDTE.Constants.vsDocumentKindText)
            {
                var document = window.Document as Document;

                if (document != null && document.Type == "Text")
                {
                    var textDocument = document.Object("") as TextDocument;

                    if (textDocument != null)
                    {
                        ThreadPool.QueueUserWorkItem(
                            new WaitCallback(
                                delegate(object obj)
                                {
                                    this.TryExecuteCommand("Edit.StopOutlining");
                                    this.TryExecuteCommand("Edit.StartAutomaticOutlining");
                                    this.TryExecuteCommand("Edit.CollapsetoDefinitions");
                                }));
                    }
                }
            }
        }

        /// <summary>
        /// Tries the execute command.
        /// </summary>
        /// <param name="command">The command.</param>
        private void TryExecuteCommand(string command)
        {
            try
            {
                this._applicationObject.ExecuteCommand(command);
            }
            catch
            {
            }
        }
                
        #endregion
	}
}