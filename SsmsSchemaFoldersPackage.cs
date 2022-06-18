extern alias Ssms18;
extern alias Ssms19;
extern alias Ssms2012;
extern alias Ssms2014;
extern alias Ssms2016;
extern alias Ssms2017;

namespace SsmsSchemaFolders
{
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(PackageGuidString)]
    [ProvideAutoLoad("d114938f-591c-46cf-a785-500a82d97410")]
    [ProvideOptionPage(typeof(SchemaFolderOptions), "SQL Server Object Explorer", "Schema Folders", 114, 116, true)]
    public sealed class SsmsSchemaFoldersPackage : Package, IDebugOutput
    {
        public const string PackageGuidString = "a88a775f-7c86-4a09-b5a6-890c4c38261b";

        public const string SchemaFolderNodeTag = "SchemaFolder";
        public static readonly Guid PackageGuid = new Guid(PackageGuidString);

        private IObjectExplorerExtender _objectExplorerExtender;

#pragma warning disable CS0649
        private IVsOutputWindowPane _outputWindowPane;
#pragma warning restore CS0649

        public SchemaFolderOptions Options { get; set; }

        protected override void Initialize()
        {
            base.Initialize();

#if DEBUG
            var outputWindow = (IVsOutputWindow)GetService(typeof(SVsOutputWindow));
            var guidPackage = new Guid(PackageGuidString);
            outputWindow.CreatePane(guidPackage, "Schema Folders debug output", 1, 0);
            outputWindow.GetPane(ref guidPackage, out _outputWindowPane);
#endif

            (this as IVsPackage).GetAutomationObject("SQL Server Object Explorer.Schema Folders", out var obj);
            Options = (SchemaFolderOptions)obj;

            _objectExplorerExtender = GetObjectExplorerExtender();

            if (_objectExplorerExtender != null)
            {
                AttachTreeViewEvents();
            }

            DelayAddSkipLoadingReg();
        }

        private void ActivityLogEntry(__ACTIVITYLOG_ENTRYTYPE entryType, string message)
        {
            DebugMessage(message);

            if (!(GetService(typeof(SVsActivityLog)) is IVsActivityLog log))
            {
                return;
            }

            log.LogEntryGuid((uint)entryType, ToString(), message, PackageGuid);
        }

        private void AddSkipLoadingReg()
        {
            var myPackage = UserRegistryRoot.CreateSubKey(@"Packages\{" + PackageGuidString + "}");
            myPackage?.SetValue("SkipLoading", 1);
        }

        private void AttachTreeViewEvents()
        {
            var treeView = _objectExplorerExtender.GetObjectExplorerTreeView();

            if (treeView != null)
            {
                treeView.BeforeExpand += ObjectExplorerTreeViewBeforeExpandCallback;
                treeView.AfterExpand += ObjectExplorerTreeViewAfterExpandCallback;
            }
            else
            {
                ActivityLogEntry(__ACTIVITYLOG_ENTRYTYPE.ALE_ERROR, "Object Explorer TreeView == null");
            }
        }

        private void DelayAddSkipLoadingReg()
        {
            var delay = new Timer();

            delay.Tick += delegate
                          {
                              delay.Stop();
                              AddSkipLoadingReg();
                          };

            delay.Interval = 1000;
            delay.Start();
        }

        private IObjectExplorerExtender GetObjectExplorerExtender()
        {
            try
            {
                var ssmsInterfacesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SqlWorkbench.Interfaces.dll");

                if (File.Exists(ssmsInterfacesPath))
                {
                    var ssmsInterfacesVersion = FileVersionInfo.GetVersionInfo(ssmsInterfacesPath);

                    switch (ssmsInterfacesVersion.FileMajorPart)
                    {
                        case 16:
                            DebugMessage("SsmsVersion:19");

                            return new Ssms19::SsmsSchemaFolders.ObjectExplorerExtender(this, Options);
                        case 15:
                            DebugMessage("SsmsVersion:18");

                            return new Ssms18::SsmsSchemaFolders.ObjectExplorerExtender(this, Options);

                        case 14:
                            DebugMessage("SsmsVersion:2017");

                            return new Ssms2017::SsmsSchemaFolders.ObjectExplorerExtender(this, Options);

                        case 13:
                            DebugMessage("SsmsVersion:2016");

                            return new Ssms2016::SsmsSchemaFolders.ObjectExplorerExtender(this, Options);

                        case 12:
                            DebugMessage("SsmsVersion:2014");

                            return new Ssms2014::SsmsSchemaFolders.ObjectExplorerExtender(this, Options);

                        case 11:
                            DebugMessage("SsmsVersion:2012");

                            return new Ssms2012::SsmsSchemaFolders.ObjectExplorerExtender(this, Options);

                        default:
                            ActivityLogEntry(__ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION,
                                             $"SqlWorkbench.Interfaces.dll v{ssmsInterfacesVersion.FileMajorPart}:{ssmsInterfacesVersion.FileMinorPart}");

                            break;
                    }
                }

                ActivityLogEntry(__ACTIVITYLOG_ENTRYTYPE.ALE_WARNING, "Unknown SSMS Version. Defaulting to 2016.");

                return new Ssms2016::SsmsSchemaFolders.ObjectExplorerExtender(this, Options);
            }
            catch (Exception ex)
            {
                ActivityLogEntry(__ACTIVITYLOG_ENTRYTYPE.ALE_ERROR, ex.ToString());

                return null;
            }
        }

        private void ObjectExplorerTreeViewAfterExpandCallback(object sender, TreeViewEventArgs e)
        {
            DebugMessage("\nObjectExplorerTreeViewAfterExpandCallback");

            try
            {
                DebugMessage("Node.Count:{0}", e.Node.GetNodeCount(false));

                if (!Options.Enabled)
                {
                    return;
                }

                if (e.Node.TreeView.InvokeRequired)
                {
                    DebugMessage("TreeView.InvokeRequired");
                }

                if (!_objectExplorerExtender.GetNodeExpanding(e.Node))
                {
                    return;
                }

                DebugMessage("node.Expanding");

                var waitCount = 0;

                e.Node.TreeView.Cursor = Cursors.AppStarting;

                var nodeExpanding = new Timer();
                nodeExpanding.Interval = 10;

                void NodeExpandingEvent(object o, EventArgs e2)
                {
                    DebugMessage("nodeExpanding:{0}", waitCount);
                    waitCount++;

                    if (e.Node.TreeView.InvokeRequired)
                    {
                        DebugMessage("TreeView.InvokeRequired");
                    }

                    DebugMessage("Node.Count:{0}", e.Node.GetNodeCount(false));

                    if (_objectExplorerExtender.GetNodeExpanding(e.Node))
                    {
                        return;
                    }

                    // ReSharper disable AccessToDisposedClosure
                    nodeExpanding.Tick -= NodeExpandingEvent;
                    nodeExpanding.Stop();
                    nodeExpanding.Dispose();

                    // ReSharper restore AccessToDisposedClosure

                    ReorganizeFolders(e.Node, true);

                    e.Node.TreeView.Cursor = Cursors.Default;
                }

                nodeExpanding.Tick += NodeExpandingEvent;
                nodeExpanding.Start();
            }
            catch (Exception ex)
            {
                ActivityLogEntry(__ACTIVITYLOG_ENTRYTYPE.ALE_ERROR, ex.ToString());
            }
        }

        private void ObjectExplorerTreeViewBeforeExpandCallback(object sender, TreeViewCancelEventArgs e)
        {
            DebugMessage("\nObjectExplorerTreeViewBeforeExpandCallback");

            try
            {
                if (!Options.Enabled)
                {
                    return;
                }

                DebugMessage("Node.Count:{0}", e.Node.GetNodeCount(false));

                if (e.Node.GetNodeCount(false) == 1)
                {
                    return;
                }

                if (_objectExplorerExtender.GetNodeExpanding(e.Node))
                {
                    DebugMessage("node.Expanding");

                    ReorganizeFolders(e.Node);
                }
            }
            catch (Exception ex)
            {
                ActivityLogEntry(__ACTIVITYLOG_ENTRYTYPE.ALE_ERROR, ex.ToString());
            }
        }

        private void ReorganizeFolders(TreeNode node, bool expand = false)
        {
            DebugMessage("ReorganizeFolders");

            try
            {
                if (node?.Parent == null || (node.Tag != null && node.Tag.ToString() == SchemaFolderNodeTag))
                {
                    return;
                }

                var urnPath = _objectExplorerExtender.GetNodeUrnPath(node);

                if (string.IsNullOrEmpty(urnPath))
                {
                    return;
                }

                switch (urnPath)
                {
                    case "Server/Database/UserTablesFolder":
                    case "Server/Database/ViewsFolder":
                    case "Server/Database/SynonymsFolder":
                    case "Server/Database/StoredProceduresFolder":
                    case "Server/Database/Table-valuedFunctionsFolder":
                    case "Server/Database/Scalar-valuedFunctionsFolder":
                    case "Server/Database/SystemTablesFolder":
                    case "Server/Database/SystemViewsFolder":
                    case "Server/Database/SystemStoredProceduresFolder":
                        var schemaFolderCount = _objectExplorerExtender.ReorganizeNodes(node, SchemaFolderNodeTag);

                        if (expand && schemaFolderCount == 1)
                        {
                            node.LastNode.Expand();
                        }

                        break;

                    case "Server/Database/Table":
                    case "Server/Database/View":
                    case "Server/Database/Synonym":
                    case "Server/Database/StoredProcedure":
                    case "Server/Database/UserDefinedFunction":
                        if (Options.RenameNode)
                        {
                            DebugMessage(node.Text);
                            _objectExplorerExtender.RenameNode(node);
                        }

                        break;

                    default:
                        DebugMessage(urnPath);

                        break;
                }
            }
            catch (Exception ex)
            {
                ActivityLogEntry(__ACTIVITYLOG_ENTRYTYPE.ALE_ERROR, ex.ToString());
            }
        }

        #region IDebugOutput Members

        public void DebugMessage(string message)
        {
            if (_outputWindowPane == null)
            {
                return;
            }

            _outputWindowPane.OutputString(message);
            _outputWindowPane.OutputString("\r\n");
        }

        public void DebugMessage(string message, params object[] args)
        {
            if (_outputWindowPane == null)
            {
                return;
            }

            _outputWindowPane.OutputString(string.Format(message, args));
            _outputWindowPane.OutputString("\r\n");
        }

        #endregion
    }
}