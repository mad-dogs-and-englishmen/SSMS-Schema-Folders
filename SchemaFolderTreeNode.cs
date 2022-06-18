namespace SsmsSchemaFolders
{
    using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
    using System;
    using System.Drawing;

    internal class SchemaFolderTreeNode : HierarchyTreeNode, INodeWithMenu, IServiceProvider

    {
        private readonly object parent;

        public SchemaFolderTreeNode(object o) => parent = o;

        public override Icon Icon => (parent as INodeWithIcon)?.Icon;

        public override Icon SelectedIcon => (parent as INodeWithIcon)?.SelectedIcon;

        public override int State => (parent as INodeWithIcon)?.State ?? 0;

        public override bool ShowPolicyHealthState
        {
            get => false;
            set
            {
            }
        }

        #region INodeWithMenu Members

        public void DoDefaultAction()
        {
            (parent as INodeWithMenu)?.DoDefaultAction();
        }

        public void ShowContextMenu(Point screenPos)
        {
            (parent as INodeWithMenu)?.ShowContextMenu(screenPos);
        }

        #endregion

        #region IServiceProvider Members

        public object GetService(Type serviceType) => (parent as IServiceProvider)?.GetService(serviceType);

        #endregion
    }
}