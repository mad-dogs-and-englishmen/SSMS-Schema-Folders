namespace SsmsSchemaFolders
{
    using JetBrains.Annotations;
    using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Windows.Forms;

    public class ObjectExplorerExtender : IObjectExplorerExtender
    {
        public ObjectExplorerExtender(IServiceProvider package, ISchemaFolderOptions options)
        {
            Package = package;
            Options = options;
        }

        private ISchemaFolderOptions Options { get; }
        private IServiceProvider Package { get; }

        private static string FormatNodeName([NotNull] INodeInformation nodeInformation) =>
            nodeInformation.InvariantName.EndsWith("." + nodeInformation.Name)
                ? nodeInformation.InvariantName.Replace("." + nodeInformation.Name, string.Empty)
                : null;

        private static INodeInformation GetNodeInformation(TreeNode node)
        {
            INodeInformation result = null;

            if (node is IServiceProvider serviceProvider)
            {
                result = serviceProvider.GetService(typeof(INodeInformation)) as INodeInformation;
            }

            return result;
        }

        private static string GetNodeSchema(TreeNode node)
        {
            var ni = GetNodeInformation(node);

            return ni == null ? null : FormatNodeName(ni);
        }

        private void DebugMessage(string message)
        {
            if (Package is IDebugOutput output)
            {
                output.DebugMessage(message);
            }
        }

        #region IObjectExplorerExtender Members

        public bool GetNodeExpanding(TreeNode node) => node is ILazyLoadingNode lazyNode && lazyNode.Expanding;

        public string GetNodeUrnPath(TreeNode node)
        {
            var ni = GetNodeInformation(node);

            return ni?.UrnPath;
        }

        public TreeView GetObjectExplorerTreeView()
        {
            var objectExplorerService = (IObjectExplorerService)Package.GetService(typeof(IObjectExplorerService));

            if (objectExplorerService == null)
            {
                return null;
            }

            var oesTreeProperty = objectExplorerService.GetType()
                                                       .GetProperty("Tree", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.IgnoreCase);

            return oesTreeProperty != null ? (TreeView)oesTreeProperty.GetValue(objectExplorerService, null) : null;
        }

        public void RenameNode(TreeNode node)
        {
            node.Text = node.Text.Substring(node.Text.IndexOf('.') + 1);
        }

        public int ReorganizeNodes(TreeNode node, string nodeTag)
        {
            DebugMessage("ReorganizeNodes");

            if (node.Nodes.Count <= 1)
            {
                return 0;
            }

            node.TreeView.BeginUpdate();

            var schemas = new Dictionary<string, List<TreeNode>>();

            foreach (TreeNode childNode in node.Nodes)
            {
                if (childNode.Tag != null && childNode.Tag.ToString() == nodeTag)
                {
                    if (!schemas.ContainsKey(childNode.Name))
                    {
                        schemas.Add(childNode.Name, new List<TreeNode>());
                    }

                    continue;
                }

                var schema = GetNodeSchema(childNode);

                if (string.IsNullOrEmpty(schema))
                {
                    continue;
                }

                if (!node.Nodes.ContainsKey(schema))
                {
                    TreeNode schemaNode;

                    if (Options.CloneParentNode)
                    {
                        schemaNode = new SchemaFolderTreeNode(node);
                        _ = node.Nodes.Add(schemaNode);
                    }
                    else
                    {
                        schemaNode = node.Nodes.Add(schema);
                    }

                    schemaNode.Name = schema;
                    schemaNode.Text = schema;
                    schemaNode.Tag = nodeTag;

                    if (Options.AppendDot)
                    {
                        schemaNode.Text += ".";
                    }

                    if (Options.UseObjectIcon)
                    {
                        schemaNode.ImageIndex = childNode.ImageIndex;
                        schemaNode.SelectedImageIndex = childNode.ImageIndex;
                    }
                    else
                    {
                        schemaNode.ImageIndex = node.ImageIndex;
                        schemaNode.SelectedImageIndex = node.ImageIndex;
                    }
                }

                if (!schemas.TryGetValue(schema, out var schemaNodeList))
                {
                    schemaNodeList = new List<TreeNode>();
                    schemas.Add(schema, schemaNodeList);
                }

                schemaNodeList.Add(childNode);
            }

            foreach (var schema in schemas.Keys)
            {
                var schemaNode = node.Nodes[schema];

                foreach (var childNode in schemas[schema])
                {
                    node.Nodes.Remove(childNode);

                    if (Options.RenameNode)
                    {
                        RenameNode(childNode);
                    }

                    _ = schemaNode.Nodes.Add(childNode);
                }
            }

            node.TreeView.EndUpdate();

            return schemas.Count;
        }

        #endregion
    }
}