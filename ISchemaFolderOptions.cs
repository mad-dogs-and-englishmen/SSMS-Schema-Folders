namespace SsmsSchemaFolders
{
    public interface ISchemaFolderOptions
    {
        bool AppendDot { get; }
        bool CloneParentNode { get; }
        bool Enabled { get; }
        bool RenameNode { get; }
        bool UseObjectIcon { get; }
    }
}