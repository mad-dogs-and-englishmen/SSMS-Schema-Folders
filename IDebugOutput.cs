namespace SsmsSchemaFolders
{
    public interface IDebugOutput
    {
        void DebugMessage(string message);

        void DebugMessage(string message, params object[] args);

    }
}