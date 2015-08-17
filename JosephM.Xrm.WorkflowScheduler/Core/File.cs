namespace JosephM.Xrm.WorkflowScheduler.Core
{
    public abstract class XrmFile
    {
        protected XrmFile(string fileName)
        {
            FileName = fileName;
        }

        public string FileName { get; set; }

        public abstract string FileMask { get; }
    }
}