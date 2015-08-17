namespace JosephM.Xrm.WorkflowScheduler.Core
{
    public class ExcelFile : XrmFile
    {
        public ExcelFile(string filePath)
            : base(filePath)
        {
        }

        public override string FileMask
        {
            get { return FileMasks.ExcelFile; }
        }
    }
}