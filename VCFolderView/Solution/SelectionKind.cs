using System;

namespace VCFolderView
{
    [Flags]
    public enum SelectionFlags
    {
        None = 0,
        Project = 1 << 0,
        File = 1 << 1,
        Filter = 1 << 2,
    }
}
