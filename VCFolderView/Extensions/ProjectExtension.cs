using EnvDTE;
using System;

namespace VCFolderView
{
    public static class ProjectExtension
    {
        public static bool IsKind(this Project project, string guid)
        {
            return project.Kind.Equals(guid, StringComparison.OrdinalIgnoreCase);
        }
    }
}
