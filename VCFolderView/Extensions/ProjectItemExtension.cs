using EnvDTE;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VCFolderView
{
    public static class ProjectItemExtension
    {
        public static bool IsKind(this ProjectItem projectItem, string guid)
        {
            return projectItem.Kind.Equals(guid, StringComparison.OrdinalIgnoreCase);
        }

        public static string GetFullPath(this ProjectItem projectItem)
        {
            return projectItem.Properties.Item("FullPath").Value as string;
        }

        public static bool IsFile(this ProjectItem projectItem)
        {
            return projectItem.IsKind(ProjectItemKinds.FILE);
        }

        public static bool IsFilter(this ProjectItem projectItem)
        {
            return projectItem.IsKind(ProjectItemKinds.FILTER);
        }
    }
}
