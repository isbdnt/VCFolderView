using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.Design;

namespace VCFolderView
{
    public static class SelectedItemExtension
    {
        public static ProjectItems GetFolderCollection(this SelectedItem selectedItem)
        {
            if (selectedItem.Project != null)
            {
                return selectedItem.Project.ProjectItems;
            }
            else
            {
                if (selectedItem.ProjectItem.IsFile())
                {
                    return selectedItem.ProjectItem.Collection;
                }
                else
                {
                    return selectedItem.ProjectItem.ProjectItems;
                }
            }
        }

        public static string GetProjectPath(this SelectedItem selectedItem)
        {
            string projectPath;
            if (selectedItem.Project != null)
            {
                return selectedItem.Project.FullName;
            }
            else
            {
                return selectedItem.ProjectItem.ContainingProject.FullName;
            }
        }
    }
}
