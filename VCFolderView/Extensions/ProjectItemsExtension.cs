using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;

namespace VCFolderView
{
    public static class ProjectItemsExtension
    {
        public static string GetSafeChildName(this ProjectItems projectItems, string basaName, bool isCopy = false)
        {
            int nameIndex = 1;
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(basaName);
            string extension = Path.GetExtension(basaName);
            string safeName = $"{nameWithoutExtension}{extension}";
            HashSet<string> projectItemNameSet = new HashSet<string>();
            string copy = isCopy ? "_Copy" : "";
            foreach (ProjectItem projectItem in projectItems)
            {
                projectItemNameSet.Add(projectItem.Name);
            }
            if (projectItemNameSet.Contains(safeName))
            {
                safeName = $"{nameWithoutExtension}{copy}{extension}";
            }
            while (projectItemNameSet.Contains(safeName))
            {
                safeName = $"{nameWithoutExtension}{copy}_{++nameIndex}{extension}";
            }
            return safeName;
        }

        public static ProjectItem FindChild(this ProjectItems projectItems, string name)
        {
            foreach (ProjectItem projectItem in projectItems)
            {
                if (projectItem.Name == name)
                {
                    return projectItem;
                }
            }
            return null;
        }
    }
}
