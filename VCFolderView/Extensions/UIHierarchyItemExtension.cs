using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VCFolderView
{
    public static class UIHierarchyItemExtension
    {
        public static bool TryGetChild(this UIHierarchyItem hierarchyItem, string name, out UIHierarchyItem childHierarchyItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (UIHierarchyItem child in hierarchyItem.UIHierarchyItems)
            {
                if (child.Name == name)
                {
                    childHierarchyItem = child;
                    return true;
                }
            }
            childHierarchyItem = null;
            return false;
        }
    }
}
