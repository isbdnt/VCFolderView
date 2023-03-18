using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace VCFolderView
{
    public static class DTEUtility
    {
        public static bool IsVCXProj()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (SelectedItem selectedItem in MainPackage.instance.dte.SelectedItems)
            {
                if (selectedItem.Project == null)
                {
                    if (selectedItem.ProjectItem == null)
                    {
                        return false;
                    }
                    if (!selectedItem.ProjectItem.ContainingProject.IsKind(ProjectKinds.VC))
                    {
                        return false;
                    }
                }
                else
                {
                    if (!selectedItem.Project.IsKind(ProjectKinds.VC))
                    {
                        return false;
                    }
                }
            }
            return MainPackage.instance.dte.SelectedItems.Count > 0;
        }

        public static bool IsSingleSelected()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return MainPackage.instance.dte.SelectedItems.Count == 1;
        }

        public static bool IsSelected(SelectionFlags flags)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (SelectedItem selectedItem in MainPackage.instance.dte.SelectedItems)
            {
                if (selectedItem.Project == null)
                {
                    if (selectedItem.ProjectItem == null)
                    {
                        return false;
                    }
                    if (!selectedItem.ProjectItem.IsFile() && !selectedItem.ProjectItem.IsFilter())
                    {
                        return false;
                    }
                    if (selectedItem.ProjectItem.IsFile() && !flags.HasFlag(SelectionFlags.File))
                    {
                        return false;
                    }
                    else if (selectedItem.ProjectItem.IsFilter() && !flags.HasFlag(SelectionFlags.Filter))
                    {
                        return false;
                    }
                }
                else
                {
                    if (!flags.HasFlag(SelectionFlags.Project))
                    {
                        return false;
                    }
                }
            }
            return MainPackage.instance.dte.SelectedItems.Count > 0;
        }

        public static bool TryGetSelectedItem(out SelectedItem selectedItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SelectedItems selectedItems = MainPackage.instance.dte.SelectedItems;
            if (selectedItems.Count == 1)
            {
                var enumerator = selectedItems.GetEnumerator();
                if (enumerator.MoveNext())
                {
                    selectedItem = enumerator.Current as SelectedItem;
                    return selectedItem != null;
                }
            }
            selectedItem = null;
            return false;
        }

        public static List<SelectedItem> GetSelections()
        {
            var selections = new List<SelectedItem>();
            foreach (SelectedItem selectedItem in MainPackage.instance.dte.SelectedItems)
            {
                selections.Add(selectedItem);
            }
            return selections;
        }

        public static bool TryGetSelectedHierarchyItem(out UIHierarchyItem hierarchyItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            UIHierarchy solExplorer = MainPackage.instance.dte.ToolWindows.SolutionExplorer;
            var selectedHierarchyItems = solExplorer.SelectedItems as UIHierarchyItem[];
            if (selectedHierarchyItems.Length == 1)
            {
                hierarchyItem = selectedHierarchyItems[0];
                return hierarchyItem != null;
            }
            hierarchyItem = null;
            return false;
        }

        public static void SelectAndRename(string name)
        {
            if (Select(name))
            {
                //MainPackage.instance.dte.ExecuteCommand("File.Rename");
            }
        }

        public static bool Select(string name)
        {
            if (DTEUtility.TryGetSelectedHierarchyItem(out UIHierarchyItem hierarchyItem))
            {
                var projectItem = hierarchyItem.Object as ProjectItem;
                if (projectItem != null && projectItem.IsFile())
                {
                    hierarchyItem = hierarchyItem.Collection.Parent as UIHierarchyItem;
                }
                if (hierarchyItem.TryGetChild(name, out hierarchyItem))
                {
                    hierarchyItem.Select(vsUISelectionType.vsUISelectionTypeSelect);
                    return true;
                }
            }
            return false;
        }

        public static bool SelectParent()
        {
            if (DTEUtility.TryGetSelectedHierarchyItem(out UIHierarchyItem hierarchyItem))
            {
                var parent = hierarchyItem.Collection.Parent as UIHierarchyItem;
                if (parent != null)
                {
                    parent.Select(vsUISelectionType.vsUISelectionTypeSelect);
                    return true;
                }
            }
            return false;
        }

        public static string CreateFolders(SelectedItem selectedItem)
        {
            if (selectedItem.ProjectItem != null)
            {
                return CreateFolders(selectedItem.ProjectItem);
            }
            return MainPackage.instance.GetConfig(selectedItem.Project.FullName).WorkingDirectory;
        }

        public static string CreateFolders(ProjectItem projectItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var path = MainPackage.instance.GetConfig(projectItem.ContainingProject.FullName).WorkingDirectory;
            ProjectItem currentProjectItem;
            if (projectItem.IsFile())
            {
                currentProjectItem = projectItem.Collection.Parent as ProjectItem;
            }
            else
            {
                currentProjectItem = projectItem;
            }
            var folders = new Stack<string>();
            while (currentProjectItem != null)
            {
                folders.Push(currentProjectItem.Name);
                currentProjectItem = currentProjectItem.Collection.Parent as ProjectItem;
            }
            while (folders.Count > 0)
            {
                var folderName = folders.Pop();
                path = Path.Combine(path, folderName);
                try
                {
                    Directory.CreateDirectory(path);
                }
                catch (Exception) { }
            }
            return path;
        }

        public static void CopyItemToSelected(ProjectItem originalItem, SelectedItem selectedItem, bool select = false)
        {
            ProjectItems collection = selectedItem.GetFolderCollection();
            var originalPath = DTEUtility.CreateFolders(originalItem);
            var targetFolderPath = DTEUtility.CreateFolders(selectedItem);
            var name = collection.GetSafeChildName(originalItem.Name, true);
            var targetPath = Path.Combine(targetFolderPath, name);
            try
            {
                if (originalItem.IsFile())
                {
                    File.Copy(originalItem.GetFullPath(), targetPath, true);
                    collection.AddFromFile(targetPath);
                }
                else
                {
                    CopyDirectory(originalPath, targetPath);
                    Refresh(collection, targetFolderPath);
                }
            }
            catch (Exception) { }
            if (select)
            {
                DTEUtility.SelectAndRename(name);
            }
        }

        public static void MoveItemToSelected(ProjectItem originalItem, SelectedItem selectedItem, bool select = false)
        {
            ProjectItems collection = selectedItem.GetFolderCollection();
            if (originalItem.Collection == collection)
            {
                return;
            }
            var originalPath = DTEUtility.CreateFolders(originalItem);
            var targetFolderPath = DTEUtility.CreateFolders(selectedItem);
            var name = originalItem.Name;
            var targetPath = Path.Combine(targetFolderPath, name);
            try
            {
                if (originalItem.IsFile())
                {
                    if (!File.Exists(targetPath) || MessageBox.Show($"The destination already has a file named \"{name}\"", "Folder View", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        if (File.Exists(targetPath))
                        {
                            collection.FindChild(name)?.Remove();
                            File.Delete(targetPath);
                        }
                        File.Move(originalItem.GetFullPath(), targetPath);
                        originalItem.Remove();
                        collection.AddFromFile(targetPath);
                        if (select)
                        {
                            Select(name);
                        }
                    }
                }
                else
                {
                    if (!Directory.Exists(targetPath) || MessageBox.Show($"The destination already has a folder named \"{name}\"", "Folder View", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        if (Directory.Exists(targetPath))
                        {
                            collection.FindChild(name)?.Remove();
                            Directory.Delete(targetPath, true);
                        }
                        Directory.Move(originalPath, targetPath);
                        originalItem.Remove();
                        Refresh(collection, targetFolderPath);
                        if (select)
                        {
                            Select(name);
                        }
                    }
                }
            }
            catch (Exception) { }
        }

        static void CopyDirectory(string from, string to)
        {
            var dirs = Directory.GetDirectories(from);
            Directory.CreateDirectory(to);
            foreach (var file in Directory.GetFiles(from))
            {
                File.Copy(file, Path.Combine(to, Path.GetFileName(file)), true);
            }
            foreach (var dir in dirs)
            {
                var toDir = Path.Combine(to, Path.GetFileName(dir));
                Directory.CreateDirectory(toDir);
                CopyDirectory(dir, toDir);
            }
        }

        static List<ProjectItem> _invalidItems = new List<ProjectItem>();

        public static void Refresh(ProjectItems projectItems, string directory)
        {
            Refresh(MainPackage.instance.GetConfig(projectItems.ContainingProject.FullName), projectItems, directory);
        }

        public static void Refresh(Config config, ProjectItems projectItems, string directory)
        {
            Directory.CreateDirectory(directory);
            foreach (ProjectItem childProjectItem in projectItems)
            {
                _invalidItems.Add(childProjectItem);
            }
            foreach (var invalidItem in _invalidItems)
            {
                try
                {
                    invalidItem.Remove();
                }
                catch (Exception) { }
            }
            _invalidItems.Clear();
            foreach (var childFile in Directory.GetFiles(directory))
            {
                if (!config.IsExcluded(childFile))
                {
                    try
                    {
                        projectItems.AddFromFile(childFile);
                    }
                    catch (Exception) { }
                }
            }
            foreach (var childDirectory in Directory.GetDirectories(directory))
            {
                if (!config.IsExcluded(childDirectory))
                {
                    var name = Path.GetFileName(childDirectory);
                    try
                    {
                        var childFolder = projectItems.AddFolder(name, ProjectItemKinds.FILTER);
                        Refresh(config, childFolder.ProjectItems, childDirectory);
                    }
                    catch (Exception) { }
                }
            }
        }
    }
}
