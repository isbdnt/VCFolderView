using EnvDTE;
using Microsoft.VisualStudio.PlatformUI;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace VCFolderView
{
    [Guid("f88da20c-0c11-43ca-9b44-d2fca1ad8a93")]
    public class RenameDialog : DialogWindow
    {
        public ProjectItem projectItem => _projectItem;

        ProjectItem _projectItem;

        public RenameDialog(ProjectItem projectItem) : base("")
        {
            _projectItem = projectItem;
            Title = "Rename";
            Content = new RanameDialogControl(this);
            Width = 396;
            Height = 126;
            ResizeMode = System.Windows.ResizeMode.NoResize;
        }

        public void Rename(string newName)
        {
            try
            {
                if (_projectItem.IsFile())
                {
                    var parent = DTEUtility.CreateFolders(_projectItem);
                    var oldPath = Path.Combine(parent, _projectItem.Name);
                    var newPath = Path.Combine(parent, newName);
                    Directory.Move(oldPath, newPath);
                }
                else if (_projectItem.IsFilter())
                {
                    var oldPath = DTEUtility.CreateFolders(_projectItem);
                    var parent = Path.GetDirectoryName(oldPath);
                    var newPath = Path.Combine(parent, newName);
                    Directory.Move(oldPath, newPath);
                }
                _projectItem.Name = newName;
            }
            catch (Exception) { }
            Close();
        }
    }
}
