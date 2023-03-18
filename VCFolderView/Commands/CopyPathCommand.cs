using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace VCFolderView
{
    public sealed class CopyPathCommand
    {
        public static CopyPathCommand instance { get; private set; }

        public const int commandId = 0x110C;
        public static readonly Guid commandSet = new Guid("5e1c6ef5-1e4f-4523-a97d-f92c88fd4d58");

        public static CopyPathCommand NewInstance()
        {
            instance = new CopyPathCommand();
            return instance;
        }

        CopyPathCommand()
        {
            var menuCommandID = new CommandID(commandSet, commandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            MainPackage.instance.commandService.AddCommand(menuItem);
        }

        void OnBeforeQueryStatus(object sender, EventArgs args)
        {
            var typedSender = sender as OleMenuCommand;
            typedSender.Enabled = DTEUtility.IsVCXProj() && DTEUtility.IsSingleSelected() && DTEUtility.IsSelected(SelectionFlags.Project | SelectionFlags.File | SelectionFlags.Filter);
        }

        void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (DTEUtility.TryGetSelectedItem(out SelectedItem selectedItem))
            {
                var path = DTEUtility.CreateFolders(selectedItem);
                if (selectedItem.Project != null)
                {
                    Clipboard.SetText(Path.GetDirectoryName(selectedItem.Project.FullName));
                }
                else if (selectedItem.ProjectItem.IsFile())
                {
                    Clipboard.SetText(selectedItem.ProjectItem.GetFullPath());
                }
                else
                {
                    Clipboard.SetText(path);
                }
            }
        }
    }
}
