using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.IO;

namespace VCFolderView
{
    public sealed class DeleteCommand
    {
        public static DeleteCommand instance { get; private set; }

        public const int commandId = 0x1106;
        public static readonly Guid commandSet = new Guid("5e1c6ef5-1e4f-4523-a97d-f92c88fd4d58");

        public static DeleteCommand NewInstance()
        {
            instance = new DeleteCommand();
            return instance;
        }

        DeleteCommand()
        {
            var menuCommandID = new CommandID(commandSet, commandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            MainPackage.instance.commandService.AddCommand(menuItem);
        }

        void OnBeforeQueryStatus(object sender, EventArgs args)
        {
            var typedSender = sender as OleMenuCommand;
            typedSender.Enabled = DTEUtility.IsVCXProj() && DTEUtility.IsSelected(SelectionFlags.File | SelectionFlags.Filter);
        }

        void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (SelectedItem selectedItem in DTEUtility.GetSelections())
            {
                var path = DTEUtility.CreateFolders(selectedItem.ProjectItem);
                try
                {
                    if (selectedItem.ProjectItem.IsFile())
                    {
                        File.Delete(selectedItem.ProjectItem.GetFullPath());
                    }
                    else
                    {
                        Directory.Delete(path, true);
                    }
                    selectedItem.ProjectItem.Remove();
                }
                catch (Exception) { }
            }
        }
    }
}
