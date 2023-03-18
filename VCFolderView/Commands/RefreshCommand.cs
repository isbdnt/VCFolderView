using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;

namespace VCFolderView
{
    public sealed class RefreshCommand
    {
        public static RefreshCommand instance { get; private set; }

        public const int commandId = 0x1104;
        public static readonly Guid commandSet = new Guid("5e1c6ef5-1e4f-4523-a97d-f92c88fd4d58");

        public static RefreshCommand NewInstance()
        {
            instance = new RefreshCommand();
            return instance;
        }

        RefreshCommand()
        {
            var menuCommandID = new CommandID(commandSet, commandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            MainPackage.instance.commandService.AddCommand(menuItem);
        }

        void OnBeforeQueryStatus(object sender, EventArgs args)
        {
            var typedSender = sender as OleMenuCommand;
            typedSender.Enabled = DTEUtility.IsVCXProj() && DTEUtility.IsSingleSelected() && DTEUtility.IsSelected(SelectionFlags.Project | SelectionFlags.Filter);
        }

        void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (DTEUtility.TryGetSelectedItem(out SelectedItem selectedItem))
            {
                MainPackage.instance.RefreshConfig(selectedItem.GetProjectPath());
                var path = DTEUtility.CreateFolders(selectedItem);
                DTEUtility.Refresh(selectedItem.GetFolderCollection(), path);
            }
        }
    }
}
