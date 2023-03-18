using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;

namespace VCFolderView
{
    public sealed class NewFolderCommand
    {
        public static NewFolderCommand instance { get; private set; }

        public const int commandId = 0x1102;
        public static readonly Guid commandSet = new Guid("5e1c6ef5-1e4f-4523-a97d-f92c88fd4d58");

        public static NewFolderCommand NewInstance()
        {
            instance = new NewFolderCommand();
            return instance;
        }

        NewFolderCommand()
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
                var collection = selectedItem.GetFolderCollection();
                var newFolder = collection.AddFolder(collection.GetSafeChildName("NewFolder"), ProjectItemKinds.FILTER);
                DTEUtility.CreateFolders(newFolder);
                DTEUtility.SelectAndRename(newFolder.Name);
            }
        }
    }
}
