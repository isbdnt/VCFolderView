using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.IO;

namespace VCFolderView
{
    public sealed class NewFileCommand
    {
        public static NewFileCommand instance { get; private set; }

        public const int commandId = 0x1101;
        public static readonly Guid commandSet = new Guid("5e1c6ef5-1e4f-4523-a97d-f92c88fd4d58");

        public static NewFileCommand NewInstance()
        {
            instance = new NewFileCommand();
            return instance;
        }

        NewFileCommand()
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
                var path = DTEUtility.CreateFolders(selectedItem);
                var name = collection.GetSafeChildName("NewFile");
                path = Path.Combine(path, name);
                try
                {
                    File.WriteAllText(path, "");
                }
                catch (Exception) { }
                collection.AddFromFile(path);
                DTEUtility.SelectAndRename(name);
            }
        }
    }
}
