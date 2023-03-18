using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;

namespace VCFolderView
{
    public sealed class RevealCommand
    {
        public static RevealCommand instance { get; private set; }

        public const int commandId = 0x110A;
        public static readonly Guid commandSet = new Guid("5e1c6ef5-1e4f-4523-a97d-f92c88fd4d58");

        public static RevealCommand NewInstance()
        {
            instance = new RevealCommand();
            return instance;
        }

        RevealCommand()
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
                string args;
                if (selectedItem.ProjectItem?.IsFile() == true)
                {
                    args = $"/select,\"{selectedItem.ProjectItem.GetFullPath()}\"";
                }
                else
                {
                    args = path;
                }
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    Arguments = args,
                    FileName = "explorer.exe",
                };
                System.Diagnostics.Process.Start(startInfo);
            }
        }
    }
}
