using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;

namespace VCFolderView
{
    public sealed class CopyCommand
    {
        public static CopyCommand instance { get; private set; }

        public const int commandId = 0x1107;
        public static readonly Guid commandSet = new Guid("5e1c6ef5-1e4f-4523-a97d-f92c88fd4d58");
        public List<ProjectItem> projectItems { get; } = new List<ProjectItem>();

        public static CopyCommand NewInstance()
        {
            instance = new CopyCommand();
            return instance;
        }

        CopyCommand()
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
            projectItems.Clear();
            foreach (SelectedItem selectedItem in DTEUtility.GetSelections())
            {
                projectItems.Add(selectedItem.ProjectItem);
            }
            CutCommand.instance.projectItems.Clear();
        }
    }
}
