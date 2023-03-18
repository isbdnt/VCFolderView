using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;

namespace VCFolderView
{
    public sealed class PasteCommand
    {
        public static PasteCommand instance { get; private set; }

        public const int commandId = 0x1109;
        public static readonly Guid commandSet = new Guid("5e1c6ef5-1e4f-4523-a97d-f92c88fd4d58");

        public static PasteCommand NewInstance()
        {
            instance = new PasteCommand();
            return instance;
        }

        PasteCommand()
        {
            var menuCommandID = new CommandID(commandSet, commandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            MainPackage.instance.commandService.AddCommand(menuItem);
        }

        void OnBeforeQueryStatus(object sender, EventArgs args)
        {
            var typedSender = sender as OleMenuCommand;
            typedSender.Enabled = DTEUtility.IsVCXProj() && DTEUtility.IsSingleSelected() && DTEUtility.IsSelected(SelectionFlags.Project | SelectionFlags.File | SelectionFlags.Filter) && (CopyCommand.instance.projectItems.Count > 0 || CutCommand.instance.projectItems.Count > 0);
        }

        void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (DTEUtility.TryGetSelectedItem(out SelectedItem selectedItem))
            {
                if (CopyCommand.instance.projectItems.Count > 0)
                {
                    if (CopyCommand.instance.projectItems.Count == 1)
                    {
                        DTEUtility.CopyItemToSelected(CopyCommand.instance.projectItems[0], selectedItem, true);
                    }
                    else
                    {
                        foreach (var projectItem in CopyCommand.instance.projectItems)
                        {
                            DTEUtility.CopyItemToSelected(projectItem, selectedItem);
                        }
                    }
                }
                else if (CutCommand.instance.projectItems.Count > 0)
                {
                    if (CutCommand.instance.projectItems.Count == 1)
                    {
                        DTEUtility.MoveItemToSelected(CutCommand.instance.projectItems[0], selectedItem, true);
                    }
                    else
                    {
                        foreach (var projectItem in CutCommand.instance.projectItems)
                        {
                            DTEUtility.MoveItemToSelected(projectItem, selectedItem);
                        }
                    }
                    CutCommand.instance.projectItems.Clear();
                }
            }
        }
    }
}
