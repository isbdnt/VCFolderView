using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace VCFolderView
{
    public partial class RanameDialogControl : UserControl
    {
        RenameDialog _renameDialog;

        public RanameDialogControl(RenameDialog renameDialog)
        {
            _renameDialog = renameDialog;
            InitializeComponent();
            NameText.Text = renameDialog.projectItem.Name;
            NameText.SelectAll();
            NameText.Focus();
            var enterCmd = new RoutedCommand();
            enterCmd.InputGestures.Add(new KeyGesture(Key.Enter));
            CommandBindings.Add(new CommandBinding(enterCmd, OKButton_Click));
            var escCmd = new RoutedCommand();
            escCmd.InputGestures.Add(new KeyGesture(Key.Escape));
            CommandBindings.Add(new CommandBinding(escCmd, CancelButton_Click));
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            _renameDialog.Rename(NameText.Text);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _renameDialog.Close();
        }
    }
}