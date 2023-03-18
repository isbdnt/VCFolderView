using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using MiniJSON;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace VCFolderView
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(RenameDialog))]
    public sealed class MainPackage : AsyncPackage
    {
        public const string PackageGuidString = "7c020ffe-3c9c-4196-878d-88f8502c1dd7";
        public static MainPackage instance { get; private set; }
        public DTE2 dte { get; private set; }
        public OleMenuCommandService commandService { get; private set; }
        public IVsOutputWindowPane output { get; private set; }

        Dictionary<string, Config> _configMap = new Dictionary<string, Config>();

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            instance = this;
            dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            commandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            output = await GetServiceAsync(typeof(SVsGeneralOutputWindowPane)) as IVsOutputWindowPane;
            RefreshCommand.NewInstance();
            RevealCommand.NewInstance();
            NewFileCommand.NewInstance();
            NewFolderCommand.NewInstance();
            DuplicateCommand.NewInstance();
            CopyCommand.NewInstance();
            CutCommand.NewInstance();
            PasteCommand.NewInstance();
            RemoveCommand.NewInstance();
            DeleteCommand.NewInstance();
            RenameCommand.NewInstance();
            CopyPathCommand.NewInstance();
        }

        public Config RefreshConfig(string projectPath)
        {
            var projectDirectory = Path.GetDirectoryName(projectPath);
            var configPath = Path.Combine(projectDirectory, "VCFolderView.json");
            Config config = null;
            if (File.Exists(configPath))
            {
                Dictionary<string, object> json = null;
                try
                {
                    json = Json.Deserialize(File.ReadAllText(configPath)) as Dictionary<string, object>;
                }
                catch (Exception) { }
                if (json != null)
                {
                    config = new Config()
                    {
                        WorkingDirectory = json.ContainsKey("WorkingDirectory") ? Path.GetFullPath(Path.Combine(projectDirectory, (string)json["WorkingDirectory"])) : projectDirectory,
                        ExcludeRules = new List<Regex>(),
                    };
                    if (json.ContainsKey("ExcludeRules"))
                    {
                        foreach (string rule in json["ExcludeRules"] as List<object>)
                        {
                            config.ExcludeRules.Add(new Regex(rule));
                        }
                    }
                }
            }
            if (config == null)
            {
                config = new Config
                {
                    WorkingDirectory = projectDirectory,
                    ExcludeRules = new List<Regex>(),
                };
            }
            _configMap[projectPath] = config;
            return config;
        }

        public Config GetConfig(string projectPath)
        {
            if (!_configMap.TryGetValue(projectPath, out Config config))
            {
                return RefreshConfig(projectPath);
            }
            return config;
        }
    }
}
