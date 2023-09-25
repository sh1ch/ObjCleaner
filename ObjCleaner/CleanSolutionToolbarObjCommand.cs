using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ObjCleaner.Properties;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace ObjCleaner
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class CleanSolutionToolbarObjCommand
	{
		private readonly DTE2 _DTE2;

		private Cleaner Cleaner = null;

		private readonly IVsOutputWindowPane _Output;

		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 4129;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("5610e463-3ba8-4627-be71-238be8f85011");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly AsyncPackage package;

		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static CleanSolutionToolbarObjCommand Instance
		{
			get;
			private set;
		}

		/// <summary>
		/// Gets the service provider from the owner package.
		/// </summary>
		private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
		{
			get
			{
				return this.package;
			}
		}

		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		public static async Task InitializeAsync(AsyncPackage package)
		{
			// Switch to the main thread - the call to AddCommand in ObjToolbar's constructor requires
			// the UI thread.
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

			var dte = await package.GetServiceAsync<DTE, DTE2>();

			Instance = new CleanSolutionToolbarObjCommand(package, commandService, dte);
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="CleanSolutionToolbarObjCommand"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		/// <param name="commandService">Command service to add command to, not null.</param>
		private CleanSolutionToolbarObjCommand(AsyncPackage package, OleMenuCommandService commandService, DTE2 dte2)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			_DTE2 = dte2;

			this.package = package ?? throw new ArgumentNullException(nameof(package));
			commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

			var menuCommandID = new CommandID(CommandSet, CommandId);
			var menuItem = new MenuCommand(this.Execute, menuCommandID);

			// 出力ウィンドウ
			var outputWindow = (IVsOutputWindow)Package.GetGlobalService(typeof(SVsOutputWindow));
			var generalPaneGuid = VSConstants.GUID_BuildOutputWindowPane; // or GUID_OutWindowDebugPane

			outputWindow.GetPane(ref generalPaneGuid, out _Output);
			_Output.Activate();

			Cleaner = new Cleaner(_Output);

			commandService.AddCommand(menuItem);
		}

		/// <summary>
		/// This function is the callback used to execute the command when the menu item is clicked.
		/// See the constructor to see how the menu item is associated with this function using
		/// OleMenuCommandService service and MenuCommand class.
		/// </summary>
		/// <param name="sender">Event sender.</param>
		/// <param name="e">Event args.</param>
		private void Execute(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// 対象が０件（未指定）のときは、なにもしない
			if (CleanData.Directories.Length <= 0) return;

			_Output.WriteLineThreadSafe(TextResources.CleanStart);

			var selectedProject = Cleaner.GetSelectedProject(_DTE2);

			if (selectedProject == null)
			{
				return;
			}

			_Output.WriteLineThreadSafe(string.Format(TextResources.CleanTargetProject, selectedProject.Name));

			var projectRootDirectory = Cleaner.GetProjectRootDirectory(selectedProject);

			Cleaner.Clean(projectRootDirectory);
		}
	}
}
