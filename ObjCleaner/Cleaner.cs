using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ObjCleaner.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjCleaner
{
	public class Cleaner
	{
		private IVsOutputWindowPane _Output;

		/// <summary>
		/// <see cref="Cleaner"/> クラスの新しいインスタンスを初期化します。
		/// </summary>
		public Cleaner(IVsOutputWindowPane output)
		{
			_Output = output;
		}

		public void Clean(string projectRootPath)
		{
			var (success1, failure1) = CleanFiles(projectRootPath);
			var (success2, failure2) = CleanDirectories(projectRootPath);

			int success = (success1 + success2);
			int failure = (failure1 + failure2);

			_Output.WriteLineThreadSafe(string.Format(TextResources.CleanReport, success, failure));
		}

		public (int, int) CleanDirectories(string projectRootPath)
		{
			if (string.IsNullOrEmpty(projectRootPath))
			{
				return (0, 0);
			}

			int success = 0;
			int failure = 0;

			foreach (var info in CleanData.Directories
						.Select(p => Path.Combine(projectRootPath, p))
						.Where(Directory.Exists)
						.Select(p => new DirectoryInfo(p)))
			{
				foreach (var dir in info.EnumerateDirectories())
				{
					_Output.WriteThreadSafe(string.Format(TextResources.DeleteStart, dir.Name));

					try
					{
						dir.Delete(true);
						success += 1;

						_Output.WriteLineThreadSafe(TextResources.DeleteSuccessed);
					}
					catch (Exception ex)
					{
						_Output.WriteLineThreadSafe(ex.ToString());
						failure += 1;

						_Output.WriteLineThreadSafe(string.Format(TextResources.DeleteFailed, ex));
					}
				}
			}

			return (success, failure);
		}

		public (int, int) CleanFiles(string projectRootPath)
		{
			if (string.IsNullOrEmpty(projectRootPath))
			{
				return (0, 0);
			}

			int success = 0;
			int failure = 0;

			foreach (var info in CleanData.Directories
						.Select(p => Path.Combine(projectRootPath, p))
						.Where(Directory.Exists)
						.Select(p => new DirectoryInfo(p)))
			{
				foreach (var file in info.EnumerateFiles())
				{
					_Output.WriteThreadSafe(string.Format(TextResources.DeleteStart, file.Name));

					try
					{
						file.Delete();
						success += 1;

						_Output.WriteLineThreadSafe(TextResources.DeleteSuccessed);
					}
					catch (Exception ex)
					{
						_Output.WriteLineThreadSafe(ex.ToString());
						failure += 1;

						_Output.WriteLineThreadSafe(string.Format(TextResources.DeleteFailed, ex));
					}
				}
			}

			return (success, failure);
		}

		public static string GetProjectRootDirectory(Project project)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (string.IsNullOrEmpty(project.FullName))
			{
				return "";
			}

			var fullPath = "";
			var canGet = project.Properties.TryGetPropertyValue("FullPath", out fullPath);

			if (!canGet)
			{
				canGet = project.Properties.TryGetPropertyValue("ProjectDirectory", out fullPath);

				if (!canGet)
				{
					canGet = project.Properties.TryGetPropertyValue("ProjectPath", out fullPath);

					if (!canGet)
					{
						fullPath = "";
					}
				}
			}

			// プロパティで取得できないとき
			if (string.IsNullOrEmpty(fullPath))
			{
				fullPath = File.Exists(project.FullName) ? Path.GetDirectoryName(project.FullName) : "";
			}

			return fullPath;
		}

		public static IEnumerable<Project> GetActiveProjects(DTE2 dte2)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var projects = new List<Project>();

			if (dte2.ActiveSolutionProjects is Array activeProjects && (activeProjects?.Length ?? 0) > 0)
			{
				var temp = activeProjects.Cast<Project>().ToList();

				projects.AddRange(temp);
			}

			if (projects.Count <= 0)
			{
				var temp = dte2.Solution.Projects.Cast<Project>().ToList();

				projects.AddRange(temp);
			}

			return projects;
		}

		public static Project GetSelectedProject(DTE2 dte2)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Project selectedProject = null;

			// 対象が１件以外のときは、なにもしない
			if (dte2.SelectedItems.Count != 1) return null;

			selectedProject = dte2.SelectedItems.Item(1)?.Project ?? null;

			return selectedProject;
		}
	}
}
