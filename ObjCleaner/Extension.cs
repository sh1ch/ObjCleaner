using EnvDTE;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ObjCleaner
{
	internal static class Extension
	{
		internal static bool TryGetPropertyValue<T>(this EnvDTE.Properties properties, string index, out T result)
		{
			Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

			T value = default(T);
			bool canGet = false;

			try
			{
				value = (T)properties.Item(index).Value;
				canGet = true;
			}
			catch
			{
				canGet = false;
			}

			result = value;

			return canGet;
		}

		internal static void WriteLineThreadSafe(this IVsOutputWindowPane pane, string text)
		{
			Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

			pane?.OutputStringThreadSafe(text + Environment.NewLine);
		}

		internal static void WriteThreadSafe(this IVsOutputWindowPane pane, string text)
		{
			Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

			pane?.OutputStringThreadSafe(text);
		}
	}
}
