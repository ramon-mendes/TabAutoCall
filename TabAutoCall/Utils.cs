using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;

namespace TabAutoCall
{
	static class Utils
	{
		static Utils()
		{
			DTE = ServiceProvider.GlobalProvider.GetService(typeof(DTE)) as DTE2;
			_vsMEFcontainer = ServiceProvider.GlobalProvider.GetService(typeof(SComponentModel)) as IComponentModel;
		}

		private static readonly DTE2 DTE;
		private static IComponentModel _vsMEFcontainer;

		public static T GetService<T>() where T : class
		{
			var TT = ServiceProvider.GlobalProvider.GetService(typeof(IMenuCommandService)) as IMenuCommandService;
			return _vsMEFcontainer.GetService<T>();
		}
	}
}