using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace TabAutoCall
{
	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[Guid(TabAutoCallPackage.PackageGuidString)]
	[ProvideService(typeof(TabAutoCallPackage), IsAsyncQueryable = true)]
	public sealed class TabAutoCallPackage : AsyncPackage
	{
		public static IMenuCommandService MenuService;
		public const string PackageGuidString = "61483a86-fe46-425d-a0bb-e0a895a5f92f";

		protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
		{
			MenuService = (IMenuCommandService) await GetServiceAsync(typeof(IMenuCommandService));

			// When initialized asynchronously, the current thread may be a background thread at this point.
			// Do any initialization that requires the UI thread after switching to the UI thread.
			await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
		}
	}
}