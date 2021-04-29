using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Threading;
using Microsoft.VisualStudio.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace TabAutoCall
{
    [Export(typeof(IWpfTextViewConnectionListener))]
	[ContentType("text")]
	[TextViewRole(PredefinedTextViewRoles.Interactive)]
	internal class CompletionController : IWpfTextViewConnectionListener
	{
		public static string[] SupportedContentTypes = new string[] { "CSharp", "C/C++", "Basic", "Javascript", "TypeScript" };

		[Import]
		internal IVsEditorAdaptersFactoryService AdaptersFactory = null;

		[Import]
		internal IEditorOperationsFactoryService OperationsService = null;

		[Import]
		internal SVsServiceProvider ServiceProvider = null;

		[Import]
		internal ICompletionBroker CompletionBroker = null;

		public void SubjectBuffersConnected(IWpfTextView textView, ConnectionReason reason, Collection<ITextBuffer> subjectBuffers)
		{
			IVsShell shell = ServiceProvider.GetService(typeof(IVsShell)) as IVsShell;
			if(shell == null) return;

			IVsPackage package = null;
			Guid PackageToBeLoadedGuid =
				new Guid(TabAutoCallPackage.PackageGuidString);
			shell.LoadPackage(ref PackageToBeLoadedGuid, out package);

			IVsTextView textViewAdapter = AdaptersFactory.GetViewAdapter(textView);
			IEditorOperations operations = OperationsService.GetEditorOperations(textView);
			TrackState state = textView.Properties.GetOrCreateSingletonProperty(() => new TrackState());

			foreach(var item in subjectBuffers)
			{
				if(SupportedContentTypes.Contains(item.ContentType.TypeName, StringComparer.OrdinalIgnoreCase))
				{
					textView.Properties.GetOrCreateSingletonProperty(() => new CommandFilter(state, textView, textViewAdapter, operations, ServiceProvider, CompletionBroker));
					textView.Properties.GetOrCreateSingletonProperty(() => new CommandFilterHighPriority(state, textView, textViewAdapter, ServiceProvider, CompletionBroker));
				}
			}
		}

		public void SubjectBuffersDisconnected(IWpfTextView textView, ConnectionReason reason, Collection<ITextBuffer> subjectBuffers)
		{
			IVsTextView textViewAdapter = AdaptersFactory.GetViewAdapter(textView);
			foreach(var item in subjectBuffers)
			{
				if(textView.Properties.ContainsProperty(typeof(CommandFilterHighPriority)))
				{
					textViewAdapter.RemoveCommandFilter(textView.Properties[typeof(CommandFilter)] as IOleCommandTarget);
					textViewAdapter.RemoveCommandFilter(textView.Properties[typeof(CommandFilterHighPriority)] as IOleCommandTarget);
				}
			}
		}
	}

	internal class TrackState
	{
		public ICompletionSession _activeSession;
		public bool _justCompletedFunc;
		public bool _justCompletedAutoBrace;
		public ITrackingSpan _justCompletedAutoBraceTrack;
		public ITextVersion _justCompletedWordVer;
	}

	internal class CommandFilterHighPriority : IOleCommandTarget
	{
		private ITextView TextView;
		private SVsServiceProvider Service;
		private ICompletionBroker Broker;

		private IOleCommandTarget _nextCommandTarget;
		private TrackState _state;

		public CommandFilterHighPriority(TrackState state, ITextView textView, IVsTextView textViewAdapter, SVsServiceProvider service, ICompletionBroker broker)
		{
			ThreadHelper.JoinableTaskFactory.Run(async delegate {
				await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
				// needed, else you don't catch AUTOCOMPLETE/COMPLETEWORD
				ErrorHandler.ThrowOnFailure(textViewAdapter.AddCommandFilter(this, out _nextCommandTarget));
			});

			_state = state;
			TextView = textView;
			Service = service;
			Broker = broker;
		}

		public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
		{
			if(VsShellUtilities.IsInAutomationFunction(Service))
				return _nextCommandTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

			if(pguidCmdGroup == VSConstants.VSStd2K)
			{
				switch((VSConstants.VSStd2KCmdID)nCmdID)
				{
					case VSConstants.VSStd2KCmdID.AUTOCOMPLETE:
					case VSConstants.VSStd2KCmdID.COMPLETEWORD:
						// let it autocomplete
						var vbefore = TextView.TextSnapshot.Version;
						_nextCommandTarget.Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

						// call our low priority to check for session
						Check4Session();

						// detects if it autocompleted
						ThreadHelper.JoinableTaskFactory.Run(async delegate {
							await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
							var vafter = TextView.TextSnapshot.Version;
							if(vbefore != vafter)
							{
								_state._justCompletedWordVer = TextView.TextSnapshot.Version;
								_state._justCompletedFunc = true;
							}
						});

						return VSConstants.S_OK;
				}
			}

			//if(pguidCmdGroup == VSConstants.VSStd2K || pguidCmdGroup == VSConstants.GUID_VSStandardCommandSet97)
			{
				Check4Session();
			}

			return _nextCommandTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
		}

		public void Check4Session()
		{
			if(_state._activeSession != null)
			{
				var sessions = Broker.GetSessions(TextView);
				//Debug.Assert(sessions[0] == _state._activeSession);
			}
			else
			{
				var sessions = Broker.GetSessions(TextView);
				Debug.Assert(sessions.Count <= 1);

				if(sessions.Count != 0)
				{
					var selection = TextView.BufferGraph.MapDownToInsertionPoint(TextView.Caret.Position.BufferPosition, PointTrackingMode.Positive,
						ts => CompletionController.SupportedContentTypes.Contains(ts.ContentType.TypeName, StringComparer.OrdinalIgnoreCase));

					if(selection.HasValue)
					{
						_state._activeSession = sessions[0];
						_state._activeSession.Committed += (object sender, EventArgs e) =>
						{
							// is never called for C# editor
							// but does works for other editors
							_state._justCompletedFunc = true;
						};

						_state._activeSession.Dismissed += (object sender, EventArgs e) =>
						{
							if(_state._activeSession.SelectedCompletionSet.SelectionStatus.Completion != null)
							{
								string completionText = _state._activeSession.SelectedCompletionSet.SelectionStatus.Completion.DisplayText;
								var applic = _state._activeSession.SelectedCompletionSet.ApplicableTo;

								ThreadHelper.JoinableTaskFactory.Run(async delegate {
									CaretPosition pos = TextView.Caret.Position;
									if(pos.BufferPosition.Position != 0)
									{
										string wordString = applic.GetText(applic.TextBuffer.CurrentSnapshot);

										if(wordString == completionText)
											_state._justCompletedFunc = true;
									}
								});
							}
							
							_state._activeSession = null;
						};
					}
				}
			}
		}

		public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
		{
			return _nextCommandTarget.QueryStatus(pguidCmdGroup, cCmds, prgCmds, pCmdText);
		}
	}

	internal class CommandFilter : IOleCommandTarget
	{
		private static DTE2 DTE;

		private ITextView TextView;
		private IEditorOperations Operations;
		private ICompletionBroker Broker;
		private SVsServiceProvider Service;

		private IOleCommandTarget _nextCommandTarget;
		private TrackState _state;

		public CommandFilter(TrackState state, ITextView textView, IVsTextView textViewAdapter, IEditorOperations operations, SVsServiceProvider service, ICompletionBroker broker)
		{
			ErrorHandler.ThrowOnFailure(textViewAdapter.AddCommandFilter(this, out _nextCommandTarget));

			if(DTE==null)
			{
				DTE = service.GetService(typeof(DTE)) as DTE2;
			}

			_state = state;
			TextView = textView;
			Operations = operations;
			Broker = broker;
			Service = service;

			textView.Caret.PositionChanged += (s, e) =>
			{
				_state._justCompletedFunc = false;

				if(_state._justCompletedAutoBrace && !_state._justCompletedAutoBraceTrack.GetSpan(e.NewPosition.BufferPosition.Snapshot).Contains(e.NewPosition.BufferPosition))
				{
					string txt = _state._justCompletedAutoBraceTrack.GetSpan(e.NewPosition.BufferPosition.Snapshot).GetText();
					_state._justCompletedAutoBrace = false;
					_state._justCompletedAutoBraceTrack = null;
				}
			};

			textView.TextBuffer.Changed += (s, e) =>
			{
				_state._justCompletedFunc = false;
				_state._justCompletedAutoBrace = false;

				if(_state._justCompletedWordVer != null && e.AfterVersion != _state._justCompletedWordVer)
					_state._justCompletedFunc = true;

				_state._justCompletedWordVer = null;
			};
		}

		public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
		{
			if(VsShellUtilities.IsInAutomationFunction(Service))
				return _nextCommandTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
			
			if(pguidCmdGroup == VSConstants.VSStd2K)
			{
				switch((VSConstants.VSStd2KCmdID) nCmdID)
				{
					case VSConstants.VSStd2KCmdID.TAB:
						if(_state._justCompletedFunc)
						{
							var caret_pos = TextView.Caret.Position.BufferPosition;
							int caret = caret_pos.Position;

							//IntPtr pointer = Marshal.AllocHGlobal(1024);
							//Marshal.GetNativeVariantForObject('(', pointer);

							/*object customin = '(';
							object customout = null;
							DTE.Commands.Raise(VSConstants.VSStd2K.ToString("B").ToUpper(), (int)VSConstants.VSStd2KCmdID.TYPECHAR, ref customin, ref customout);*/

							TextView.TextBuffer.Insert(caret, "()");
							TextView.Caret.MoveToPreviousCaretPosition();
							_state._justCompletedFunc = false;
							_state._justCompletedAutoBrace = true;

							caret_pos = TextView.Caret.Position.BufferPosition;
							_state._justCompletedAutoBraceTrack = TextView.TextBuffer.CurrentSnapshot.CreateTrackingSpan(
								Span.FromBounds(caret + 1, caret_pos.GetContainingLine().End.Position), SpanTrackingMode.EdgeInclusive
							);

							TabAutoCallPackage.MenuService.GlobalInvoke(new CommandID(VSConstants.VSStd2K, (int) VSConstants.VSStd2KCmdID.PARAMINFO));
							return VSConstants.S_OK;
						}
						else if(_state._justCompletedAutoBrace)
						{
							_state._justCompletedAutoBrace = false;
							_state._justCompletedAutoBraceTrack = null;
							Operations.MoveToEndOfLine(false);
							Operations.InsertText(";");
							return VSConstants.S_OK;
						}
						break;
				}
			}

			return _nextCommandTarget.Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
		}

		public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
		{
			return _nextCommandTarget.QueryStatus(pguidCmdGroup, cCmds, prgCmds, pCmdText);
		}
	}
}