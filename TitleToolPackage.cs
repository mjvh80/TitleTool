using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using Task = System.Threading.Tasks.Task;

namespace TitleTool
{
    internal class Ref<T> where T : struct
    {
        public T Value;
    }

    internal static class Extensions
    {
        public static T Update<T>(this T t, Action<Ref<T>> updater) where T : struct
        {
            var @ref = new Ref<T> { Value = t };
            updater(@ref);
            return @ref.Value;
        }
    }

    //  [ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuidString)]
    public sealed class TitleToolPackage : AsyncPackage
    {
        public const string PackageGuidString = "88ecdc8e-b5e2-4a3e-b067-5e5ce80bba82";

        /// <summary>
        /// Find a descendant control of given type or name. If both type and name are given both must match.
        /// </summary>
        private static IEnumerable<DependencyObject> FindChildDepthFirst(DependencyObject parent, string fullTypeName, string name)
        {
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (fullTypeName == null && name == null) throw new ArgumentException("need at least one type or name to search on");

            var count = VisualTreeHelper.GetChildrenCount(parent);
            for (var i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (fullTypeName != null && !child.GetType().FullName.Equals(fullTypeName, StringComparison.OrdinalIgnoreCase))
                    goto VisitChildren;

                if (name == null)
                    // have match on type so return
                    yield return child;
                else if (child is FrameworkElement f && f.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    yield return child;

                VisitChildren:
                if (VisualTreeHelper.GetChildrenCount(child) > 0)
                {
                    foreach (var grandChild in FindChildDepthFirst(child, fullTypeName, name))
                        yield return grandChild;
                }
            }
        }

        private ToolBar FindToolBarByName(IEnumerable<ToolBar> allToolBars, string name)
            => allToolBars.FirstOrDefault(tb =>
                tb.GetValue(AutomationProperties.NameProperty)?.ToString()?.Equals(name, StringComparison.OrdinalIgnoreCase) ?? false);

        private void ReportProgress(IProgress<ServiceProgressData> progress, string stepMessage, int currentStep) =>
            progress.Report(new ServiceProgressData("Loading TitleTool", stepMessage, currentStep, 2));

        private enum LogType { Info, Error }

        private IVsActivityLog _mLog;
        private void Log(LogType type, string msg)
        {
            if (_mLog == null)
            {
                Trace.WriteLine("No logger found: " + msg);
                return;
            }

            Trace.WriteLine(type + ": " + msg);

            ThreadHelper.ThrowIfNotOnUIThread();

            var entryType = type == LogType.Error ?
                __ACTIVITYLOG_ENTRYTYPE.ALE_ERROR : __ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION;

            int hr = _mLog.LogEntry((UInt32)entryType, "TitleToolExtension", msg);
            if (hr != 0)
            {
                Trace.WriteLine($"Non-zero HRESULT logging to ActivityLog: {hr} {msg}");
            }
        }

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            Trace.WriteLine("Initializing");

            ReportProgress(progress, "Initializing", 0);

            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _mLog = await GetServiceAsync(typeof(SVsActivityLog)) as IVsActivityLog;

            cancellationToken.ThrowIfCancellationRequested();
            ReportProgress(progress, "On main thread", 1);

            Trace.WriteLine("On main thread");

            Action<string> moveAction = (string type) =>
            {
                Trace.WriteLine($"MainWindow loaded {type} - moving toolbar");
                var ok = MoveToolbarToTitleBar();
                ReportProgress(progress, $"{type} Completion", 2);
                if (ok) Log(LogType.Info, "Moved toolbar to titlebar");
            };

            if (Application.Current.MainWindow.IsLoaded)
                moveAction("Direct");
            else
                Application.Current.MainWindow.Loaded += (sender, args) => moveAction("Deferred");
        }

        private enum TitleBarType
        {
            Main,
            FullScreen
        }

        private bool _mMovedToolbarOrImpossible = false;

        private bool MoveToolbarToTitleBar(bool hookLayoutUpdate = true)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_mMovedToolbarOrImpossible) return true;

            var mainWindow = Application.Current.MainWindow;

            var dockTray = (ToolBarTray)FindChildDepthFirst(mainWindow, null, "TopDockTray").FirstOrDefault();
            if (dockTray == null)
            {
                Log(LogType.Error, "Could not find dock tray called 'TopDockTray'");
                return (_mMovedToolbarOrImpossible = true);
            }

            var toolbars = FindChildDepthFirst(dockTray, "Microsoft.VisualStudio.PlatformUI.VsToolBar", null).Cast<ToolBar>();

            var standardToolBar = FindToolBarByName(toolbars, "TitleBar"); // "Standard");

            if (standardToolBar == null)
            {
                Trace.WriteLine("Toolbar not found - hooking tray layout update");

                // Hook some events to retry if a toolbar has been enabled later.
                if (hookLayoutUpdate)
                {
                    EventHandler layoutUpdateHandler = null;
                    layoutUpdateHandler = (sender, args) =>
                    {
                        Trace.WriteLine("Tray layout updated");
                        if (MoveToolbarToTitleBar(false))
                            dockTray.LayoutUpdated -= layoutUpdateHandler;
                    };
                    dockTray.LayoutUpdated += layoutUpdateHandler;

                    dockTray.Loaded += (sender, args) =>
                    {
                        foreach (var tb in dockTray.ToolBars)
                        {
                            DependencyPropertyChangedEventHandler visibleHandler = null;
                            visibleHandler = (_, __) =>
                            {
                                if (MoveToolbarToTitleBar(false))
                                    tb.IsVisibleChanged -= visibleHandler;
                            };
                            tb.IsVisibleChanged += visibleHandler;
                        }
                    };
                }

                return false;
            }

            Trace.WriteLine("Found toolbar");

            var mainWindowTitleBar = (FrameworkElement)FindChildDepthFirst(mainWindow, null, "MainWindowTitleBar").FirstOrDefault();
            if (mainWindowTitleBar == null)
            {
                Log(LogType.Error, "Could not find main window titlebar called 'MainWindowTitleBar'");
                return (_mMovedToolbarOrImpossible = true);
            }
            Trace.WriteLine("Found main window toolbar");

            var fullScreenMenuBar = (FrameworkElement)FindChildDepthFirst(mainWindow, null, "PART_MainMenuBar").FirstOrDefault();
            if (fullScreenMenuBar == null)
            {
                Log(LogType.Error, "FullScreen control PART_MainMenuBar not found: toolbar not supported in fullscreen");
            }

            var stackPanel = new StackPanel();
            stackPanel.Orientation = Orientation.Horizontal;
            stackPanel.Margin = new Thickness() { Top = 4 };

            var secondTray = (ToolBarTray)Activator.CreateInstance(dockTray.GetType());
            secondTray.Background = dockTray.Background;

            var border = new Border();
            border.BorderThickness = new Thickness(0, 0, 0, 0);
            border.Child = secondTray;

            dockTray.ToolBars.Remove(standardToolBar);
            secondTray.ToolBars.Add(standardToolBar);

            // Apply "fixups"
            standardToolBar.BorderThickness = new Thickness(0, 0, 0, 0);

            // Hide thumbs as we can't move this one.
            var thumb = FindChildDepthFirst(standardToolBar, null, "ToolBarThumb").FirstOrDefault() as Thumb;
            if (thumb != null)
                thumb.Visibility = Visibility.Collapsed;

            // Copy background, the overflow button needs this for mouseover.
            stackPanel.Children.Add(border);

            Func<TitleBarType> getTitleBarType = () => mainWindowTitleBar.IsVisible ? TitleBarType.Main : TitleBarType.FullScreen;

            Panel currentParent = null; // closure
            Action insertToolBar = () =>
            {
                Trace.WriteLine("Inserting toolbar into title grid for type " + getTitleBarType());

                if (currentParent != null)
                {
                    currentParent.Children.Remove(stackPanel);
                    currentParent.UpdateLayout();
                }

                if (getTitleBarType() == TitleBarType.Main)
                {
                    stackPanel.Margin = stackPanel.Margin.Update(m => m.Value.Top = 4);

                    Grid.SetRow(stackPanel, 0);
                    Grid.SetColumn(stackPanel, 4);

                    currentParent = ((Grid)VisualTreeHelper.GetChild(((Grid)VisualTreeHelper.GetChild(mainWindowTitleBar, 0)), 2));
                    ((Grid)currentParent).Children.Add(stackPanel);
                }
                else if (fullScreenMenuBar != null)
                {
                    // We've gone fullscreen.
                    stackPanel.Margin = stackPanel.Margin.Update(m => m.Value.Top = 0);

                    var dockPanel = (DockPanel)fullScreenMenuBar.Parent;
                    dockPanel.Children.Add(stackPanel);

                    // for some reason it becomes invisible after move
                    standardToolBar.Visibility = Visibility.Visible;

                    dockPanel.UpdateLayout(); // todo: check if these calls are in fact necessary
                    currentParent = dockPanel;
                }
            };

            insertToolBar();

            // When switching from fullscreen to normal or vice versa:
            mainWindowTitleBar.IsVisibleChanged += (_, __) => insertToolBar();

            // Adjust layout if other things are added such as the solution info control.
            mainWindowTitleBar.LayoutUpdated += (_, __) =>
            {
                if (getTitleBarType() == TitleBarType.FullScreen)
                    return;

                var solutionInfoControl = FindChildDepthFirst(mainWindowTitleBar, "Microsoft.VisualStudio.PlatformUI.SolutionInfoControl", null).FirstOrDefault();
                if (solutionInfoControl == null) return;

                var textBox = (FrameworkElement)FindChildDepthFirst(solutionInfoControl, null, "TextBorder").FirstOrDefault();
                if (textBox == null)
                {
                    Trace.WriteLine("Could not find TextBorder in SolutionInfoControl");
                    return;
                }

                stackPanel.Margin = stackPanel.Margin.Update(m => m.Value.Left = textBox.ActualWidth + 10);
            };

            return (_mMovedToolbarOrImpossible = true);
        }
    }
}
