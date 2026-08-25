using System;
using System.Windows;
using System.Windows.Threading;

namespace SpaceReportApp
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Without these, any unhandled exception closes the window instantly
            // with nothing on screen and nothing in the console - the app simply
            // disappears. Always show what went wrong instead.
            DispatcherUnhandledException += (s, a) =>
            {
                Report(a.Exception, "UI thread");
                a.Handled = true;          // stay alive; the user can retry
            };

            AppDomain.CurrentDomain.UnhandledException += (s, a) =>
                Report(a.ExceptionObject as Exception, "background thread");

            System.Threading.Tasks.TaskScheduler.UnobservedTaskException += (s, a) =>
            {
                Report(a.Exception, "background task");
                a.SetObserved();
            };
        }

        private static void Report(Exception ex, string where)
        {
            if (ex == null) return;

            // Inner exceptions carry the useful part. A XamlParseException on its
            // own only says "a type converter threw"; the inner one names the
            // property and the file that could not be found.
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Something went wrong on the " + where + ".");
            for (var e = ex; e != null; e = e.InnerException)
            {
                sb.AppendLine();
                sb.AppendLine(e.GetType().Name + ": " + e.Message);
            }
            sb.AppendLine();
            sb.AppendLine(ex.StackTrace);
            string text = sb.ToString();

            string log = null;
            try
            {
                string dir = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "SpaceReport");
                System.IO.Directory.CreateDirectory(dir);
                log = System.IO.Path.Combine(dir, "error.log");
                System.IO.File.AppendAllText(log,
                    "=== " + DateTime.Now.ToString("s") + " ===" + Environment.NewLine +
                    text + Environment.NewLine);
            }
            catch { log = null; }

            try
            {
                MessageBox.Show(text + (log == null ? "" : "\n\nWritten to: " + log),
                    "Space Report", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch { }
        }
    }
}
