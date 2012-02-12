using System;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using EnvDTE;
using Process = System.Diagnostics.Process;
using Window = EnvDTE.Window;

namespace Dotty
{
    public partial class ViewGraphControl : UserControl
    {
        private DTE applicationObject;
        private DispatcherTimer timer;
        private Window parentWindow;
        private string lastGraph;

        private static ViewGraphControl instance; // Fulkoda!

        public ViewGraphControl()
        {
            instance = this;
            lastGraph = "";
            InitializeComponent();
        }

        public static void Initialize(Window parentWindow, DTE applicationObject)
        {
            // Whole mehod is ugly as hell, and should not exist. This exists only because 
            // VS integration is being a little bitch to me, and I'm retarded enough to
            // not figure out how it works.
            instance.InitializeInstance(parentWindow, applicationObject);
        }

        public void InitializeInstance(Window parentWindow, DTE applicationObject)
        {
            this.applicationObject = applicationObject;
            this.parentWindow = parentWindow;
            timer = new DispatcherTimer();
            timer.Tick += OnTimer;
            timer.Interval = new TimeSpan(0, 0, 0, 1);
            timer.Start();
        }

        private void OnTimer(object sender, EventArgs eventArgs)
        {
            if (parentWindow.Visible)
            {
                var activeDocument = applicationObject.ActiveDocument;
                if (activeDocument != null)
                {
                    var selection = (TextSelection) activeDocument.Selection;
                    LoadGraphImage(selection.Text);
                }
              //  LoadGraphImage("digraph { a -> b }");
            }
        }

        private void LoadGraphImage(string graphData)
        {
            if (graphData != lastGraph)
            {
                lastGraph = graphData;

                var p = new Process
                            {
                                StartInfo =
                                    {
                                        FileName = @"C:\Program Files (x86)\Graphviz 2.28\bin\dot.exe",
                                        Arguments = "-Tpng",
                                        UseShellExecute = false,
                                        RedirectStandardError = true,
                                        RedirectStandardInput = true,
                                        RedirectStandardOutput = true,
                                        CreateNoWindow = true
                                    }
                            };

                p.Start();

                p.StandardInput.WriteLine(graphData);
                p.StandardInput.Close();

                var src = new BitmapImage();
                src.BeginInit();
                src.StreamSource = p.StandardOutput.BaseStream;
                src.CacheOption = BitmapCacheOption.OnLoad;
                src.EndInit();
                image.Source = src;
            }
        }
    }
}
