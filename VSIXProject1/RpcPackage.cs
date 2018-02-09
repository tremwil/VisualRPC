using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

namespace VisualRpc
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(RpcPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class RpcPackage : Package
    {
        private System.Threading.Thread UpdateThread;
        private DiscordRpc.RichPresence rpc;
        private DiscordRpc.EventHandlers handlers;
        private Stopwatch sw;

        private long epoch;

        private const int updateDelayMs = 100;
        private bool close;

        private HashSet<string> knownNames = new HashSet<string>()
        {
            "file_aspx", "file_config", "file_cpp", "file_cs", "file_css", "file_dgsl",
            "file_dll", "file_fs", "file_fxg", "file_glsl", "file_h", "file_html",
            "file_ico", "file_js", "file_json", "file_jsx", "file_m", "file_mtl",
            "file_obj", "file_php", "file_ps1", "file_py", "file_r", "file_rb",
            "file_reg", "file_snk", "file_sql", "file_stl", "file_tif", "file_ts",
            "file_txt", "file_vb", "file_vbs", "file_xaml", "file_xml", "file_zip"
        };

        /// <summary>
        /// ToggleRpcPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "95904d81-d4eb-4f6e-9814-402fc0910044";

        /// <summary>
        /// Initializes a new instance of the <see cref="ToggleRpc"/> class.
        /// </summary>
        public RpcPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.

            close = false;

            UpdateThread = new System.Threading.Thread(RpcUpdate);
            UpdateThread.Start();
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            ToggleRpc.Initialize(this);
            base.Initialize();
        }

        protected override int QueryClose(out bool canClose)
        {
            close = true;
            return base.QueryClose(out canClose);
        }

        private void RpcUpdate()
        {
            handlers = new DiscordRpc.EventHandlers();
            rpc = new DiscordRpc.RichPresence();

            sw = new Stopwatch();
            sw.Start();

            epoch = DateTimeOffset.Now.ToUnixTimeSeconds();

            DiscordRpc.Initialize("410943600705273856", ref handlers, true, null);

            while (!close)
            {
                DiscordRpc.RunCallbacks();

                DTE dte = GetGlobalService(typeof(DTE)) as DTE;

                rpc.details = "";
                if (dte != null && dte.Solution != null && dte.Solution.FullName != "")
                {
                    rpc.details = "Working on " + Path.GetFileNameWithoutExtension(dte.Solution.FileName);
                }

                rpc.state = "";
                rpc.smallImageKey = "";
                try
                {
                    if (dte != null && dte.ActiveDocument != null)
                    {
                        rpc.state = "Editing " + dte.ActiveDocument.Name;
                        string imgKey = "file_" + Path.GetExtension(dte.ActiveDocument.Name).Remove(0, 1);
                        if (knownNames.Contains(imgKey)) { rpc.smallImageKey = imgKey; }
                    }
                }
                catch (Exception)
                {

                }

                rpc.largeImageKey = "logo";
                rpc.startTimestamp = epoch;

                DiscordRpc.UpdatePresence(ref rpc);
                System.Threading.Thread.Sleep(updateDelayMs);
            }

            DiscordRpc.Shutdown();
        }

        #endregion
    }
}
