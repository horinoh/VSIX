using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using Microsoft.VisualStudio.VCCodeModel;
using System.Text;
using System.Diagnostics;

namespace VSIX
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("c52dcbf0-2c7f-4241-96e4-e36e68f969be");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Command(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new Command(package);
        }

        private void ProcessNamespace(EnvDTE.CodeElements namespaces)
        {
            foreach(var i in namespaces)
            {
                var codeNamespace = (VCCodeNamespace)i;
                if(null != codeNamespace)
                {
                    var sb = new StringBuilder(codeNamespace.Name);
                    sb.Append("::");
                    var prefix = sb.ToString();

                    ProcessClass(prefix, codeNamespace.Classes);

                    ProcessFunction(prefix, codeNamespace.Functions);
                }
            }
        }
        private void ProcessClass(string prefix, EnvDTE.CodeElements classes)
        {
            foreach (var i in classes)
            {
                var codeClass = (VCCodeClass)i;
                if (null != codeClass)
                {
                    var sb = new StringBuilder(prefix);
                    sb.Append(codeClass.Name);
                    sb.Append("::");
                    var newPrefix = sb.ToString();

                    ProcessClass(newPrefix, codeClass.Classes);

                    ProcessFunction(newPrefix, codeClass.Functions);
                }
            }
        }
        private void ProcessFunction(string prefix, EnvDTE.CodeElements functions)
        {
            foreach (var i in functions)
            {
                var codeFunction = (VCCodeFunction)i;
                if (null != codeFunction)
                {
                    Debug.WriteLine("\t" + prefix + codeFunction.Name);

                    #region EDIT_TEST
                    if (false/*prefix + codeFunction.Name == ""*/)
                    {
#if false
                        var textPoint = codeFunction.GetStartPoint(EnvDTE.vsCMPart.vsCMPartBody);
                        if (null != textPoint)
                        {
                            var editPoint = textPoint.CreateEditPoint();
                            if (null != editPoint)
                            {
                                //!< "XXX" が見つかると editPoint は "XXX" の先頭、endPoint は末尾になる
                                EnvDTE.EditPoint endPoint = null;
                                if (editPoint.FindPattern("XXX", 0, ref endPoint))
                                {
                                    editPoint.Delete(endPoint);
                                    editPoint.DeleteWhitespace(EnvDTE.vsWhitespaceOptions.vsWhitespaceOptionsVertical);
                                }
                                else
                                {
                                    editPoint.Insert("\n");
                                    editPoint.Indent();
                                    editPoint.Insert("XXX\n");
                                    editPoint.Indent();
                                }
                            }
                        }
#else
                        //!< GetStartPoint() に(デフォルト以外の)引数を渡すと例外を発生することあったので自前でやる…
                        var textPoint = codeFunction.GetStartPoint();
                        if(null != textPoint)
                        {
                            var editPoint = textPoint.CreateEditPoint();
                            if(null != editPoint)
                            {
                                //!< "{" が見つかると editPoint は "{" の先頭になる
                                if (editPoint.FindPattern("{"))
                                {
                                    //!< "{" の次へ (これで vsCMPartBody 相当になる)
                                    editPoint.CharRight();

                                    //!< "XXX" が見つかると editPoint は "XXX" の先頭、endPoint は末尾になる
                                    EnvDTE.EditPoint endPoint = null;
                                    if (editPoint.FindPattern("XXX", 0, ref endPoint))
                                    {
                                        editPoint.Delete(endPoint);
                                        editPoint.DeleteWhitespace(EnvDTE.vsWhitespaceOptions.vsWhitespaceOptionsVertical);
                                    }
                                    else
                                    {
                                        editPoint.Insert("\n");
                                        editPoint.Indent();
                                        editPoint.Insert("XXX\n");
                                        editPoint.Indent();
                                    }
                                }
                            }
                        }
#endif
                    }
#endregion

                    foreach (var j in codeFunction.Parameters)
                    {
                        var codeParameter = (VCCodeParameter)j;
                        if (null != codeParameter)
                        {
                            Debug.WriteLine("\t" + "\t" + codeParameter.TypeString + " " + codeParameter.Name);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var dte = (EnvDTE.DTE)ServiceProvider.GetService(typeof(EnvDTE.DTE));
            if(null != dte)
            {
                var solution = dte.Solution;
                if(null != solution)
                {
                    for (var i = 1; i <= solution.Count; ++i)
                    {
                        Debug.WriteLine(solution.Item(i).Name);

                        var codeModel = (VCCodeModel)solution.Item(i).CodeModel;
                        if (null != codeModel)
                        {
                            ProcessNamespace(codeModel.Namespaces);

                            ProcessClass("", codeModel.Classes);

                            ProcessFunction("", codeModel.Functions);
                        }
                    }
                }
            }

            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "Command";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
