using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using System.Diagnostics;
using Microsoft.VisualStudio.VCCodeModel;

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
                    #region Solution
                    //!< ソリューション列挙
                    for (var i = 1; i <= solution.Count; ++i)
                    {
                        Debug.WriteLine(solution.Item(i).Name);

                        var codeModel = (VCCodeModel)solution.Item(i).CodeModel;
                        if (null != codeModel)
                        {
                            var classes = codeModel.Classes;
                            if(null != classes)
                            {
                                #region Class
                                //!< クラス列挙
                                foreach(var j in classes)
                                {
                                    var codeClass = (VCCodeClass)j;
                                    if (null != codeClass)
                                    {
                                        Debug.WriteLine("\t" + codeClass.Name);

                                        var functions = codeClass.Functions;
                                        if (null != functions)
                                        {
                                            #region Function
                                            //!< クラス関数列挙
                                            foreach(var k in functions)
                                            {
                                                var codeFunction = (VCCodeFunction)k;
                                                if (null != codeFunction)
                                                {
                                                    Debug.WriteLine("\t" + "\t" + codeFunction.Name);
                                                    //!< 引数列挙
                                                    foreach(var l in codeFunction.Parameters)
                                                    {
                                                        var codeParameter = (VCCodeParameter)l;
                                                        if(null != codeParameter)
                                                        {
                                                            Debug.WriteLine("\t" + "\t" + "\t" + codeParameter.TypeString + " " + codeParameter.Name);
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                }
                                #endregion
                            }

                            var globalFunctions = codeModel.Functions;
                            if (null != globalFunctions)
                            {
                                #region GlobalFunction
                                //!< グローバル関数列挙
                                foreach(var j in globalFunctions)
                                {
                                    var codeFunction = (VCCodeFunction)j;
                                    if (null != codeFunction)
                                    {
                                        Debug.WriteLine("\t" + codeFunction.Name);

#if false
                                        //!< 関数名が "XXX" の場合、関数ボディの先頭に "YYY" を挿入するテスト
                                        if (codeFunction.Name == "XXX")
                                        {
                                            var textPoint = codeFunction.GetStartPoint(EnvDTE.vsCMPart.vsCMPartBody);
                                            if (null != textPoint)
                                            {
                                                var editPoint = textPoint.CreateEditPoint();
                                                editPoint.WordRight();
                                                editPoint.Insert("YYY");
                                            }
                                        }
#endif
                                    }
                                }
                                #endregion
                            }
                        }
                    }
                    #endregion
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
