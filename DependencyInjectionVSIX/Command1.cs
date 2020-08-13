using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace DependencyInjectionVSIX
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command1
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("3680491b-cf18-4e34-92b5-266895e0a1e7");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command1(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command1 Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new Command1(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var dte = this.ServiceProvider.GetServiceAsync(typeof(DTE)).ConfigureAwait(true).GetAwaiter().GetResult() as DTE2;
                var doc = dte.ActiveDocument;
                var name = doc.FullName;
                var language = doc.Language;

                TextDocument textDoc = (TextDocument)doc.Object("");

                var classCode = (CodeClass)textDoc.Selection.ActivePoint.CodeElement[vsCMElement.vsCMElementClass];
                string classname = classCode.Name;

                foreach (CodeElement member in classCode.Members)
                {
                    if (member.Kind == vsCMElement.vsCMElementFunction && member.Name == classCode.Name)
                    {
                        //构造函数
                        CodeFunction func = (CodeFunction)member;
                        if (func.FunctionKind == vsCMFunction.vsCMFunctionConstructor &&
                            func.Access == vsCMAccess.vsCMAccessPublic && func.Parameters.Count > 0)
                        {
                            handleFunction(func, classCode);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(
             this.package,
             ex.Message,
             "Error",
             OLEMSGICON.OLEMSGICON_WARNING,
             OLEMSGBUTTON.OLEMSGBUTTON_OK,
             OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }          
        }

        bool hasField(CodeClass classCode,string fieldName,CodeTypeRef fieldType)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (CodeElement member in classCode.Members)
            {
                if (member.Kind == vsCMElement.vsCMElementVariable )
                {
                    CodeVariable field = (CodeVariable)member;
                    if(field.Type == fieldType)
                    {
                        if (field.Name.ToLower() == "_" + fieldName.ToLower())
                            return true;
                    }
                }
            }
            return false;
        }

        void handleFunction(CodeFunction func,CodeClass classCode)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            List<CodeParameter> parameters = new List<CodeParameter>();
            foreach ( CodeParameter p in func.Parameters )
            {
                if( hasField(classCode,p.Name , p.Type) == false)
                {
                    parameters.Add(p);
                   
                }
            }

            foreach( var p in parameters )
            {
                var fieldName = "_" + p.Name.Substring(0, 1).ToLower() + (p.Name.Length > 1 ? p.Name.Substring(1) : "");
                //添加依赖注入
                var point = func.EndPoint.CreateEditPoint();

                point.CharLeft(1);
                point.Insert($"this.{fieldName} = {p.Name};\r\n");

                point = func.StartPoint.CreateEditPoint();
                point.Insert($"{p.Type.CodeType.Name} {fieldName};\r\n");
            }

            classCode.StartPoint.CreateEditPoint().SmartFormat(func.EndPoint);
        }
    }
}
