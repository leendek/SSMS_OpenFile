using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Extensibility;
using EnvDTE;
using SSMSW08;

namespace SsmsOpenFile
{
    /// <summary>The object for implementing an Add-in.</summary>
    /// <seealso class='IDTExtensibility2' />
    public class Connect : IDTExtensibility2, IDTCommandTarget
    {

        private SSMS _ssms;
        private const string HOST_MENU_BAR_NAME = "SSMS_OpenFile2";
        private const string SUB_MENU_ITEM = "Open File2";
        private const string OPEN_FILE_COMMAND_NAME = "open_file2";
        private const string INVALIDATE_CACHE_COMMAND_NAME = "invalidate_cache";
        private const string TOOLTIP = "Open File from selected text.2";

        /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
        /// <param name='application'>Root object of the host application.</param>
        /// <param name='connectMode'>Describes how the Add-in is being loaded.</param>
        /// <param name='addInInst'>Object representing this Add-in.</param>
        /// <param name="custom"></param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _ssms = new SSMS(addInInst, @"C:\kiln\SSMS_OpenFile\SSMS_OpenFile\bin");

            if (connectMode == ext_ConnectMode.ext_cm_Startup)
            {
                _ssms.command_manager.create_popup_menu("Tools", HOST_MENU_BAR_NAME, "A tooltip", 0);
                _ssms.command_manager.create_popup_menu_and_context_menu_command(HOST_MENU_BAR_NAME, OPEN_FILE_COMMAND_NAME, SUB_MENU_ITEM, TOOLTIP, 0, null, OpenSelectedFileName, "Global::Ctrl+k, Ctrl+o");
                _ssms.command_manager.create_popup_menu_command(HOST_MENU_BAR_NAME, INVALIDATE_CACHE_COMMAND_NAME, INVALIDATE_CACHE_COMMAND_NAME, TOOLTIP, 0, null, InvalidateFileCache);
            }
        }

        /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
        /// <param name='disconnectMode'>Describes how the Add-in is being unloaded.</param>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
        }

        /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />		
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
            //MessageBox.Show("Addin Loaded", "SsmsOpenFile");
        }

        /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
            _ssms.command_manager.cleanup();
        }

        /// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
        /// <param name='commandName'>The name of the command to determine state for.</param>
        /// <param name='neededText'>Text that is needed for the command.</param>
        /// <param name='status'>The state of the command in the user interface.</param>
        /// <param name='commandText'>Text requested by the neededText parameter.</param>
        /// <seealso class='Exec' />
        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            _ssms.command_manager.QueryStatus(commandName, neededText, ref status, ref commandText);
        }

        /// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
        /// <param name='commandName'>The name of the command to execute.</param>
        /// <param name='executeOption'>Describes how the command should be run.</param>
        /// <param name='varIn'>Parameters passed from the caller to the command handler.</param>
        /// <param name='varOut'>Parameters passed from the command handler to the caller.</param>
        /// <param name='handled'>Informs the caller if the command was handled or not.</param>
        /// <seealso class='Exec' />
        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            _ssms.command_manager.Exec(commandName, executeOption, ref varIn, ref varOut, ref handled);
        }

        public void InvalidateFileCache(vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            Helper.Instance().InvalidateFileCache();
        }
        public void OpenSelectedFileName(vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            var doc = _ssms.addin.DTE.ActiveDocument;
            if (doc == null) return;
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            var selectedText = GetSelectedWord(doc);
            LogToFeedbackRegion(string.Format("Attempting to open file: {0}", selectedText));
            var path = Helper.Instance().GetPath(selectedText, doc.Path);

            if (File.Exists(path))
            {
                 _ssms.DTE2.ItemOperations.OpenFile(path);
                 LogToFeedbackRegion(string.Format("Opened file: {0} -- {1}ms", path, stopwatch.ElapsedMilliseconds));
            }
            else
            {
                 LogToFeedbackRegion(string.Format("Failed to open file: {0} -- {1}ms", path ?? selectedText, stopwatch.ElapsedMilliseconds));
            }
        }

        private string GetSelectedWord(Document doc)
        {
            _ssms.DTE2.ExecuteCommand("Edit.SelectCurrentWord");
            var txtSelection = (TextSelection) doc.Selection;
            var selectedText = txtSelection.Text;
            return selectedText;
        }

        private void LogToFeedbackRegion(string feedback)
        {
            EnvDTE.StatusBar statusBar =_ssms.DTE2.StatusBar;
            statusBar.Text = feedback;
        }
    }
}

public class Helper
{
    private static Helper _helper;
    private const string START_PATH = @"C:\svn";
    private string[] _filePathCache;

    public static void PrintSomething(string something)
    {
        MessageBox.Show(something, "HelperClass");
    }

    public static Helper Instance()
    {
        return _helper ?? (_helper = new Helper());
    }

    public void InvalidateFileCache()
    {
        FilePathCache = null;
    }

    private string[] FilePathCache
    {
        get
        {
            return _filePathCache ?? (_filePathCache = Directory.GetFiles(START_PATH, "*.sql",
                                                                          SearchOption.AllDirectories));
        }
        set { _filePathCache = value; }
    }


    public string GetPath(string fileName, string priorityDirectory)
    {
        var filePaths = Directory.GetFiles(priorityDirectory, "*.sql",
                             SearchOption.AllDirectories);

        if (!filePaths.Any(filePath => filePath.Contains(fileName + ".sql")))
        {
            filePaths = FilePathCache;
        }
        return filePaths.FirstOrDefault(filePath => filePath.Contains(fileName + ".sql"));
    }
}



