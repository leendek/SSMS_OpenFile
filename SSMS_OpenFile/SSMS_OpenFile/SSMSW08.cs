/*************************************************************************************
Copyright (c) 2011, Donald Halloran (allmhuran@gmail.com)
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are 
permitted provided that the following two conditions are met:
1) Redistributions must retain the above copyright notice and this list of conditions.
2) The name of the copyright holder may not be used to endorse or promote products 
   derived from this software without specific prior written permission.
 
**************************************************************************************

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY 
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL 
THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, 
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, 
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR 
BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY
WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH 
DAMAGE.

*************************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using SMO = Microsoft.SqlServer.Management.Smo;
using VSI = Microsoft.SqlServer.Management.UI.VSIntegration;
using BARS = Microsoft.VisualStudio.CommandBars;
using System.Configuration;

namespace SSMSW08
{

    public class connection
    {
        private SMO.RegSvrEnum.UIConnectionInfo _connection;
        internal connection set_connection(SMO.RegSvrEnum.UIConnectionInfo connection)
        {
            _connection = connection;
            return this;
        }
        public string user_name
        {
            get
            {
                return _connection.UserName;
            }
        }
        public string instance_name
        {
            get
            {
                return _connection.ServerName;
            }
        }
        public string database_name
        {
            get
            {
                return _connection.AdvancedOptions["DATABSE"];
            }
        }
    }

    public class query_window
    {
        private Document _doc;
        private TextDocument _tdoc;
        private connection _connection;
        internal query_window()
        {
            _connection = new connection();
        }
        internal query_window set_doc(Document doc)
        {
            _doc = doc;
            _tdoc = (TextDocument)_doc.Object("TextDocument");
            return this;
        }
        public string text
        {
            get
            {
                EditPoint p = _tdoc.StartPoint.CreateEditPoint();
                return p.GetText(_tdoc.EndPoint);
            }
            set
            {
                EditPoint p1 = _tdoc.StartPoint.CreateEditPoint();
                EditPoint p2 = _tdoc.EndPoint.CreateEditPoint();
                p1.Delete(p2);
                p1.Insert(value);
            }
        }
        public string current_query
        {
            get
            {
                if (_tdoc.Selection != null && _tdoc.Selection.Text != "")
                {
                    return _tdoc.Selection.Text;
                }
                else
                {
                    return this.text;
                }
            }
        }
        public string name
        {
            get
            {
                return _doc.Name;
            }
        }
        public string file_path
        {
            get
            {
                return _doc.FullName;
            }
        }
        public connection connection
        {
            get
            {
                Document old_active = null;
                if (_doc.DTE.ActiveDocument.FullName != _doc.FullName)
                {
                    old_active = _doc.DTE.ActiveDocument;
                    _doc.Activate();
                }
                _connection.set_connection(VSI.ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo);
                if (old_active != null)
                {
                    old_active.Activate();
                }
                return _connection;
            }
        }
    }

    public class window_manager
    {
        private SSMS _ssms;
        private Windows2 _SSMS_window_collection;
        private query_window _working_query_window;
        private System.Collections.Generic.List<Window> _addin_window_collection;
        public query_window query_windows(int index)
        {
            return _ssms.DTE2.Documents.Item(index) == null ? null : _working_query_window.set_doc(_ssms.DTE2.ActiveDocument);
        }

        public query_window active_query_window
        {
            get
            {
                return _ssms.DTE2.ActiveDocument == null ? null : _working_query_window.set_doc(_ssms.DTE2.ActiveDocument);
            }
        }
        public Window addin_tool_windows(int index)
        {
            return _addin_window_collection[index];
        }

        public query_window create_query_window()
        {
            object new_query_window = VSI.ServiceCache.ScriptFactory.CreateNewBlankScript(VSI.Editors.ScriptType.Sql);
            return active_query_window;
        }

        public query_window create_query_window(string server_name)
        {
            Microsoft.SqlServer.Management.Smo.RegSvrEnum.UIConnectionInfo CI = new Microsoft.SqlServer.Management.Smo.RegSvrEnum.UIConnectionInfo();
            CI.ServerName = server_name;
            object new_query_window = VSI.ServiceCache.ScriptFactory.CreateNewBlankScript(VSI.Editors.ScriptType.Sql, CI, null);
            return active_query_window;
        }
        public Window create_tool_window(string class_name, string caption, string GUID, ref object control)
        {
            Window new_window = null;
            // new_window = ((Windows2)VSI.ServiceCache.ExtensibilityModel.Windows).CreateToolWindow2(
            new_window = ((Windows2)_ssms.DTE2.Windows).CreateToolWindow2(
               _ssms.addin,
               _ssms.addin_location,
               class_name,
               caption,
               GUID,
               ref control
            );
            new_window.SetKind(vsWindowType.vsWindowTypeDesigner);
            new_window.Visible = true;
            _addin_window_collection.Add(new_window);
            return new_window;
        }
        internal window_manager(SSMS ssms)
        {
            _ssms = ssms;
            _working_query_window = new query_window();
            _SSMS_window_collection = (Windows2)_ssms.DTE2.Windows;
            _addin_window_collection = new System.Collections.Generic.List<Window>();
        }
    }

    public class event_manager
    {
        private SSMS _ssms;
        private System.Collections.Generic.List<CommandEvents> _registered_events;
        public struct SSMS_event
        {
            public string guid;
            public int id;
            public SSMS_event(string guid, int id)
            {
                this.guid = guid;
                this.id = id;
            }
        }
        public static class query
        {
            public static readonly SSMS_event execute = new SSMS_event("{52692960-56BC-4989-B5D3-94C47A513E8D}", 1);
        }
        internal event_manager(SSMS ssms)
        {
            _ssms = ssms;
            _registered_events = new System.Collections.Generic.List<CommandEvents>();
        }
        public void register_after_event(SSMS_event event_to_register, _dispCommandEvents_AfterExecuteEventHandler handler)
        {
            CommandEvents ev = _ssms.DTE.Events.get_CommandEvents(event_to_register.guid, event_to_register.id);
            ev.AfterExecute += new _dispCommandEvents_AfterExecuteEventHandler(handler);
            _registered_events.Add(ev);
        }
        public void register_before_event(SSMS_event event_to_register, _dispCommandEvents_BeforeExecuteEventHandler handler)
        {
            CommandEvents ev = _ssms.DTE.Events.get_CommandEvents(event_to_register.guid, event_to_register.id);
            ev.BeforeExecute += new _dispCommandEvents_BeforeExecuteEventHandler(handler);
            _registered_events.Add(ev);
        }
    }

    public class command_manager
    {
        public delegate void menu_command_querystatus_handler(vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText);
        public delegate void menu_command_exec_handler(vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled);
        private struct menu_command_handlers
        {
            public menu_command_querystatus_handler querystatus_handler;
            public menu_command_exec_handler exec_handler;
        }
        private System.Collections.Generic.List<BARS.CommandBar> _addin_command_bars;
        private Commands2 _SSMS_commands_collection;
        private System.Collections.Generic.Dictionary<string, menu_command_handlers> _addin_menu_commands_dictonary;
        private SSMS _ssms;

        internal command_manager(SSMS ssms)
        {
            // VSI.ServiceCache.ExtensibilityModel.Commands was NULL when called.
            // _SSMS_commands_collection = (Commands2)VSI.ServiceCache.ExtensibilityModel.Commands;
            _SSMS_commands_collection = (Commands2)ssms.addin.DTE.Commands;
            _addin_menu_commands_dictonary = new System.Collections.Generic.Dictionary<string, menu_command_handlers>();
            _ssms = ssms;
            _addin_command_bars = new System.Collections.Generic.List<BARS.CommandBar>();
        }

        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (_addin_menu_commands_dictonary.ContainsKey(commandName))
                {
                    _addin_menu_commands_dictonary[commandName].querystatus_handler(neededText, ref status, ref commandText);
                }
            }
        }

        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = false;
            if (_addin_menu_commands_dictonary.ContainsKey(commandName))
            {
                _addin_menu_commands_dictonary[commandName].exec_handler(executeOption, ref varIn, ref varOut, ref handled);
            }
        }

        private void _default_querystatus_handler(vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
        }

        private void _default_exec_handler(vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = true;
        }

        public void create_popup_menu(string host_menu_bar_name, string new_bar_name, string tooltip_text, int position)
        {
            BARS.CommandBar new_bar = null;
            // check existing
            try
            {
                // VSI.ServiceCache.ExtensibilityModel.Commands was NULL
                // new_bar = (BARS.CommandBar)((BARS.CommandBars)VSI.ServiceCache.ExtensibilityModel.CommandBars)[new_bar_name];
                new_bar =
                    (BARS.CommandBar)((BARS.CommandBars)_ssms.DTE2.CommandBars)[new_bar_name];
                new_bar.Delete();
                new_bar = null;
            }
            catch
            {
            }
            // VSI.ServiceCache.ExtensibilityModel.Commands was NULL
            // BARS.CommandBar host_bar = ((BARS.CommandBars)VSI.ServiceCache.ExtensibilityModel.CommandBars)[host_menu_bar_name];
            BARS.CommandBar host_bar = ((BARS.CommandBars)_ssms.DTE2.CommandBars)[host_menu_bar_name];
            position = (position == 0 ? host_bar.Controls.Count + 1 : position);
            new_bar = (BARS.CommandBar)_SSMS_commands_collection.AddCommandBar(new_bar_name, vsCommandBarType.vsCommandBarTypeMenu, host_bar, position);
            _addin_command_bars.Add(new_bar);
        }

        public void log_command_bar_names()
        {
            var path = @"C:\kiln\commandbars.txt";
            using (var sw = File.CreateText(path))
            {
                foreach (CommandBar commandBar in (BARS.CommandBars)_ssms.DTE2.CommandBars)
                {
                    sw.WriteLine(commandBar.Name);
                    sw.WriteLine("-------------");
                    List<string> names = (from CommandBarControl commandBarControl in commandBar.Controls select commandBarControl.accName).ToList();
                    sw.WriteLine(names.Aggregate("", (current, name) => current + ("\n" + name)));
                    sw.WriteLine();
                }
            }


        }

        public void create_popup_menu_command(
           string host_menu_bar_name,
           string command_name,
           string item_text,
           string tooltip_text,
           int position,
           menu_command_querystatus_handler querystatus_handler,
           menu_command_exec_handler exec_handler
        )
        {
            object[] contextGUIDS = new object[] { };
            BARS.CommandBar host_bar = ((BARS.CommandBars)_ssms.DTE2.CommandBars)[host_menu_bar_name];

            Command new_command = _SSMS_commands_collection.AddNamedCommand2(_ssms.addin, command_name, item_text, tooltip_text, true, 0, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStyleText, vsCommandControlType.vsCommandControlTypeButton);
            if (new_command == null)
            {
                throw new Exception("SSMSW08 could not create command bar: " + item_text);
            }
            position = (position == 0 ? host_bar.Controls.Count + 1 : position);
            new_command.AddControl(host_bar, position);
            menu_command_handlers handlers = new menu_command_handlers();
            handlers.querystatus_handler = querystatus_handler == null ? this._default_querystatus_handler : querystatus_handler;
            handlers.exec_handler = exec_handler == null ? this._default_exec_handler : exec_handler;
            this._addin_menu_commands_dictonary.Add(_ssms.addin.ProgID + "." + command_name, handlers);
        }

       public void create_popup_menu_and_context_menu_command(
       string host_menu_bar_name,
       string command_name,
       string item_text,
       string tooltip_text,
       int position,
       menu_command_querystatus_handler querystatus_handler,
       menu_command_exec_handler exec_handler,
       string keyBindings = null
        )
        {
            object[] contextGUIDS = new object[] { };
            string context_menu_name = "SQL Files Editor Context";

            BARS.CommandBar host_bar = ((BARS.CommandBars)_ssms.DTE2.CommandBars)[host_menu_bar_name];
            BARS.CommandBar context_menu = ((BARS.CommandBars)_ssms.DTE2.CommandBars)[context_menu_name];

            Command new_command = _SSMS_commands_collection.AddNamedCommand2(_ssms.addin, command_name, item_text, tooltip_text, true, 0, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStyleText, vsCommandControlType.vsCommandControlTypeButton);
            if (new_command == null)
            {
                throw new Exception("SSMSW08 could not create command bar: " + item_text);
            }
            position = (position == 0 ? host_bar.Controls.Count + 1 : position);
            new_command.AddControl(host_bar, position);
            new_command.AddControl(context_menu, context_menu.Controls.Count + 1);
           if(keyBindings!= null)
               new_command.Bindings = keyBindings;
            menu_command_handlers handlers = new menu_command_handlers();
            handlers.querystatus_handler = querystatus_handler == null ? this._default_querystatus_handler : querystatus_handler;
            handlers.exec_handler = exec_handler == null ? this._default_exec_handler : exec_handler;
            this._addin_menu_commands_dictonary.Add(_ssms.addin.ProgID + "." + command_name, handlers);
        }


        public void cleanup()
        {
            // foreach (Command c in VSI.ServiceCache.ExtensibilityModel.Commands)
            foreach (Command c in _ssms.DTE2.Commands)
            {
                if (_addin_menu_commands_dictonary.ContainsKey(c.Name))
                {
                    c.Delete();
                }
            }
            foreach (BARS.CommandBar bar in _addin_command_bars)
            {
                bar.Delete();
            }
        }
    }

    public class SSMS
    {
        private EnvDTE80.DTE2 _DTE2;
        private EnvDTE.DTE _DTE;
        private AddIn _addin;
        private command_manager _command_manager;
        private window_manager _window_manager;
        private event_manager _event_manager;
        public string addin_location;
        private System.Configuration.Configuration _cfg;
        public SSMS(object addin_instance, string assembly_location)
        {
            _addin = (AddIn)addin_instance;
            _DTE2 = (DTE2)_addin.DTE;
            _DTE = (DTE)_addin.DTE;
            addin_location = assembly_location;
            _command_manager = new command_manager(this);
            _window_manager = new window_manager(this);
            _event_manager = new event_manager(this);
        }
        public bool load_config()
        {
            bool found = false;
            try
            {
                _cfg = System.Configuration.ConfigurationManager.OpenExeConfiguration(addin_location);
                found = true;
            }
            catch
            {
            }
            return found;
        }
        public System.Configuration.Configuration config
        {
            get
            {
                return _cfg;
            }
        }
        public EnvDTE80.DTE2 DTE2
        {
            get
            {
                return _DTE2;
            }
        }
        public EnvDTE.DTE DTE
        {
            get
            {
                return _DTE;
            }
        }
        public AddIn addin
        {
            get
            {
                return _addin;
            }
        }
        public command_manager command_manager
        {
            get
            {
                return _command_manager;
            }
        }
        public window_manager window_manager
        {
            get
            {
                return _window_manager;
            }
        }
        public event_manager event_manager
        {
            get
            {
                return _event_manager;
            }
        }
        public void OnDisconnection()
        {
            _command_manager.cleanup();
        }
    }
}