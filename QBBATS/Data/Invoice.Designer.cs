﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.18408
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------
//------------------------------------------------------------------------------

namespace QBBATS.Data {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "12.0.0.0")]
    internal sealed partial class Invoice : global::System.Configuration.ApplicationSettingsBase {
        
        private static Invoice defaultInstance = ((Invoice)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Invoice())));
        
        public static Invoice Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("John")]
        public string Customer_Job {
            get {
                return ((string)(this["Customer_Job"]));
            }
            set {
                this["Customer_Job"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Intuit Product Invoice")]
        public string Template {
            get {
                return ((string)(this["Template"]));
            }
            set {
                this["Template"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("DHL")]
        public string VIA {
            get {
                return ((string)(this["VIA"]));
            }
            set {
                this["VIA"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("FOB")]
        public string FOB {
            get {
                return ((string)(this["FOB"]));
            }
            set {
                this["FOB"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Net 15")]
        public string REP {
            get {
                return ((string)(this["REP"]));
            }
            set {
                this["REP"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("1")]
        public string Quantity {
            get {
                return ((string)(this["Quantity"]));
            }
            set {
                this["Quantity"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Simulator")]
        public string Item {
            get {
                return ((string)(this["Item"]));
            }
            set {
                this["Item"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("100HT")]
        public string Class {
            get {
                return ((string)(this["Class"]));
            }
            set {
                this["Class"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Accounts Receivable")]
        public string Account {
            get {
                return ((string)(this["Account"]));
            }
            set {
                this["Account"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Invoice")]
        public string Test_Name {
            get {
                return ((string)(this["Test_Name"]));
            }
            set {
                this["Test_Name"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("BATS")]
        public string Module_Name {
            get {
                return ((string)(this["Module_Name"]));
            }
            set {
                this["Module_Name"] = value;
            }
        }
    }
}
