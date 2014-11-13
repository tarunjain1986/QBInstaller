
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.WindowStripControls;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using System.IO;
using TestStack.White.UIItems.MenuItems;
using TestStack.White;
using System.Diagnostics;
using Xunit;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;


namespace FrameworkLibraries.ActionLibs.WhiteAPI
{
    public class Actions
    {
        public static Property conf = Property.GetPropertyInstance();
        public static string Execution_Speed = conf.get("ExecutionSpeed");
        public static string Sync_Timeout = conf.get("SyncTimeOut");

        //**************************************************************************************************************************************************************

        public static bool SelectMenu(TestStack.White.Application app, Window win, string level1, string level2)
        {
            try
            {
                Logger.logMessage("Function call @ :" + DateTime.Now);
                MenuBar menu = win.MenuBar;
                //MenuBar qbMenu = app.GetWindow(win.Name).MenuBar;
                menu.MenuItem(level1, level2).Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SelectMenu " + level1 + "->" + level2 + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return true;
            }
            catch (Exception e)
            {
                Logger.logMessage("SelectMenu " + level1 + "->" + level2 + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************
      

        public static void SendF2ToWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.F2);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SendF2ToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendF2ToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************
        public static bool CheckMenuEnabled(TestStack.White.Application app, Window win, string level1)
        {
            try
            {
                Logger.logMessage("Function call @ :" + DateTime.Now);
                MenuBar qbMenu = app.GetWindow(win.Name).MenuBar;
                var status = qbMenu.MenuItem(level1).Enabled;
                if(status)
                    Logger.logMessage("CheckMenuEnabled " + level1 + "->" + " - Enabled");
                else
                    Logger.logMessage("CheckMenuEnabled " + level1 + "->" + " - Disabled");    
                Logger.logMessage("------------------------------------------------------------------------------");
                return status;
            }
            catch (Exception e)
            {
                Logger.logMessage("CheckMenuEnabled " + level1 + "->" + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static bool SelectMenu(TestStack.White.Application app, Window win, string level1, string level2, string level3)
        {
            try
            {
                Logger.logMessage("Function call @ :" + DateTime.Now);
                MenuBar qbMenu = app.GetWindow(win.Name).MenuBar;
                TestStack.White.UIItems.MenuItems.Menu m1 = qbMenu.MenuItem(level1);
                m1.SubMenu(level2).SubMenu(level3).Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SelectMenu " + level1 + "->" + level2 + "->" + level3 + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return true;
            }
            catch (Exception e)
            {
                Logger.logMessage("SelectMenu " + level1 + "->" + level2 + "->" + level3 + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static bool SelectMenu(TestStack.White.Application app, Window win, String[] args)
        {
            try
            {
                Logger.logMessage("Function call @ :" + DateTime.Now);
                MenuBar qbMenu = app.GetWindow(win.Name).MenuBar;

                foreach (String item in args)
                {
                    qbMenu.MenuItem(item).Click();
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                return true;
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static Window GetWindow(Window win, String winName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Window window = null;

            try
            {
                List<Window> modalWins = win.ModalWindows();
                foreach (Window item in modalWins)
                {
                    Logger.logMessage(item.Name);

                    if (item.Name.Equals(winName) || item.Name.Contains(winName))
                    {
                        window = item;
                        window.Focus();
                        window.DoubleClick();
                        Thread.Sleep(int.Parse(Execution_Speed));
                    }
                }
                Logger.logMessage("GetWindow " + winName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return window;

            }
            catch (Exception e)
            {
                Logger.logMessage("GetWindow " + winName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static UIItemCollection GetWindowItems(Window win)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            UIItemCollection items = null;

            try
            {
                items = win.Items;
                foreach (var item in items)
                {
                    Logger.logMessage(item.ToString());
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("GetWindowItems " + win + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return items;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetWindowItems " + win + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void HighLightWindowElements(UIItemCollection collection)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                foreach (IUIItem item in collection)
                {
                    try
                    {
                        item.RightClick();
                        Thread.Sleep(int.Parse(Execution_Speed));
                    }
                    catch (Exception)
                    {
                    }
                }
                Logger.logMessage("HighLightWindowElements " + collection + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("HighLightWindowElements " + collection + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void ClickWPFButton(Window win, UIItemCollection collection, String text)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                foreach (var item in collection)
                {
                    if (item.GetType().Name.Equals("WPFButton"))
                    {
                        AutomationProperty[] properties = item.AutomationElement.GetSupportedProperties();
                        foreach (AutomationProperty p in properties)
                        {
                            if (item.AutomationElement.GetCurrentPropertyValue(p).Equals(text))
                            {
                                item.Click();
                                Thread.Sleep(int.Parse(Execution_Speed));
                                break;
                            }
                        }
                    }
                    Logger.logMessage("ClickWPFButton " + win + "->" + collection + "->" + text + " - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
            }
            catch (Exception e)
            {
                Logger.logMessage("ClickWPFButton " + collection + "->" + text + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static bool ClickButtonByOrientation(UIItemCollection item, String identifier)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            AutomationProperty p = AutomationElementIdentifiers.OrientationProperty;
            try
            {
                foreach (IUIItem element in item)
                {
                    if (element.GetType().Name.Contains("Button") || element.GetType().Name.Equals("Button"))
                    {
                        if (element.Name.Equals(identifier) || element.Id.Equals(identifier) || element.PrimaryIdentification.Equals(identifier))
                        {
                            element.Click();
                            Thread.Sleep(int.Parse(Execution_Speed));
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void HighLightAndShowProperties(UIItemCollection item, String elementType)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            String spy = null;

            try
            {
                foreach (var element in item)
                {
                    if (element.GetType().Name.Contains(elementType) || element.GetType().Name.Equals(elementType))
                    {
                        element.RightClick();
                        var name = element.Name;
                        var id = element.Id;
                        var pid = element.PrimaryIdentification;
                        MessageBox.Show(" Name= " + name + " ID= " + id + " PrimaryID= " + pid + " / ");
                        AutomationProperty[] properties = element.AutomationElement.GetSupportedProperties();
                        foreach (var p in properties)
                        {
                            try
                            {

                                var value = element.AutomationElement.GetCurrentPropertyValue(p);
                                var property = p.ProgrammaticName;
                                spy = spy + " -- " + value + " / " + property;
                            }
                            catch (Exception e)
                            {
                                var err = e.Message;
                            }
                        }
                        MessageBox.Show(spy.ToString());
                        spy = null;
                    }
                }
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void ShowWindowElementTypes(Window win)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            String spy = null;
            SortedSet<String> Types = new SortedSet<string>();

            try
            {
                UIItemCollection allElements = win.Items;

                foreach (var element in allElements)
                {
                    Types.Add(element.GetType().Name);
                }

                foreach (String item in Types)
                {
                    spy = spy + " / " + item;
                }

                MessageBox.Show(spy);
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void GetCurrsorToFirstTextBox(Window win)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                UIItemCollection allElements = win.Items;

                foreach (var element in allElements)
                {
                    if (element.GetType().Name.Equals("TextBox"))
                    {
                        element.Focus();
                        element.Click();
                        Thread.Sleep(int.Parse(Execution_Speed));
                        break;
                    }
                }
                Logger.logMessage("GetCurrsorToFirstTextBox " + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("GetCurrsorToFirstTextBox " + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void ClickElementByAutomationID(Window win, String automationID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                win.Get(SearchCriteria.ByAutomationId(automationID)).Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("ClickElementByAutomationID " + automationID + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("ClickElementByAutomationID " + automationID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static bool CheckElementExistsByAutomationID(Window win, String automationID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            bool exists = false;

            try
            {
                try { exists = win.Get(SearchCriteria.ByAutomationId(automationID)).Visible; }
                catch { }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("CheckElementExistsByAutomationID " + win + "->" + automationID + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return exists;
            }
            catch (Exception e)
            {
                Logger.logMessage("CheckElementExistsByAutomationID " + win + "->" + automationID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static bool CheckElementExistsByName(Window win, String name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            bool exists = false;

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                try { exists = win.Get(SearchCriteria.ByNativeProperty(p, name)).Visible; }
                catch { }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("CheckElementExistsByName " + win + "->" + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return exists;
            }
            catch (Exception e)
            {
                Logger.logMessage("CheckElementExistsByName " + win + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void ShowWindowElementAutomationIDs(Window win)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            String spy = null;

            try
            {
                UIItemCollection elements = win.Items;

                foreach (IUIItem e in elements)
                {
                    AutomationProperty p = AutomationElementIdentifiers.AutomationIdProperty;
                    spy = spy + " / " + e.AutomationElement.GetCurrentPropertyValue(p).ToString();
                }
                MessageBox.Show(spy);
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SetFocusOnWindow(Window win)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                win.Focus();
                win.Click();
                Logger.logMessage("SetFocusOnWindow " + win + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                Thread.Sleep(int.Parse(Execution_Speed));
            }
            catch (Exception e)
            {
                Logger.logMessage("SetFocusOnWindow " + win + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void ClickElementByName(Window win, String name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            string windowName = null;

            try
            {
                windowName = win.Name;
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                win.Get(SearchCriteria.ByNativeProperty(p, name)).Click();
                win.WaitWhileBusy();
                Logger.logMessage("ClickElementByName " + windowName + "->" + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                Thread.Sleep(int.Parse(Execution_Speed));
            }
            catch (Exception e)
            {
                Logger.logMessage("ClickElementByName " + windowName + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void ClickElementByMatchingName(Window win, String matchingName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                UIItemCollection allItems = win.Items;
                foreach (IUIItem item in allItems)
                {
                    if (item.Name.Contains(matchingName))
                    {
                        item.Click();
                    }
                }
                Logger.logMessage("ClickElementByMatchingName " + win + "->" + matchingName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                Thread.Sleep(int.Parse(Execution_Speed));
            }
            catch (Exception e)
            {
                Logger.logMessage("ClickElementByMatchingName " + win + "->" + matchingName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SetTextOnElementByName(Window win, String name, String value)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                win.Get(SearchCriteria.ByNativeProperty(p, name)).Enter(value);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SetTextOnElementByName " + win + "->" + name + "->" + value + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SetTextOnElementByName " + win + "->" + name + "->" + value + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void ClickButtonByAutomationID(Window win, String automationID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                TestStack.White.UIItems.Button b = (TestStack.White.UIItems.Button)win.Get(SearchCriteria.ByAutomationId(automationID));
                b.Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("ClickButtonByAutomationID " + win + "->" + automationID + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("ClickButtonByAutomationID " + win + "->" + automationID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void ClickMenuItemByName(Window win, String name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                TestStack.White.UIItems.MenuItems.Menu b = (TestStack.White.UIItems.MenuItems.Menu)win.Get(SearchCriteria.ByNativeProperty(p, name));
                b.Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("ClickMenuItemByName " + win + "->" + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("ClickMenuItemByName " + win + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static void SetTextByAutomationID(Window win, String automationID, String value)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                TestStack.White.UIItems.TextBox t = (TestStack.White.UIItems.TextBox)win.Get(SearchCriteria.ByAutomationId(automationID));
                t.SetValue(value);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SetTextByAutomationID " + win + "->" + automationID + "->" + value + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("SetTextByAutomationID " + win + "->" + automationID + "->" + value + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SetTextByName(Window win, String name, String value)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                TestStack.White.UIItems.TextBox t = (TestStack.White.UIItems.TextBox)win.Get(SearchCriteria.ByNativeProperty(p, name));
                t.SetValue(value);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SetTextByName " + win + "->" + name + "->" + value + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SetTextByName " + win + "->" + name + "->" + value + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void FileOutAutomationIDs(UIItemCollection collection, String fileName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                FileInfo test = new FileInfo(@fileName);
                string temp = string.Empty;

                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@fileName))
                {
                    foreach (IUIItem item in collection)
                    {
                        try
                        {
                            AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                            var element = item.AutomationElement.ToString();
                            var value = item.AutomationElement.GetCurrentPropertyValue(p).ToString();
                            temp = element + " / " + item.GetType().ToString() + " = " + value;
                            file.WriteLine(temp);
                            test.AppendText().WriteLine(temp);
                            test.AppendText().Flush();
                        }
                        catch (Exception)
                        {
                        }
                    }
                }

                System.IO.File.WriteAllText(@fileName, temp);

            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }



        //**************************************************************************************************************************************************************

        public static void SendTABToWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.TAB);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SendTABToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendTABToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void SendDOWNToWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DOWN);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SendDOWNToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendDOWNToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************


        public static void SendKeysToWindow(Window window, String key)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                foreach (char c in key)
                {
                    kb.Enter(c.ToString());
                    Thread.Sleep(25);
                }
                Thread.Sleep(200);
                Logger.logMessage("SendKeysToWindow " + window + "->" + key + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendKeysToWindow " + window + "->" + key + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void SendSHIFT_TABToWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SHIFT);
                kb.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.TAB);
                kb.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SHIFT);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SendSHIFT_TABToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendSHIFT_TABToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void SendNumbersToWindow(Window window, int input)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                foreach (var c in input.ToString())
                {
                    kb.Enter("" + Int32.Parse(c.ToString()) + "");
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SendNumbersToWindow " + window + "->" + input + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendNumbersToWindow " + window + "->" + input + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void SendBCKSPACEToWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.BACKSPACE);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SendBCKSPACEToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendBCKSPACEToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void CloseAllChildWindows(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                List<Window> modalWindows = window.ModalWindows();
                foreach (Window win in modalWindows)
                {
                    win.Focus();

                    try { FrameworkLibraries.ActionLibs.WhiteAPI.Actions.ClickElementByName(win, "Close"); }
                    catch { }

                    try { win.Close(); }
                    catch { }
                }
                Logger.logMessage("CloseAllChildWindows " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("CloseAllChildWindows " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void ClickButtonByName(Window win, String name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                TestStack.White.UIItems.Button b = (TestStack.White.UIItems.Button)win.Get(SearchCriteria.ByNativeProperty(p, name));
                b.Click();
                win.WaitWhileBusy();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("ClickButtonByName " + win + "->" + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("ClickButtonByName " + win + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void HighLightWindowElementsAndShowType(Window win)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                UIItemCollection allElements = win.Items;
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;

                foreach (var element in allElements)
                {
                    element.RightClick();
                    MessageBox.Show(element.GetType().ToString() + " - " + element.AutomationElement.GetCurrentPropertyValue(p));
                }
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<IUIItem> GetAllGroupBoxes(UIItemCollection collection)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<IUIItem> groupBoxes = new List<IUIItem>();

            try
            {
                foreach (IUIItem element in collection)
                {
                    if (element.GetType().Name.Contains("Group") || element.GetType().Equals("Group"))
                    {
                        groupBoxes.Add(element);
                    }
                }
                Logger.logMessage("GetAllGroupBoxes " + collection + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return groupBoxes;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllGroupBoxes " + collection + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<IUIItem> GetAllListItems(UIItemCollection collection)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<IUIItem> listItems = new List<IUIItem>();

            try
            {
                foreach (IUIItem element in collection)
                {
                    if (element.GetType().Name.Contains("List") || element.GetType().Equals("List"))
                    {
                        listItems.Add(element);
                    }
                }
                Logger.logMessage("GetAllListItems " + collection + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return listItems;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllListItems " + collection + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SelectListBoxItemByText(Window win, String listBoxElementAutoID, String matchText)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.UIItems.ListBoxItems.ListBox l = win.Get<TestStack.White.UIItems.ListBoxItems.ListBox>(SearchCriteria.ByAutomationId(listBoxElementAutoID));
                List<TestStack.White.UIItems.ListBoxItems.ListItem> k = l.Items;
                foreach (var item in k)
                {
                    if (item.Text.Equals(matchText) || item.Text.Contains(matchText))
                    {
                        item.Focus();
                        //item.SetValue(matchText);
                        item.Click();
                        Thread.Sleep(int.Parse(Execution_Speed));
                    }
                }
                Logger.logMessage("SelectListBoxItemByText " + win + "->" + listBoxElementAutoID + "->" + matchText + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SelectListBoxItemByText " + win + "->" + listBoxElementAutoID + "->" + matchText + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void GetElementCountOfType(Window win, String type)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            int count = 0;
            SortedSet<String> Types = new SortedSet<string>();

            try
            {
                UIItemCollection allElements = win.Items;

                foreach (var element in allElements)
                {
                    if (element.GetType().Name.Equals(type) || element.GetType().Name.Contains(type))
                        count++;

                }

                MessageBox.Show(type + " elements = " + count.ToString());
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<IUIItem> GetAllTextBoxes(UIItemCollection collection)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<IUIItem> textBoxes = new List<IUIItem>();

            try
            {
                foreach (IUIItem element in collection)
                {
                    if (element.GetType().Name.Contains("Text") || element.GetType().Equals("Text"))
                    {
                        TestStack.White.UIItems.TextBox x = (TestStack.White.UIItems.TextBox)element;
                        textBoxes.Add(x);
                    }
                }
                Logger.logMessage("GetAllTextBoxes " + collection + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return textBoxes;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllTextBoxes " + collection + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<IUIItem> GetAllLabels(UIItemCollection collection)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<IUIItem> labels = new List<IUIItem>();

            try
            {
                foreach (IUIItem element in collection)
                {
                    if (element.GetType().Name.Contains("Text") || element.GetType().Equals("Text"))
                    {
                        TestStack.White.UIItems.Label x = (TestStack.White.UIItems.Label)element;
                        labels.Add(x);
                    }
                }
                Logger.logMessage("GetAllLabels " + collection + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return labels;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllLabels " + collection + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<IUIItem> GetAllPanels(UIItemCollection collection)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<IUIItem> panels = new List<IUIItem>();

            try
            {
                foreach (IUIItem element in collection)
                {
                    if (element.GetType().Name.Contains("Pane") || element.GetType().Equals("Pane"))
                    {
                        TestStack.White.UIItems.Panel x = (TestStack.White.UIItems.Panel)element;
                        panels.Add(x);
                    }
                }
                Logger.logMessage("GetAllPanels " + collection + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return panels;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllPanels " + collection + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static TestStack.White.UIItems.Panel GetPanelElementByName(UIItemCollection collection, string elementName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            TestStack.White.UIItems.Panel p = null;

            try
            {
                foreach (IUIItem element in collection)
                {
                    if (element.GetType().Name.Contains("Pane") && element.Name.Equals(elementName))
                    {
                        p = (TestStack.White.UIItems.Panel)element;
                        break;
                    }
                }
                Logger.logMessage("GetPanelElementByName " + collection + "->" + elementName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return p;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetPanelElementByName " + collection + "->" + elementName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<IUIItem> GetAllCheckboxes(UIItemCollection collection)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<IUIItem> checkBoxes = new List<IUIItem>();

            try
            {
                foreach (IUIItem element in collection)
                {
                    if (element.GetType().Name.Contains("Check") || element.GetType().Equals("Check"))
                    {
                        TestStack.White.UIItems.CheckBox x = (TestStack.White.UIItems.CheckBox)element;
                        checkBoxes.Add(x);
                    }
                }
                Logger.logMessage("GetAllCheckboxes " + collection + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return checkBoxes;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllCheckboxes " + collection + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void UIA_SetFocusOfFirstTextBox(AutomationElement uiaWindow, Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                PropertyCondition textCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                AutomationElementCollection textBoxes = uiaWindow.FindAll(TreeScope.Descendants, textCondition);
                foreach (AutomationElement textBox in textBoxes)
                {
                    TestStack.White.UIItems.TextBox t = new TestStack.White.UIItems.TextBox(textBox, window.ActionListener);
                    t.Focus();
                    t.Click();
                    Thread.Sleep(int.Parse(Execution_Speed));
                    break;
                }
                Logger.logMessage("UIA_SetFocusOfFirstTextBox " + uiaWindow + "->" + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_SetFocusOfFirstTextBox " + uiaWindow + "->" + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void UIA_SetTextByName(AutomationElement uiaWindow, Window window, string name, string value)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                PropertyCondition textCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                AutomationElementCollection textBoxes = uiaWindow.FindAll(TreeScope.Descendants, textCondition);
                foreach (AutomationElement e in textBoxes)
                {
                    if (e.Current.Name.Equals(name))
                    {
                        TestStack.White.UIItems.TextBox t = new TestStack.White.UIItems.TextBox(e, window.ActionListener);
                        t.Text = value;
                    }
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_SetTextByName " + uiaWindow + "->" + window + "->" + name + "->" + value + "->" + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_SetTextByName " + uiaWindow + "->" + window + "->" + name + "->" + value + "->" + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************
        public static void UIA_ClickTextByName(AutomationElement uiaWindow, Window window, string name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                PropertyCondition textCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                AutomationElementCollection texts = uiaWindow.FindAll(TreeScope.Descendants, textCondition);
                foreach (AutomationElement e in texts)
                {
                    if (e.Current.Name.Equals(name))
                    {
                        TestStack.White.UIItems.Label t = new TestStack.White.UIItems.Label(e, window.ActionListener);
                        t.Click();
                    }
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_ClickTextByName " + uiaWindow + "->" + window + "->" + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_ClickTextByName " + uiaWindow + "->" + window + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void UIA_ClickEditControlByName(AutomationElement uiaWindow, Window window, string name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                PropertyCondition textCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                AutomationElementCollection texts = uiaWindow.FindAll(TreeScope.Descendants, textCondition);
                foreach (AutomationElement e in texts)
                {
                    if (e.Current.Name.Equals(name))
                    {
                        TestStack.White.UIItems.TextBox t = new TestStack.White.UIItems.TextBox(e, window.ActionListener);
                        t.Click();
                    }
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_ClickTextByName " + uiaWindow + "->" + window + "->" + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_ClickTextByName " + uiaWindow + "->" + window + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void UIA_SelectCheckBoxByName(AutomationElement uiaWindow, Window window, string name, bool state)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                PropertyCondition checkBoxCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox);
                AutomationElementCollection checkBoxes = uiaWindow.FindAll(TreeScope.Descendants, checkBoxCondition);
                foreach (AutomationElement e in checkBoxes)
                {
                    if (e.Current.Name.Equals(name))
                    {
                        TestStack.White.UIItems.CheckBox t = new TestStack.White.UIItems.CheckBox(e, window.ActionListener);
                        if (state)
                            t.Select();
                        break;
                    }
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_SelectCheckBoxByName " + uiaWindow + "->" + window + "->" + name + "->" + "state" + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_SelectCheckBoxByName " + uiaWindow + "->" + window + "->" + name + "->" + "state" + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static void UIA_ClickButtonByName(AutomationElement uiaWindow, Window window, string name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                PropertyCondition textCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                AutomationElementCollection buttons = uiaWindow.FindAll(TreeScope.Descendants, textCondition);
                foreach (AutomationElement e in buttons)
                {
                    if (e.Current.Name.Equals(name))
                    {
                        TestStack.White.UIItems.Button t = new TestStack.White.UIItems.Button(e, window.ActionListener);
                        t.Click();
                    }
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_ClickButtonByName " + uiaWindow + "->" + window + "->" + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_ClickButtonByName " + uiaWindow + "->" + window + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************



        public static void UIA_ClickTextByAutomationID(AutomationElement uiaWindow, Window window, string automationID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                PropertyCondition textCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                AutomationElementCollection texts = uiaWindow.FindAll(TreeScope.Descendants, textCondition);
                foreach (AutomationElement e in texts)
                {
                    if (e.Current.AutomationId.Equals(automationID))
                    {
                        TestStack.White.UIItems.Label t = new TestStack.White.UIItems.Label(e, window.ActionListener);
                        t.Click();
                    }
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_ClickTextByAutomationID " + uiaWindow + "->" + window + "->" + automationID + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_ClickTextByAutomationID " + uiaWindow + "->" + window + "->" + automationID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static void UIA_SetTextByAutomationID(AutomationElement uiaWindow, Window window, string automationID, string value)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                PropertyCondition textCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                AutomationElementCollection textBoxes = uiaWindow.FindAll(TreeScope.Descendants, textCondition);
                foreach (AutomationElement e in textBoxes)
                {
                    if (e.Current.AutomationId.Equals(automationID))
                    {
                        TestStack.White.UIItems.TextBox t = new TestStack.White.UIItems.TextBox(e, window.ActionListener);
                        t.Text = value;
                    }
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_SetTextByAutomationID " + uiaWindow + "->" + window + "->" + automationID + "->" + value + "->" + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_SetTextByAutomationID " + uiaWindow + "->" + window + "->" + automationID + "->" + value + "->" + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static AutomationElement UIA_GetAppWindow(string windowName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                PropertyCondition windowTypeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
                PropertyCondition windowAutomationIDCondition = new PropertyCondition(AutomationElement.NameProperty, windowName);
                AndCondition windowCondition = new AndCondition(windowTypeCondition, windowAutomationIDCondition);
                AutomationElement window = AutomationElement.RootElement.FindFirst(TreeScope.Children, windowCondition);
                Logger.logMessage("UIA_GetAppWindow " + windowName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return window;
            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_GetAppWindow " + windowName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static AutomationElement UIA_GetChildWindow(AutomationElement appWindow, string childWindowName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            AutomationElement childWindow = null;

            try
            {
                PropertyCondition windowTypeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
                PropertyCondition windowNameCondition = new PropertyCondition(AutomationElement.NameProperty, childWindowName);
                AndCondition windowCondition = new AndCondition(windowTypeCondition, windowNameCondition);
                AutomationElement window = appWindow.FindFirst(TreeScope.Children, windowCondition);

                AutomationElementCollection windows = appWindow.FindAll(TreeScope.Descendants, windowTypeCondition);

                foreach (AutomationElement w in windows)
                {
                    if (w.Current.Name.Equals(childWindowName) || w.Current.Name.Contains(childWindowName))
                    {
                        childWindow = w;
                        break;
                    }
                }
                Logger.logMessage("UIA_GetChildWindow " + appWindow + "->" + childWindowName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return childWindow;
            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_GetChildWindow " + appWindow + "->" + childWindowName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************
        public static Window GetChildWindow(Window mainWindow, string childWindowName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Window childWindow = null;

            try
            {
                List<Window> allChildWindows = mainWindow.ModalWindows();

                foreach (Window w in allChildWindows)
                {
                    if (w.Name.Equals(childWindowName) || w.Name.Contains(childWindowName))
                    {
                        childWindow = w;
                        break;
                    }
                }
                Logger.logMessage("GetChildWindow " + mainWindow + "->" + childWindowName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return childWindow;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetChildWindow " + mainWindow + "->" + childWindowName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static bool WaitForChildWindow(Window mainWindow, string childWindowName, long timeOut)
        {
            var qbApp = QuickBooks.GetApp("QuickBooks");
            var qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("                 WaitForChildWindow " + mainWindow + "->" + childWindowName + " - Begin Sync");
            bool windowFound = false;
            long elapsedTime = 0;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                do
                {
                    if (Actions.CheckDesktopWindowExists("Alert"))
                        Actions.CheckForAlertAndClose("Alert");

                    try { Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Warning"), "OK"); }
                    catch (Exception) { }

                    //Crash handler
                    if (Actions.CheckDesktopWindowExists("QuickBooks - Unrecoverable Error"))
                    {
                        Actions.QBCrashHandler();
                        break;
                    }

                    if (windowFound)
                        break;

                    elapsedTime = stopwatch.ElapsedMilliseconds;

                    List<Window> allChildWindows = mainWindow.ModalWindows();

                    foreach (Window w in allChildWindows)
                    {

                        if (Actions.CheckDesktopWindowExists("Alert"))
                            Actions.CheckForAlertAndClose("Alert");

                        try { Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Warning"), "OK"); }
                        catch (Exception) { }

                        //Crash handler
                        if (Actions.CheckDesktopWindowExists("QuickBooks - Unrecoverable Error"))
                        {
                            Actions.QBCrashHandler();
                            break;
                        }


                        if (w.Name.Equals(childWindowName) || w.Name.Contains(childWindowName))
                        {
                            windowFound = true;
                            w.WaitWhileBusy();
                            break;
                        }
                    }
                }
                while (elapsedTime <= timeOut);
                Logger.logMessage("                 WaitForChildWindow " + mainWindow + "->" + childWindowName + " - End Sync");
                Logger.logMessage("------------------------------------------------------------------------------");
                return windowFound;
            }
            catch (Exception e)
            {
                Logger.logMessage("WaitForChildWindow " + mainWindow + "->" + childWindowName + " - Terminated");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static bool WaitForWindow(string windowName, long timeOut)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("                 WaitForWindow " + windowName + " - Begin Sync");
            bool windowFound = false;
            long elapsedTime = 0;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                do
                {

                    elapsedTime = stopwatch.ElapsedMilliseconds;

                    List<Window> allChildWindows = Desktop.Instance.Windows();

                    foreach (Window w in allChildWindows)
                    {
                        if (w.Name.Equals(windowName) || w.Name.Contains(windowName))
                        {
                            windowFound = true;
                            w.WaitWhileBusy();
                            break;
                        }
                    }
                }
                while (elapsedTime <= timeOut && windowFound==false);
                Logger.logMessage("                 WaitForChildWindow " + windowName + " - End Sync");
                Logger.logMessage("------------------------------------------------------------------------------");
                return windowFound;
            }
            catch (Exception e)
            {
                Logger.logMessage("                 WaitForChildWindow " + windowName + " - Terminated");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static bool WaitForAnyChildWindow(Window mainWindow, string currentWindowName, long timeOut)
        {
            var qbApp = QuickBooks.GetApp("QuickBooks");
            var qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("                 WaitForAnyChildWindow " + mainWindow + "->" + currentWindowName + " - Begin Sync");

            bool windowFound = false;
            long elapsedTime = 0;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                do
                {
                    //Alert window handler
                    if (Actions.CheckDesktopWindowExists("Alert"))
                        Actions.CheckForAlertAndClose("Alert");

                    try { Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Warning"), "OK"); }
                    catch (Exception) { }

                    //Crash handler
                    if (Actions.CheckDesktopWindowExists("QuickBooks - Unrecoverable Error"))
                    {
                        Actions.QBCrashHandler();
                        break;
                    }

                    if (windowFound)
                        break;

                    elapsedTime = stopwatch.ElapsedMilliseconds;

                    List<Window> allChildWindows = mainWindow.ModalWindows();

                    foreach (Window w in allChildWindows)
                    {

                        if (Actions.CheckDesktopWindowExists("Alert"))
                            Actions.CheckForAlertAndClose("Alert");

                        try { Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Warning"), "OK"); }
                        catch (Exception) { }

                        //Crash handler
                        if (Actions.CheckDesktopWindowExists("QuickBooks - Unrecoverable Error"))
                        {
                            Actions.QBCrashHandler();
                            break;
                        }

                        if (!w.Name.Equals(currentWindowName) || !w.Name.Contains(currentWindowName))
                        {
                            windowFound = true;
                            w.WaitWhileBusy();
                            break;
                        }
                    }
                }
                while (elapsedTime <= timeOut);
                Logger.logMessage("                 WaitForAnyChildWindow " + mainWindow + "->" + currentWindowName + " - End Sync");
                Logger.logMessage("------------------------------------------------------------------------------");

                return windowFound;
            }
            catch (Exception e)
            {
                Logger.logMessage("WaitForAnyChildWindow " + mainWindow + "->" + currentWindowName + " - Terminated");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static bool WaitForElementEnabledOrTransformed(Window window, string elementName, string transformName, long timeOut)
        {

            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("                 WaitForElementEnabledOrTransformed " + window + "->" + elementName + "->"+ transformName + " - Begin Sync");

            bool enabled = false;
            long elapsedTime = 0;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                do
                {
                    elapsedTime = stopwatch.ElapsedMilliseconds;

                    var elements = window.Items;

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        if (Actions.CheckElementExistsByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Files in Use"))
                        {
                            Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Do not close applications. (A reboot will be required.)");
                            Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "OK");
                        }
                    }
                    catch { }

                    foreach (var w in elements)
                    {

                        if (w.Name.Equals(elementName))
                        {
                            enabled = w.Enabled;
                            try
                            {
                                if (enabled || Actions.CheckElementExistsByName(window, transformName))
                                    break;
                            }
                            catch { }
                        }
                    }
                }
                while (elapsedTime <= timeOut && enabled == false);
                Logger.logMessage("                 WaitForElementEnabledOrTransformed " + window + "->" + elementName + "->" + transformName + " - End Sync");
                Logger.logMessage("------------------------------------------------------------------------------");

                return enabled;
            }
            catch (Exception e)
            {
                Logger.logMessage("                 WaitForElementEnabledOrTransformed " + window + "->" + elementName + "->" + transformName + " - Terminated");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static bool WaitForElementEnabled(Window window, string elementName, long timeOut)
        {

            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("                 WaitForElementEnabled " + window + "->" + elementName + " - Begin Sync");

            bool enabled = false;
            long elapsedTime = 0;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                do
                {
                    elapsedTime = stopwatch.ElapsedMilliseconds;

                    var elements = window.Items;

                    foreach (var w in elements)
                    {

                        if (w.Name.Equals(elementName))
                        {
                            enabled = w.Enabled;
                            if(enabled)
                                break;
                        }
                    }
                }
                while (elapsedTime <= timeOut && enabled==false);
                Logger.logMessage("                 WaitForElementEnabled " + window + "->" + elementName + " - End Sync");
                Logger.logMessage("------------------------------------------------------------------------------");

                return enabled;
            }
            catch (Exception e)
            {
                Logger.logMessage("WaitForElementEnabled " + window + "->" + elementName + " - Terminated");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static bool WaitForElementVisible(Window window, string elementName, long timeOut)
        {

            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("                 WaitForElementVisible " + window + "->" + elementName + " - Begin Sync");

            bool visible = false;
            long elapsedTime = 0;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                do
                {
                    elapsedTime = stopwatch.ElapsedMilliseconds;

                    var elements = window.Items;

                    foreach (var w in elements)
                    {
                        if (w.Name.Equals(elementName))
                        {
                            visible = w.Visible;
                            if (visible)
                                break;
                        }
                    }
                }
                while (elapsedTime <= timeOut && visible == false);
                Logger.logMessage("                 WaitForElementVisible " + window + "->" + elementName + " - End Sync");
                Logger.logMessage("------------------------------------------------------------------------------");

                return visible;
            }
            catch (Exception e)
            {
                Logger.logMessage("WaitForElementVisible " + window + "->" + elementName + " - Terminated");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static bool WaitForAppWindow(string appWindowName, long timeOut)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("                 WaitForAppWindow " + appWindowName + "->" + " - Begin Sync");

            bool windowFound = false;
            long elapsedTime = 0;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                do
                {
                    if (windowFound)
                        break;

                    elapsedTime = stopwatch.ElapsedMilliseconds;

                    List<Window> allChildWindows = Desktop.Instance.Windows();

                    foreach (Window w in allChildWindows)
                    {
                        if (Actions.CheckDesktopWindowExists("Alert"))
                            Actions.CheckForAlertAndClose("Alert");

                        //Crash handler
                        if (Actions.CheckDesktopWindowExists("QuickBooks - Unrecoverable Error"))
                        {
                            Actions.QBCrashHandler();
                            break;
                        }

                        if (w.Name.Equals(appWindowName) || w.Name.Contains(appWindowName))
                        {
                            windowFound = true;
                            w.WaitWhileBusy();
                            Thread.Sleep(int.Parse(Execution_Speed));
                            break;
                        }
                    }
                }
                while (elapsedTime <= timeOut);

                Logger.logMessage("                 WaitForAppWindow " + appWindowName + "->" + " - End Sync");
                Logger.logMessage("------------------------------------------------------------------------------");
                return windowFound;
            }
            catch (Exception e)
            {
                Logger.logMessage("WaitForAppWindow " + appWindowName + "->" + " - Terminated");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SendSHIFT_ENDToWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SHIFT);
                kb.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.END);
                kb.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SHIFT);
                Thread.Sleep(200);
                Logger.logMessage("SendSHIFT_ENDToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendSHIFT_ENDToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void SendENTERoWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SendENTERoWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendENTERoWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************


        public static void SelectComboBoxItemByText(Window win, String comboBoxAutoID, String matchText)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.UIItems.ListBoxItems.ComboBox c = win.Get<TestStack.White.UIItems.ListBoxItems.ComboBox>(SearchCriteria.ByAutomationId(comboBoxAutoID));
                var k = c.Items;
                foreach (var item in k)
                {
                    if (item.Text.Equals(matchText) || item.Text.Contains(matchText))
                    {
                        item.Focus();
                        item.Select();
                        Thread.Sleep(int.Parse(Execution_Speed));
                    }
                }
                Logger.logMessage("SelectComboBoxItemByText " + win + "->" + comboBoxAutoID + "->" + matchText + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SelectComboBoxItemByText " + win + "->" + comboBoxAutoID + "->" + matchText + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void SelectCheckBox(Window win, String checkBoxAutoID, bool state)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.UIItems.CheckBox c = win.Get<TestStack.White.UIItems.CheckBox>(SearchCriteria.ByAutomationId(checkBoxAutoID));
                if (state)
                {
                    c.Select();
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                else
                {
                    c.UnSelect();
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                Logger.logMessage("SelectCheckBox " + win + "->" + checkBoxAutoID + "->" + state + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("SelectCheckBox " + win + "->" + checkBoxAutoID + "->" + state + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SelectCheckBoxByName(Window win, String checkBoxName, bool state)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                TestStack.White.UIItems.CheckBox c = win.Get<TestStack.White.UIItems.CheckBox>(SearchCriteria.ByNativeProperty(p, checkBoxName));
                if (state)
                {
                    c.Select();
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                else
                {
                    c.UnSelect();
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                Logger.logMessage("SelectCheckBoxByName " + win + "->" + checkBoxName + "->" + state + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SelectCheckBoxByName " + win + "->" + checkBoxName + "->" + state + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SelectRadioButtonByName(Window win, String radioButtonName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                TestStack.White.UIItems.RadioButton r = win.Get<TestStack.White.UIItems.RadioButton>(SearchCriteria.ByNativeProperty(p, radioButtonName));
                r.Select();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SelectRadioButtonByName " + win + "->" + radioButtonName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SelectRadioButtonByName " + win + "->" + radioButtonName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SelectRadioButton(Window win, String radioButtonAutoID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.UIItems.RadioButton r = win.Get<TestStack.White.UIItems.RadioButton>(SearchCriteria.ByAutomationId(radioButtonAutoID));
                r.Select();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SelectRadioButton " + win + "->" + radioButtonAutoID + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SelectRadioButton " + win + "->" + radioButtonAutoID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static void SetTextOnElementByAutomationID(Window win, String automationID, String value)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                win.Get(SearchCriteria.ByAutomationId(automationID)).Enter(value);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SetTextOnElementByAutomationID " + win + "->" + automationID + "->" + value + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("SetTextOnElementByAutomationID " + win + "->" + automationID + "->" + value + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static bool CheckWindowExists(Window mainWindow, string childWindowName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            bool window = false;

            try
            {
                List<Window> allChildWindows = mainWindow.ModalWindows();

                foreach (Window w in allChildWindows)

                {
                    Logger.logMessage(w.ToString());
                    if (w.Name.Equals(childWindowName) || w.Name.Contains(childWindowName))
                    {
                        window = true;
                        Thread.Sleep(int.Parse(Execution_Speed));
                        break;
                    }
                }
                Logger.logMessage("CheckWindowExists " + mainWindow + "->" + childWindowName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return window;
            }
            catch (Exception e)
            {
                Logger.logMessage("CheckWindowExists " + mainWindow + "->" + childWindowName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SetFocusOnElementByAutomationID(Window win, String automationID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                win.Get(SearchCriteria.ByAutomationId(automationID)).Focus();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SetFocusOnElementByAutomationID " + win + "->" + automationID + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("SetFocusOnElementByAutomationID " + win + "->" + automationID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<IUIItem> GetAllButtons(UIItemCollection collection)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            List<IUIItem> buttons = new List<IUIItem>();

            try
            {
                foreach (IUIItem element in collection)
                {
                    if (element.GetType().Name.Contains("Button") || element.GetType().Equals("Button"))
                    {
                        TestStack.White.UIItems.Button x = (TestStack.White.UIItems.Button)element;
                        buttons.Add(x);
                    }
                }
                Logger.logMessage("GetAllButtons " + collection + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return buttons;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllButtons " + collection + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<IUIItem> GetAllCustomControls(UIItemCollection collection)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            List<IUIItem> customControls = new List<IUIItem>();

            try
            {
                foreach (IUIItem element in collection)
                {
                    if (element.GetType().Name.Contains("Custom") || element.GetType().Equals("Custom"))
                    {
                        var x = (TestStack.White.UIItems.Custom.CustomUIItem)element;
                        customControls.Add(x);
                    }
                }
                Logger.logMessage("GetAllCustomControls " + collection + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return customControls;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllCustomControls " + collection + " - Successful");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static void UIA_ClickOnPaneItem(AutomationElement uiaWindow, Window window, int index)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                PropertyCondition paneCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane);
                AutomationElementCollection allPanes = uiaWindow.FindAll(TreeScope.Descendants, paneCondition);

                TestStack.White.UIItems.Panel p = new TestStack.White.UIItems.Panel(allPanes[index], window.ActionListener);
                p.Focus();
                p.Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_ClickOnPaneItem " + uiaWindow + "->" + window + "->" + index + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_ClickOnPaneItem " + uiaWindow + "->" + window + "->" + index + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void UIA_ClickMenuItem(AutomationElement uiaWindow, Window window, string name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem);

                AutomationElement menuItem = uiaWindow.FindFirst(TreeScope.Descendants, condition);

                TestStack.White.UIItems.MenuItems.Menu p = new TestStack.White.UIItems.MenuItems.Menu(menuItem, window.ActionListener);
                p.Focus();
                p.Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_ClickMenuItem " + uiaWindow + "->" + window + "->" + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_ClickMenuItem " + uiaWindow + "->" + window + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void UIA_ClickItemByName(AutomationElement uiaWindow, Window window, string name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                PropertyCondition condition = new PropertyCondition(p, name);
                AutomationElement element = uiaWindow.FindFirst(TreeScope.Descendants, condition);

                TestStack.White.UIItems.UIItem e = new TestStack.White.UIItems.UIItem(element, window.ActionListener);
                e.Focus();
                e.Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_ClickItemByName " + uiaWindow + "->" + window + "->" + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_ClickItemByName " + uiaWindow + "->" + window + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void UIA_ClickItemByAutomationID(AutomationElement uiaWindow, Window window, string automationID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.AutomationIdProperty;
                PropertyCondition condition = new PropertyCondition(p, automationID);
                AutomationElement element = uiaWindow.FindFirst(TreeScope.Descendants, condition);

                TestStack.White.UIItems.UIItem e = new TestStack.White.UIItems.UIItem(element, window.ActionListener);
                e.Focus();
                e.Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("UIA_ClickItemByAutomationID " + uiaWindow + "->" + window + "->" + automationID + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("UIA_ClickItemByAutomationID " + uiaWindow + "->" + window + "->" + automationID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void DesktopInstance_ClickElementByName(string name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                SearchCriteria x = SearchCriteria.ByNativeProperty(p, name);
                var e = TestStack.White.Desktop.Instance.Get(x);
                e.Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("DesktopInstance_ClickElementByName " + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("DesktopInstance_ClickElementByName " + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void DesktopInstance_ClickElementByAutomationID(string automationID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.AutomationIdProperty;
                SearchCriteria x = SearchCriteria.ByAutomationId(automationID);
                var e = TestStack.White.Desktop.Instance.Get(x);
                e.Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("DesktopInstance_ClickElementByAutomationID " + automationID + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("DesktopInstance_ClickElementByAutomationID " + automationID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static bool DesktopInstance_CheckElementExistsByName(string name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.NameProperty;
                SearchCriteria x = SearchCriteria.ByNativeProperty(p, name);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("DesktopInstance_CheckElementExistsByName " + name + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return TestStack.White.Desktop.Instance.Exists(x);
            }
            catch (Exception e)
            {
                Logger.logMessage("DesktopInstance_CheckElementExistsByName " + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static bool DesktopInstance_CheckElementExistsByAutomationID(string automationID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                AutomationProperty p = AutomationElementIdentifiers.AutomationIdProperty;
                SearchCriteria x = SearchCriteria.ByAutomationId(automationID);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("DesktopInstance_CheckElementExistsByAutomationID " + automationID + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return TestStack.White.Desktop.Instance.Exists(x);
            }
            catch (Exception e)
            {
                Logger.logMessage("DesktopInstance_CheckElementExistsByAutomationID " + automationID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static Window GetAlertWindow(string alertWindowName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            Window alertWindow = null;
            string alertText = null;

            try
            {
                List<Window> allChildWindows = Desktop.Instance.Windows();

                foreach (Window w in allChildWindows)
                {
                    if (w.Name.Equals(alertWindowName) || w.Name.Contains(alertWindowName))
                    {
                        alertWindow = w;

                        var elements = alertWindow.Items;

                        foreach (var item in elements)
                        {
                            if (item.GetType().Name.Equals("Label"))
                            {
                                alertText = item.Name;
                            }
                        }
                        Logger.logMessage("GetAlertWindow " + alertWindowName + " - Successful");
                        Logger.logMessage(alertText);
                        Logger.logMessage("------------------------------------------------------------------------------");
                        break;
                    }
                }

                return alertWindow;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAlertWindow " + alertWindowName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static Window GetDesktopWindow(string windowName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            Window win = null;

            try
            {
                List<Window> allChildWindows = Desktop.Instance.Windows();

                foreach (Window w in allChildWindows)

                {
                    
                    
                    if (w.Name.Equals(windowName) || w.Name.Contains(windowName))
                    {
                        win = w;
                        break;
                    }
                }
                Logger.logMessage("GetDesktopWindow " + windowName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return win;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetDesktopWindow " + windowName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void QBCrashHandler()
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            Window crashWindow = null;
            Window reportWindow = null;

            try
            {
                crashWindow = Actions.GetDesktopWindow("QuickBooks - Unrecoverable Error");
                Actions.ClickElementByName(crashWindow, "View report.");

                reportWindow = Actions.GetDesktopWindow("View Error Report");
                Actions.ClickElementByName(reportWindow, "QBWin");

                var elements = reportWindow.Items;
                foreach (var item in elements)
                {
                    if (item.GetType().Name.Equals("TextBox"))
                    {
                        var text = item.Name;
                        Logger.logMessage("---------------QBW32 C++ Log-----------------------");
                        Logger.logMessage(text);
                        break;
                    }
                }

                Actions.SendDOWNToWindow(reportWindow);
                Actions.SendDOWNToWindow(reportWindow);
                Actions.SendDOWNToWindow(reportWindow);
                Actions.SendDOWNToWindow(reportWindow);
                Actions.SendDOWNToWindow(reportWindow);
                Actions.SendDOWNToWindow(reportWindow);

                Actions.ClickElementByName(reportWindow, "qbw32DOTNET");
                var elements_2 = reportWindow.Items;
                foreach (var item in elements_2)
                {
                    if (item.GetType().Name.Equals("TextBox"))
                    {
                        var text = item.Name;
                        Logger.logMessage("---------------QBW32 DOT NET Log-----------------------");
                        Logger.logMessage(text);
                        break;
                    }
                }
                Actions.ClickElementByName(reportWindow, "Close");
                crashWindow = Actions.GetDesktopWindow("QuickBooks - Unrecoverable Error");
                Actions.ClickElementByName(crashWindow, "Send");

                try
                {
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks"), "Close the program");
                }
                catch (Exception)
                { }

                Utils.OSOperations.KillProcess("qbw32");

                Logger.logMessage("QBCrashHandler - Successful");
                Logger.logMessage("Killing QBProcess..");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("QBCrashHandler - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void CloseAlertWindow(string alertWindowName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            Window alertWindow = null;
            string alertText = null;

            try
            {
                List<Window> allChildWindows = Desktop.Instance.Windows();

                foreach (Window w in allChildWindows)
                {
                    if (w.Name.Equals(alertWindowName) || w.Name.Contains(alertWindowName))
                    {
                        alertWindow = w;

                        var elements = alertWindow.Items;

                        foreach (var item in elements)
                        {
                            if (item.GetType().Name.Equals("Label"))
                            {
                                alertText = item.Name;
                            }
                        }

                        alertWindow.Close();

                        Logger.logMessage("CloseAlertWindow " + alertWindowName + " - Successful");
                        Logger.logMessage(alertText);
                        Logger.logMessage("------------------------------------------------------------------------------");
                        break;
                    }
                }

            }
            catch (Exception e)
            {
                Logger.logMessage("CloseAlertWindow " + alertWindowName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static bool XunitAssertEuqals(string obj1, string obj2)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                Assert.Equal(obj1, obj2);
                Logger.logMessage("XunitAssertEuqals " + obj1 + "->" + obj2 + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return true;
            }
            catch (Exception e)
            {
                Logger.logMessage("XunitAssertEuqals " + obj1 + "->" + obj2 + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static bool XunitAssertContains(string obj1, string obj2)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                Assert.Contains(obj1, obj2);
                Logger.logMessage("XunitAssertContains " + obj1 + "->" + obj2 + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return true;
            }
            catch (Exception e)
            {
                Logger.logMessage("XunitAssertContains " + obj1 + "->" + obj2 + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void CheckForAlertAndClose(string alertWindowName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            string alertText = null;
            List<Window> allChildWindows = null;
            int iteration = 0;

            try
            {
                do
                {

                    iteration = iteration + 1;

                    if (iteration > 10)
                        break;

                    allChildWindows = Desktop.Instance.Windows();

                    foreach (Window w in allChildWindows)
                    {
                        if (w.Name.Equals(alertWindowName) || w.Name.Contains(alertWindowName))
                        {
                            var elements = w.Items;

                            foreach (var item in elements)
                            {
                                if (item.GetType().Name.Equals("Label"))
                                {
                                    alertText = item.Name;
                                }
                            }
                            Logger.logMessage(alertText);
                            Logger.logMessage("------------------------------------------------------------------------------");

                            try
                            {
                                Logger.logMessage("---------------Try-Catch Block------------------------");
                                Actions.ClickElementByName(w, "OK");
                                Thread.Sleep(int.Parse(Execution_Speed));
                            }
                            catch (Exception) { }

                            try
                            {
                                Logger.logMessage("---------------Try-Catch Block------------------------");
                                w.Close();
                                Actions.ClickElementByName(w, "No");
                                Thread.Sleep(int.Parse(Execution_Speed));
                            }
                            catch (Exception) { }
                        }
                    }
                }
                while (!allChildWindows.Contains(Actions.GetAlertWindow("Alert")));
                Logger.logMessage("CheckForAlertAndClose - Successful");
            }
            catch (Exception e)
            {
                Logger.logMessage("CheckForAlertAndClose - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static bool CheckDesktopWindowExists(string alertWindowName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            bool exists = false;
            List<Window> allChildWindows = null;

            try
            {
                Logger.logMessage("CheckDesktopWindowExists - Checking for "+alertWindowName);
                allChildWindows = Desktop.Instance.Windows();
                foreach (Window w in allChildWindows)
                {
                   
                    if (w.Name.Equals(alertWindowName) || w.Name.Contains(alertWindowName))
                    {
                        exists = true;
                        Logger.logMessage("CheckDesktopWindowExists "+alertWindowName + " - Found");
                        break;
                    }
                }
                Logger.logMessage("------------------------------------------------------------------------------");
                return exists;
            }
            catch (Exception e)
            {
                Logger.logMessage("CheckDesktopWindowExists - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static TestStack.White.UIItems.Panel GetPaneByName(Window window, string paneName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            TestStack.White.UIItems.Panel pane = null;

            try
            {
                Logger.logMessage("GetPaneByName " + window + "->" + paneName);
                var allPanes = Actions.GetAllPanels(window.Items);
                foreach (IUIItem panel in allPanes)
                {
                    if (panel.Name.Equals(paneName))
                    {
                        pane = (TestStack.White.UIItems.Panel)panel;
                        Logger.logMessage("GetPaneByName " + window + "->" + paneName + " - Successful");
                        break;
                    }
                }
                Logger.logMessage("------------------------------------------------------------------------------");
                return pane;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetPaneByName " + window + "->" + paneName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static TestStack.White.UIItems.Panel GetPaneByAutomationID(Window window, string autoID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            TestStack.White.UIItems.Panel pane = null;

            try
            {
                Logger.logMessage("GetPaneByName " + window + "->" + autoID);
                var allPanes = Actions.GetAllPanels(window.Items);
                foreach (IUIItem panel in allPanes)
                {
                    if (panel.AutomationElement.Current.AutomationId.Equals(autoID))
                    {
                        pane = (TestStack.White.UIItems.Panel)panel;
                        Logger.logMessage("GetPaneByName " + window + "->" + autoID + " - Successful");
                        break;
                    }
                }
                Logger.logMessage("------------------------------------------------------------------------------");
                return pane;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetPaneByName " + window + "->" + autoID + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static TestStack.White.Application GetApp(string appName, string processName)
        {
            int processID = 0;
            TestStack.White.Application app = null;

            try
            {
                List<Window> allWin = Desktop.Instance.Windows();
                foreach (Window item in allWin)
                {
                    if (item.Name.Contains(appName))
                    {
                        foreach (Process p in Process.GetProcesses("."))
                        {
                            if (p.ProcessName.Contains(processName) || p.ProcessName.Contains(processName.ToUpper()) || p.ProcessName.Contains(processName.ToLower()))
                            {
                                processID = p.Id;
                                app = TestStack.White.Application.Attach(processID);
                                app.WaitWhileBusy();
                                Thread.Sleep(int.Parse(Execution_Speed));
                                break;
                            }
                        }
                    }
                }

                return app;
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static Window GetAppWindow(TestStack.White.Application app, string winName)
        {
            Window win = null;

            try
            {
                List<Window> allWin = app.GetWindows();

                foreach (Window item in allWin)
                {
                    if (item.Name.Contains(winName))
                    {
                        win = item;
                        break;
                    }
                }

                return win;
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void ClickTextInsidePanel(Window window, TestStack.White.UIItems.Panel pane, string text)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                Logger.logMessage("ClickTextInsidePanel " + window + "->" + pane + "->" + text);
                PropertyCondition textCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                AutomationElementCollection textElements = pane.AutomationElement.FindAll(TreeScope.Descendants, textCondition);

                foreach (AutomationElement item in textElements)
                {
                    if (item.Current.Name.Equals(text))
                    {
                        var t = new TestStack.White.UIItems.Label(item, window.ActionListener);
                        t.Focus();
                        t.Click();
                        Logger.logMessage("ClickTextInsidePanel " + window + "->" + pane + "->" + text + " - Successful");
                        Logger.logMessage("------------------------------------------------------------------------------");
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Logger.logMessage("ClickTextInsidePanel " + window + "->" + pane + "->" + text + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static TestStack.White.UIItems.TableItems.Table GetTableInsideAPaneByIndex(Window window, TestStack.White.UIItems.Panel pane, int index)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                Logger.logMessage("GetTableInsideAPaneByIndex " + window + "->" + pane + "->" + index);

                PropertyCondition tableCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Table);
                AutomationElementCollection tableElements = pane.AutomationElement.FindAll(TreeScope.Descendants, tableCondition);
                TestStack.White.UIItems.TableItems.Table table = new TestStack.White.UIItems.TableItems.Table(tableElements[index], window.ActionListener);

                Logger.logMessage("GetTableInsideAPaneByIndex " + window + "->" + pane + "->" + index + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                
                return table;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetTableInsideAPaneByIndex " + window + "->" + pane + "->" + index + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.Label> GetAllTableTextElements(TestStack.White.UIItems.TableItems.Table table, Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.Label> elements = null;

            try
            {
                Logger.logMessage("GetAllTableTextElements " + table + "->" + window);

                PropertyCondition tableElementsCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                AutomationElementCollection allTableElements = table.AutomationElement.FindAll(TreeScope.Descendants, tableElementsCondition);

                foreach(AutomationElement item in allTableElements)
                {
                    TestStack.White.UIItems.Label l = new TestStack.White.UIItems.Label(item, window.ActionListener);
                    elements.Add(l);
                }

                Logger.logMessage("GetAllTableTextElements " + table + "->" + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return elements;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllTableTextElements " + table + "->" + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.TextBox> GetAllTableEditBoxElements(TestStack.White.UIItems.TableItems.Table table, Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.TextBox> elements = null;

            try
            {
                Logger.logMessage("GetAllTableEditBoxElements " + table + "->" + window);

                PropertyCondition tableElementsCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                AutomationElementCollection allTableElements = table.AutomationElement.FindAll(TreeScope.Descendants, tableElementsCondition);

                foreach (AutomationElement item in allTableElements)
                {
                    TestStack.White.UIItems.TextBox l = new TestStack.White.UIItems.TextBox(item, window.ActionListener);
                    elements.Add(l);
                }

                Logger.logMessage("GetAllTableEditBoxElements " + table + "->" + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return elements;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllTableEditBoxElements " + table + "->" + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.Button> GetAllTableButtonElements(TestStack.White.UIItems.TableItems.Table table, Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.Button> elements = null;

            try
            {
                Logger.logMessage("GetAllTableButtonElements " + table + "->" + window);

                PropertyCondition tableElementsCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                AutomationElementCollection allTableElements = table.AutomationElement.FindAll(TreeScope.Descendants, tableElementsCondition);

                foreach (AutomationElement item in allTableElements)
                {
                    TestStack.White.UIItems.Button l = new TestStack.White.UIItems.Button(item, window.ActionListener);
                    elements.Add(l);
                }

                Logger.logMessage("GetAllTableButtonElements " + table + "->" + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return elements;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllTableButtonElements " + table + "->" + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.CheckBox> GetAllTableCheckBoxElements(TestStack.White.UIItems.TableItems.Table table, Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.CheckBox> elements = null;

            try
            {
                Logger.logMessage("GetAllTableCheckBoxElements " + table + "->" + window);

                PropertyCondition tableElementsCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox);
                AutomationElementCollection allTableElements = table.AutomationElement.FindAll(TreeScope.Descendants, tableElementsCondition);

                foreach (AutomationElement item in allTableElements)
                {
                    TestStack.White.UIItems.CheckBox l = new TestStack.White.UIItems.CheckBox(item, window.ActionListener);
                    elements.Add(l);
                }

                Logger.logMessage("GetAllTableCheckBoxElements " + table + "->" + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return elements;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllTableCheckBoxElements " + table + "->" + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************


        public static int GetTableRowNumberByMatchingText(TestStack.White.UIItems.TableItems.Table table, int columnCount, string text)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            
            int count = 0;
            int rowNumber = 0;

            try
            {
                Logger.logMessage("GetTableRowNumberByMatchingText " + table + "->" + text);

                PropertyCondition tableElementsCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                AutomationElementCollection allTableElements = table.AutomationElement.FindAll(TreeScope.Descendants, tableElementsCondition);

                foreach (AutomationElement item in allTableElements)
                {
                    count = count + 1;
                    if (item.Current.Name.Equals(text))
                    {
                        rowNumber = (count / columnCount) + 1;
                        Logger.logMessage("GetTableRowNumberByMatchingText " + table + "->" + text + " - Successful");
                        Logger.logMessage("------------------------------------------------------------------------------");
                        break;
                    }
                }

                return rowNumber;
                
            }
            catch (Exception e)
            {
                Logger.logMessage("GetTableRowNumberByMatchingText " + table + "->" + text + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void SendALT_KeyToWindow(Window window, string key)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.ALT);
                kb.Enter(key);
                kb.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.ALT);
                Thread.Sleep(200);
                Logger.logMessage("SendALT_KeyToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendALT_KeyToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void SendCTRL_KeyToWindow(Window window, string key)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.CONTROL);
                kb.Enter(key);
                kb.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.CONTROL);
                Thread.Sleep(200);
                Logger.logMessage("SendCTRL_KeyToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendCTRL_KeyToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************


        public static void SendSPACEToWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.SPACE);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SendSPACEToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendSPACEToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static TestStack.White.Application LaunchApp(string exePath, string appName)
        {
            Logger.logMessage("Initialize " + exePath);

            int processID = 0;
            TestStack.White.Application app = null;

            try
            {
                List<Window> allWin = Desktop.Instance.Windows();
                foreach (Window item in allWin)
                {
                    if (item.Name.Contains(appName))
                    {
                        foreach (Process p in Process.GetProcesses("."))
                        {
                            if (p.ProcessName.Contains(appName) || p.ProcessName.Contains(appName.ToUpper()) || p.ProcessName.Contains(appName.ToLower()))
                            {
                                processID = p.Id;
                                app = TestStack.White.Application.Attach(processID);
                                app.WaitWhileBusy();
                                Actions.WaitForAppWindow(appName, int.Parse(Sync_Timeout));
                                Logger.logMessage("Existing App instance found..!!");
                                return app;
                            }
                        }
                    }
                }

                Logger.logMessage("No existing App instance, so launching - " + exePath);
                Process proc = new Process();
                proc.StartInfo.FileName = exePath;
                proc.Start();
                Thread.Sleep(7500);

                foreach (Process p in Process.GetProcesses("."))
                {
                    if (p.ProcessName.Contains(appName) || p.ProcessName.Contains(appName.ToUpper()) || p.ProcessName.Contains(appName.ToLower()))
                    {
                        processID = p.Id;
                    }
                }
                app = TestStack.White.Application.Attach(processID);
                app.WaitWhileBusy();
                Thread.Sleep(int.Parse(Execution_Speed));

                Logger.logMessage("Initialize " + exePath + " - Sucessful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return app;
            }
            catch (Exception e)
            {
                Logger.logMessage("Initialize " + exePath + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        
        public static void SendESCAPEToWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.InputDevices.AttachedKeyboard kb = window.Keyboard;
                kb.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.ESCAPE);
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SendESCAPEToWindow " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SendESCAPEToWindow " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static bool VerifyCheckBoxIsSelectedByAutomationID(Window win, String checkBoxAutoID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.UIItems.CheckBox c = win.Get<TestStack.White.UIItems.CheckBox>(SearchCriteria.ByAutomationId(checkBoxAutoID));
                bool state = c.IsSelected;
                if (state)
                {
                    Logger.logMessage("VerifyCheckBoxIsSelectedByAutomationID " + win + "->" + checkBoxAutoID + "->" + state + " - Successful");
                    return true;
                }
                else
                {
                    Logger.logMessage("VerifyCheckBoxIsSelectedByAutomationID " + win + "->" + checkBoxAutoID + "->" + state + " - Successful");
                    return false;
                }
            }
            catch (Exception e)
            {
                Logger.logMessage("SelectCheckBox " + win + "->" + checkBoxAutoID + "->" + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.Panel> GetAllPanesInWindow(Window window)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.Panel> collection = new List<TestStack.White.UIItems.Panel>();

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane);
                AutomationElementCollection elements = window.AutomationElement.FindAll(System.Windows.Automation.TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    TestStack.White.UIItems.Panel x = new TestStack.White.UIItems.Panel(item, window.ActionListener);
                    collection.Add(x);
                }

                Logger.logMessage("GetAllPanels " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return collection;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllPanels " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void SetText(TestStack.White.UIItems.TextBox element, String value)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                element.Text = value;
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("SetText " + element + "->" + value + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("SetText " + element + "->" + value + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void ClickElement(TestStack.White.UIItems.IUIItem element)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                element.Click();
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("ClickElement " + element + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("ClickElement " + element + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static void ClickButtonInsidePanelByName(Window window, TestStack.White.UIItems.Panel pane, string name)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);

            try
            {
                Logger.logMessage("ClickButtonInsidePanelByName " + window + "->" + pane + "->" + name);

                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    if (item.Current.Name.Equals(name))
                    {
                        var t = new TestStack.White.UIItems.Button(item, window.ActionListener);
                        t.Focus();
                        t.Click();
                        Logger.logMessage("ClickButtonInsidePanelByName " + window + "->" + pane + "->" + name + " - Successful");
                        Logger.logMessage("------------------------------------------------------------------------------");
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Logger.logMessage("ClickTextInsidePanel " + window + "->" + pane + "->" + name + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static bool WaitForTextVisibleInsidePane(Window window, TestStack.White.UIItems.Panel pane, string elementName, long timeOut)
        {

            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("                     WaitForTextVisibleInsidePane " + window + "->" + pane + "->" + elementName + " - Begin Sync");

            bool visible = false;
            long elapsedTime = 0;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                do
                {
                    elapsedTime = stopwatch.ElapsedMilliseconds;

                    PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                    AutomationElementCollection elements = pane.AutomationElement.FindAll(TreeScope.Descendants, condition);

                    foreach (AutomationElement w in elements)
                    {
                        if (w.Current.Name.Equals(elementName))
                        {
                            visible = !w.Current.IsOffscreen;
                            if (visible)
                                break;
                        }
                    }
                }
                while (elapsedTime <= timeOut && visible == false);
                Logger.logMessage("WaitForTextVisibleInsidePane " + window + "->" + pane + "->" + elementName + " - End Sync");
                Logger.logMessage("------------------------------------------------------------------------------");
                Thread.Sleep(int.Parse(Execution_Speed));
                return visible;
            }
            catch (Exception e)
            {
                Logger.logMessage("WaitForTextVisibleInsidePane " + window + "->" + pane + "->" + elementName + " - Terminated");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************


        public static bool CheckTextExistsInsidePane(TestStack.White.UIItems.Panel pane, string text)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            bool exists = false;

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    if (item.Current.Name.Equals(text))
                    {
                        exists = true;
                        break;
                    }
                }

                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("CheckTextExistsInsidePane " + pane + "->" + text + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                return exists;
            }
            catch (Exception e)
            {
                Logger.logMessage("CheckTextExistsInsidePane " + pane + "->" + text + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.Label> GetMatchingLablesInsideAPane(Window window, TestStack.White.UIItems.Panel pane, string text)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.Label> collection = new List<TestStack.White.UIItems.Label>();

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(System.Windows.Automation.TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    if (item.Current.Name.Equals(text) || item.Current.Name.Contains(text))
                    {
                        TestStack.White.UIItems.Label t = new TestStack.White.UIItems.Label(item, window.ActionListener);
                        collection.Add(t);
                    }
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("GetMatchingLablesInsideAPane " + window + "->" + pane + "->" + text + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return collection;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetMatchingLablesInsideAPane " + window + "->" + pane + "->" + text + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.MenuItems.Menu> GetAllMenuItemsInsideAPane(Window window, TestStack.White.UIItems.Panel pane)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.MenuItems.Menu> collection = new List<TestStack.White.UIItems.MenuItems.Menu>();

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(System.Windows.Automation.TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    TestStack.White.UIItems.MenuItems.Menu t = new TestStack.White.UIItems.MenuItems.Menu(item, window.ActionListener);
                    collection.Add(t);
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("GetAllMenuItemsInsideAPane " + window + "->" + pane + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return collection;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllMenuItemsInsideAPane " + window + "->" + pane + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.ListViewCell> GetAllListItemsInsideAPane(Window window, TestStack.White.UIItems.Panel pane)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.ListViewCell> collection = new List<TestStack.White.UIItems.ListViewCell>();

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(System.Windows.Automation.TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    var t = new TestStack.White.UIItems.ListViewCell(item, window.ActionListener);
                    collection.Add(t);
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("GetAllMenuItemsInsideAPane " + window + "->" + pane + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return collection;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllMenuItemsInsideAPane " + window + "->" + pane + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.TextBox> GetAllEditBoxesInsideAPane(Window window, TestStack.White.UIItems.Panel pane)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.TextBox> editBoxCollection = new List<TestStack.White.UIItems.TextBox>();

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(System.Windows.Automation.TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    TestStack.White.UIItems.TextBox t = new TestStack.White.UIItems.TextBox(item, window.ActionListener);
                    editBoxCollection.Add(t);
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("GetAllEditBoxesInsideAPane " + window + "->" + pane + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return editBoxCollection;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllEditBoxesInsideAPane " + window + "->" + pane + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.Button> GetAllButtonsInsideAPane(Window window, TestStack.White.UIItems.Panel pane)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.Button> collection = new List<TestStack.White.UIItems.Button>();

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(System.Windows.Automation.TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    TestStack.White.UIItems.Button t = new TestStack.White.UIItems.Button(item, window.ActionListener);
                    collection.Add(t);
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("GetAllButtonsInsideAPane " + window + "->" + pane + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return collection;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllButtonsInsideAPane " + window + "->" + pane + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.CheckBox> GetAllCheckBoxesInsideAPane(Window window, TestStack.White.UIItems.Panel pane)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.CheckBox> collection = new List<TestStack.White.UIItems.CheckBox>();

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(System.Windows.Automation.TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    TestStack.White.UIItems.CheckBox t = new TestStack.White.UIItems.CheckBox(item, window.ActionListener);
                    collection.Add(t);
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("GetAllCheckBoxesInsideAPane " + window + "->" + pane + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return collection;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllCheckBoxesInsideAPane " + window + "->" + pane + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static List<TestStack.White.UIItems.Label> GetAllLablesInsideAPane(Window window, TestStack.White.UIItems.Panel pane)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.Label> collection = new List<TestStack.White.UIItems.Label>();

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(System.Windows.Automation.TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    TestStack.White.UIItems.Label t = new TestStack.White.UIItems.Label(item, window.ActionListener);
                    collection.Add(t);
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("GetAllLablesInsideAPane " + window + "->" + pane + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return collection;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllLablesInsideAPane " + window + "->" + pane + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************


        public static List<TestStack.White.UIItems.ListBoxItems.ComboBox> GetAllComboBoxesInsideAPane(Window window, TestStack.White.UIItems.Panel pane)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            List<TestStack.White.UIItems.ListBoxItems.ComboBox> collection = new List<TestStack.White.UIItems.ListBoxItems.ComboBox>();

            try
            {
                PropertyCondition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox);
                AutomationElementCollection elements = pane.AutomationElement.FindAll(System.Windows.Automation.TreeScope.Descendants, condition);

                foreach (AutomationElement item in elements)
                {
                    TestStack.White.UIItems.ListBoxItems.ComboBox t = new TestStack.White.UIItems.ListBoxItems.ComboBox(item, window.ActionListener);
                    collection.Add(t);
                }
                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("GetAllComboBoxesInsideAPane " + window + "->" + pane + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return collection;
            }
            catch (Exception e)
            {
                Logger.logMessage("GetAllComboBoxesInsideAPane " + window + "->" + pane + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        //**************************************************************************************************************************************************************

        public static bool CheckCheckBoxIsSelected(Window win, String checkBoxAutoID)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            try
            {
                TestStack.White.UIItems.CheckBox c = win.Get<TestStack.White.UIItems.CheckBox>(SearchCriteria.ByAutomationId(checkBoxAutoID));
                bool state = c.IsSelected;
                if (state)
                {
                    Logger.logMessage("CheckCheckBoxIsSelected " + win + "->" + checkBoxAutoID + "->" + state + " - Successful");
                    return true;
                }
                else
                {
                    Logger.logMessage("CheckCheckBoxIsSelected " + win + "->" + checkBoxAutoID + "->" + state + " - Successful");
                    return false;
                }
            }
            catch (Exception e)
            {
                Logger.logMessage("SelectCheckBox " + win + "->" + checkBoxAutoID + "->" + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

    }
}

