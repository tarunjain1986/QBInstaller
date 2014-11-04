using System;

namespace FrameworkLibraries.ObjMaps.QBDT.Silk4NetAPI.Common
{
    public class Objects
    {
        public static string Main_Window = "/Window";
        public static string Reg_Remind_Later = "/Window//Control[@caption='Remind Me &Later']";
        public static string Main_Menu = "/Window//Control[@windowClassName='XTPSkinManagerMenu']";
        public static string Sub_Menu = "/Window//Control[@windowClassName='XTPSkinManagerMenu'][2]";
        public static string Recommendation_Window = "/Window//Window[@caption='QuickBooks Recommendation']";
        public static string Recommendation_Payroll_Window = "/Window//Window[@caption='QuickBooks Recommendation']//Window[@caption='Intuit QuickBooks Payroll']";
        public static string Text_Controls = "/Window//TextField";
        public static string SideBar_Menu_BankFeeds = "//WPFButton[@caption='Bank Feeds']";
        public static string Vertical_Scroll_Bar = "//VerticalScrollBar";
    }
}
