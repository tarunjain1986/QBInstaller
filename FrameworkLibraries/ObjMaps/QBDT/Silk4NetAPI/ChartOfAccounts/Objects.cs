using System;

namespace FrameworkLibraries.ObjMaps.QBDT.Silk4NetAPI.ChartOfAccounts
{
    public class Objects
    {
        public static string Accounts_Control = "/Window//Control[@caption='&Account']";
        public static string Menu_Control = "/Window//Control[@windowClassName='XTPSkinManagerMenu']";
        public static string Continue_Control = "/Window//Control[@caption='Con&tinue']";
        public static string Bank_Account_Name = "/Window//TextField[3]";
        public static string Bank_Control = "/Window//Control[@caption='Ban&k']";
        public static string Bank_SubAccount_Checkbox_Control = "/Window//Control[@caption='&Subaccount of']";
        public static string Bank_SubAccount_Text_Control = "/Window//TextField[3]";
        public static string Save_Close_Control = "/Window//Control[@caption='S&ave && Close']";
        public static string Bank_Description = "/Window//TextField[4]";
        public static string Bank_Account_Number = "/Window//TextField[3]";
        public static string Bank_Routing_Number = "/Window//TextField[3]";
        public static string Bank_Enter_Opening_Balance_Control = "/Window//Control[@caption='Enter Openin&g Balance...']";
        public static string Bank_Enter_Opening_Balance_Textfield = "/Window//Window[@caption='Enter Opening Balance: Bank Account']//TextField";
        public static string Bank_Enter_Opening_Balance_Date_Textfield = "/Window//Window[@caption='Enter Opening Balance: Bank Account']//TextField";
        public static string Bank_Cheque_Number = "/Window//TextField[3]";
        public static string Bank_Enter_Opening_Balance_Ok_Control = "/Window//Window[@caption='Enter Opening Balance: Bank Account']//Control[@caption='OK']";
        public static string Bank_Order_Cheque_Control = "/Window//Control[@caption='Order checks I can print from QuickBooks']";
        public static string Bank_Form_Control = "/Window//Control[@windowClassName='QuickBooksSubForm'][2]";
        public static string Setup_Bank_Feed_Window = "/Window//Window[@caption='Set Up Bank Feed']";
        public static string Setup_Bank_Feed_Window_No = "/Window//Window[@caption='Set Up Bank Feed']//Control[@caption='No']";
        public static string Order_Supplies_Window = "/Window//Window[@caption='Order Supplies']";
        public static string Chart_Of_Accounts_Window = "/Window//Window[@caption='Chart of Accounts']";
        public static string Expense_Account_Control = "/Window//Control[@caption='E&xpense']";
        public static string Income_Account_Control = "/Window//Control[@caption='&Income']";
    }
}