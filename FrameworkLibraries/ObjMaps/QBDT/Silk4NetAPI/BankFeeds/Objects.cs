using System;


namespace FrameworkLibraries.ObjMaps.QBDT.Silk4NetAPI.BankFeeds
{
    public class Objects
    {
        public static String File_Import_Dialog = "Window//Dialog[@caption='Open Online Data File']";
        public static String File_Import_TextField = "Window//Dialog[@caption='Open Online Data File']//TextField[@caption='File name:']";
        public static String File_Import_Open_Button = "Window//Dialog[@caption='Open Online Data File']//PushButton[@caption='Open']";
        public static String Web_Connect_Dialog = "Window//Dialog[@caption='QuickBooks Web Connect']";
        public static String Web_Connect_Ok_Button = "Window//Dialog[@caption='QuickBooks Web Connect']//PushButton[@caption='OK']";
        public static String AccountInformation_BankAccounts_ListViewItem = "//WPFUserControl[@caption='_Add Account']//WPFListView[@automationId='AccountListView']/WPFListViewItem";
        public static String Transaction_List = "//WPFUserControl[@caption='_Add Account']//WPFUserControl[@automationId='RightPaneView']/WPFButton[@caption='Transaction List']";
        public static String Transaction_List_Grid = "//WPFUserControl[@caption='Finish Later']//WPFToolkitDataGrid[@name='dGridTransactionDetails']";
        public static String Bank_Feeds_Window = "/Window//Window[@caption='Bank Feeds']";
        public static String TRL_Grid_Checkboxes = "//WPFUserControl[@caption='Finish Later']//WPFToolkitDataGrid//WPFCheckBox";
        public static String TRL_Grid_Payee_WPFComboBox = "//WPFUserControl[@caption='Finish Later']//WPFComboBox[@automationId='PayeeWPF']";
        public static String TRL_Grid_Account_ComboBox = "//WPFUserControl[@caption='Finish Later']//FormsHost[@automationId='embeddedComboQBHost']//ComboBox[@className='OLBUICustomControlLibrary.ComboBoxQB']";
        public static String TRL_Grid_Action_WPFComboBox = "//WPFUserControl[@caption='Finish Later']//WPFComboBox[@automationId='ActionCombo']";
        public static String TRL_Grid_Action_QuickAdd_WPFComboBoxItem = "//WPFUserControl[@caption='Finish Later']//WPFComboBox[@automationId='ActionCombo']/WPFComboBoxItem[@caption='Quick Add']";
        public static String Past_Transaction_Yes_Control = "//Control[@caption='Past Transactions']//Control[@caption='&Yes']";
        public static String Rule_Creation_Ok_Button = "//WPFButton[@automationID='btnSave']";
        public static String Rule_Creation_Action_WPFComboBox = "//WPFComboBox[@automationId='ActionCombo']";
        public static String Rule_Creation_IgnoreRule_WPFComboBoxItem = "//WPFComboBoxItem[@automationId='Ignore Rule']";
        public static String TRL_Need_Your_Review_TxtBlock = "//WPFUserControl[@caption='Finish Later']//WPFTextBlock[@caption='NEED YOUR REVIEW']";
        public static String TRL_Changed_By_Rules_TxtBlock = "//WPFUserControl[@caption='Finish Later']//WPFTextBlock[@caption='CHANGED BY RULES']";
        public static String TRL_Add_TxtBlock = "//WPFUserControl[@caption='Finish Later']//WPFTextBlock[@caption='ADD']";
        public static String TRL_Status_WPFComboBox = "//WPFUserControl[@caption='Finish Later']//WPFComboBox[@automationId='StatusCombo']";
        public static String TRL_Type_WPFComboBox = "//WPFUserControl[@caption='Finish Later']//WPFComboBox[2]";
        public static String TRL_Rules_Label = "//WPFUserControl[@caption='Finish Later']//WPFToggleButton[@automationId='OlbRulesButton']/WPFButton[@automationId='PART_Button']/WPFLabel[@caption='Rules']";
        public static String TRL_Batch_Actions_WPFButton = "//WPFUserControl[@caption='Finish Later']//WPFButton[@caption='Batch Actions']";
        public static String TRL_Batch_Actions_WPFToggleButton = "//WPFUserControl[@caption='Finish Later']//WPFToggleButton[@caption='Batch Actions']";
        public static String TRL_Batch_Actions_ContextMenu = "//WPFUserControl[@caption='Finish Later']//WPFContextMenu";
        public static String TRL_Record_Control = "//Control[@caption='Recor&d']";
        public static String TRL_Restore_Control = "//Control[@caption='Restore']";
        public static String TRL_OneLine_Control = "//Control[@caption='&1-Line']";
        public static String TRL_Batch_Actions_Add_WPFMenuItem = "//ContextMenu//WPFMenuItem[@caption='Add/Approve']";
        public static String Rule_Creation_Window = "//WPFWindow[@caption='Rule Creation']";

        //Direct connect WPF objects
        public static String Setup_SearchInput_WPFTextBox = "WPFWindow[@caption='Bank Feed Setup']//WPFTextBox[@name='SearchInput']";
        public static String Setup_Banks_WPFListView = "WPFWindow[@caption='Bank Feed Setup']//WPFListView[@TabIndex='2']";
        public static String Setup_Continue_WPFButton = "WPFWindow[@caption='Bank Feed Setup']//WPFButton[@caption='Continue']";
        public static String Setup_CustomerID_WPFTextBox = "WPFWindow[@caption='Bank Feed Setup']//WPFTextBox[@name='UserIdBox']";
        public static String Setup_PIN_WPFPasswordBox = "WPFWindow[@caption='Bank Feed Setup']//WPFPasswordBox[@automationId='PasswordBox']";
        public static String Setup_Connect_WPFButton = "WPFWindow[@caption='Bank Feed Setup']//WPFButton[@caption='Connect']";
        public static String Setup_LinkYourAccount_WPFToolkitDataGridCell = "WPFWindow[@caption='Bank Feed Setup']//WPFToolkitDataGridCell";
        public static String Setup_LinkYourAccount_WPFComboBox = "WPFWindow[@caption='Bank Feed Setup']//WPFComboBox[@Text='Select existing or create new']";
        public static String Setup_Close_WPFButton = "WPFWindow[@caption='Bank Feed Setup']//WPFButton[@caption='Close']";
        public static String Setup_Window_WPFUserControl = "//WPFUserControl[@className='LoginFIView']";

    }
}
