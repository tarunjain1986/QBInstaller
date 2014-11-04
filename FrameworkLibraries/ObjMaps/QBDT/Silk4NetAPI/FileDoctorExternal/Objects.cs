using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameworkLibraries.ObjMaps.QBDT.Silk4NetAPI.FileDoctorExternal
{
    public class Objects
    {
        public static String FileDoctorExternal_FormsWindow = "//FormsWindow[@Caption='Intuit QuickBooks File Doctor']";
        public static String MainWindow_Label_2 = "//Label[@automationId='label16']";
        public static String BrowseFile_TextField = "//TextField[@automationId='CompanyFileTextBox']";
        public static String LocalHostNDT_RadioList = "//RadioList[@automationId='LocalHostNDTButton']";
        public static String FullNDT_RadioList = "//RadioList[@automationId='FullNDTRadioButton']";
        public static String Fix6130Error_CheckBox = "//CheckBox[@automationId='Fix6130CheckBox']";
        public static String FixDataSync_CheckBox = "//CheckBox[@automationId='FixDataSyncCheckBox']";
        public static String DiagnoseFile_PushButton = "//PushButton[@caption='Diagnose File']";
        public static String CheckConnectivity_PushButton = "//PushButton[@caption='Check Connectivity']";
        public static String Cancel_PushButton = "//PushButton[@caption='Cancel']";
        public static String Close_PushButton = "//PushButton[@caption='Close']";
        public static String OK_PushButton = "//PushButton[@caption='Ok']";
        public static String UserName_TextField = "//TextField[@automationId='UserNameText']";
        public static String Password_TextField = "//TextField[@automationId='PasswordText']";
        public static String Password_PushButton = "//PushButton[@automationId='LoginButton']";
        public static String LoginWaitMessage_Label = "//Label[@Text='Trying to log in to the company file...']";
        public static String NetworkAccess_Label = "//Label[@caption='Network Access']";
        public static String NetworkAccess_WorkStation_RadioList = "//RadioList[@caption='Workstation']";
        public static String NetworkAccess_Server_RadioList = "//RadioList[@caption='Server']";
        public static String NetworkAccess_Next_PushButton = "//PushButton[@caption='Next']";
        public static String CheckConnectivityWaitMessage_Label = "//Label[@Text='Checking connectivity set up for hosting QuickBooks...']";
        public static String NetworkResult_Group = "//Group[@automationId='NetworkResultPanel']";
        public static String FileResult_Group = "//Group[@automationId='NoSuccessPanel']";
        public static String FolderSharing_Yes_RadioList = "//RadioList[@caption='Yes']";
        public static String FolderSharing_No_RadioList = "//RadioList[@caption='No']";
        public static String FolderSharing_Next_PushButton = "//PushButton[@caption='Next']";
    }
}
