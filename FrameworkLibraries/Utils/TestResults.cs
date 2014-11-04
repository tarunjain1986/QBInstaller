using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FrameworkLibraries;
using FrameworkLibraries.EntityFramework;

namespace FrameworkLibraries.Utils
{
    public static class TestResults
    {
        public static Property conf = Property.GetPropertyInstance();

        public static String GetResultUpdateQuery(String sModule, String sTestName, String sStatus, String sFailReason, String sFailCategory)
	    {
            Random random = new Random();
		    String sEnvironment = conf.get("Environment");
            String sReleaseName = conf.get("ReleaseName");
            String sProducVersion = conf.get("ProductVersion");
            String sRunName = conf.get("RunName");
            int iID = random.Next(123456, 999999);
            DateTime dExecutionDate = DateTime.Today;
		    String insertQuery =  
		    "insert into Test_Results" + 
		    "(ENVIRONMENT, RELEASE_NAME, "+
		    "PRODUCT_VERSION, "+
		    "MODULE_NAME, "+
		    "TEST_NAME, "+
		    "TEST_RESULT, "+
		    "FAILURE_REASON, "+
		    "FAILURE_CATEGORY, "+
		    "EXECUTION_DATE, "+
            "ID, " +
		    "RUN_NAME)" +
		    "VALUES('"+sEnvironment+"', "+
            "'" +sReleaseName+ "', " +
            "'" +sProducVersion+ "', " +
		    "'"+sModule+"', "+
		    "'"+sTestName+"', "+
		    "'"+sStatus+"', "+
		    "'"+sFailReason+"', "+
		    "'"+sFailCategory+"', "+
            "'" + dExecutionDate + "', " +
            "'" + iID + "', " +
		    "'"+sRunName+"')";
		
             return insertQuery;
    	}

        public static void CheckResultRecordExists(String sTestName, String sStatus)
        {
            try
            {
                using (SQLCompactDBEntities entity = new SQLCompactDBEntities())
                {
                    string runName = conf.get("RunName");
                    var recordCount = entity.Test_Results.Count(a => a.Test_Name.Equals(sTestName) && a.Run_Name.Equals(runName) && a.Execution_Date.Equals(DateTime.Today) && a.Test_Result.Equals(sStatus));
                    if (recordCount != 0)
                    {
                        entity.Test_Results.SqlQuery("delete from Test_Results where TEST_NAME like '" + sTestName + "' and EXECUTION_DATE like " + DateTime.Today + " and TEST_RESULT like '" + sStatus + "' and RUN_Name like '" + conf.get("RunName") + "'");
                        DbSqlQuery<Test_Results> db = entity.Test_Results.SqlQuery("delete from Test_Results where TEST_NAME like '" + sTestName + "' and TEST_RESULT like '" + sStatus + "' and RUN_Name like '" + conf.get("RunName") + "'");
                        foreach (Test_Results check in db)
                        {
                        }
                        IEnumerator<Test_Results> enu = db.GetEnumerator();
                        var c = enu.Current;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        public static void UpdateResult(String query)
        {
            try
            {
                using (SQLCompactDBEntities entity = new SQLCompactDBEntities())
                {
                    DbSqlQuery<Test_Results> db = entity.Test_Results.SqlQuery(query);
                    foreach (Test_Results check in db)
                    {
                    }
                    IEnumerator<Test_Results> enu = db.GetEnumerator();
                    var c = enu.Current;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static void GetTestResult(String testName, String moduleName, String exception, String category)
        {
            if (exception.Equals("Null"))
            {
                TestResults.CheckResultRecordExists(testName, "Pass");
                String sQuery = TestResults.GetResultUpdateQuery(moduleName, testName, "Pass", null, null);
                TestResults.UpdateResult(sQuery);
            }
            else
            {
                TestResults.CheckResultRecordExists(testName, "Fail");
                String sQuery = TestResults.GetResultUpdateQuery(moduleName, testName, "Fail", exception, "Failure");
                TestResults.UpdateResult(sQuery);
            }
        }

        public static String TrimExceptionMessage(String excep)
        {

            var charsToRemove = new string[] { "@", ",", ".", ";", "'" };
            foreach (var c in charsToRemove)
            {
                excep = excep.Replace(c, string.Empty);
            }
            return excep;
        }

        public static void UpdateRecord(String query)
        {
            try
            {
                using (SQLCompactDBEntities entity = new SQLCompactDBEntities())
                {
                    DbSqlQuery<Siebel_TestData> db = entity.Siebel_TestData.SqlQuery(query);
                    foreach (Siebel_TestData check in db)
                    {
                    }
                    IEnumerator<Siebel_TestData> enu = db.GetEnumerator();
                    var c = enu.Current;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

    }
}
