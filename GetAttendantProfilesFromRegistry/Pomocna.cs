/*
   Microsoft SQL Server Integration Services Script Task
   Write scripts using Microsoft Visual C# 2008.
   The ScriptMain is the entry point class of the script.
*/

using System;
using System.Data;
using Microsoft.Win32;
using System.Data.SqlClient;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;


namespace ST_db62adb153074a7eb8d1194bd8c6413d.csproj
{
    [Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
    public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
    {

        #region VSTA generated code
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

        /*
		The execution engine calls this method when the task executes.
		To access the object model, use the Dts property. Connections, variables, events,
		and logging features are available as members of the Dts property as shown in the following examples.

		To reference a variable, call Dts.Variables["MyCaseSensitiveVariableName"].Value;
		To post a log entry, call Dts.Log("This is my log text", 999, null);
		To fire an event, call Dts.Events.FireInformation(99, "test", "hit the help message", "", 0, true);

		To use the connections collection use something like the following:
		ConnectionManager cm = Dts.Connections.Add("OLEDB");
		cm.ConnectionString = "Data Source=localhost;Initial Catalog=AdventureWorks;Provider=SQLNCLI10;Integrated Security=SSPI;Auto Translate=False;";

		Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
		
		To open Help, press F1.
	*/

        public void Main()
        {
            try
            {
                Dts.Log("Pocetak", 999, null);
                string employeePath = Dts.Variables["User::EmployeePath"].Value.ToString();
                RegistryKey rkey = Registry.LocalMachine.CreateSubKey(employeePath);


                Dts.Log("Pre deklarisanja konekcije", 999, null);
                //SqlConnection mySqlConn = new SqlConnection(Dts.Connections["ContactCenterStaging"].ConnectionString);
                //MessageBox.Show(Dts.Variables["User::connContactCenterStaging"].Value.ToString());
                //MessageBox.Show(Dts.Connections["ContactCenterStaging"].ConnectionString);
                // SqlConnection mySqlConn = (SqlConnection)(Dts.Connections["ContactCenterStaging"].AcquireConnection(null));
                SqlConnection mySqlConn = new SqlConnection(Dts.Variables["User::connADOContactCenterStaging"].Value.ToString());

                Dts.Log("Posle kreiranja konekcije", 999, null);




                String[] names = rkey.GetSubKeyNames();

                MessageBox.Show(employeePath);

                // Print the contents of the array to the console.
                foreach (String s in names)
                {

                    MessageBox.Show("1");
                    string user = s.ToString();
                    RegistryKey registry = rkey.OpenSubKey(s);

                    if (registry != null)
                    {

                        string employeeFirstName = registry.GetValue("givenName") != null ? ((string[])(registry.GetValue("givenName")))[0] : null;
                        string employeeLastName = registry.GetValue("surName") != null ? ((string[])(registry.GetValue("surName")))[0] : null;
                        string employeeAddress = registry.GetValue("streetAddress") != null ? ((string[])(registry.GetValue("streetAddress")))[0] : null;
                        string employeeCity = registry.GetValue("localityName") != null ? ((string[])(registry.GetValue("localityName")))[0] : null;
                        string employeeZipCode = registry.GetValue("postalCode") != null ? ((string[])(registry.GetValue("postalCode")))[0] : null;
                        string employeeCountry = registry.GetValue("country") != null ? ((string[])(registry.GetValue("country")))[0] : null;
                        string employeeTitle = registry.GetValue("title") != null ? ((string[])(registry.GetValue("title")))[0] : null;
                        string employeeDepartment = registry.GetValue("departmentName") != null ? ((string[])(registry.GetValue("departmentName")))[0] : null;
                        string employeeCompany = registry.GetValue("companyName") != null ? ((string[])(registry.GetValue("companyName")))[0] : null;
                        string employeeDisplayname = registry.GetValue("displayName") != null ? ((string[])(registry.GetValue("displayName")))[0] : null;
                        string employeeHomePhone1 = registry.GetValue("phoneHome1") != null ? ((string[])(registry.GetValue("phoneHome1")))[0] : null;
                        string employeeHomePhone2 = registry.GetValue("phoneHome2") != null ? ((string[])(registry.GetValue("phoneHome2")))[0] : null;
                        string employeeBusinessPhone1 = registry.GetValue("phoneBusiness1") != null ? ((string[])(registry.GetValue("phoneBusiness1")))[0] : null;
                        string employeeBusinessPhone2 = registry.GetValue("phoneBusiness2") != null ? ((string[])(registry.GetValue("phoneBusiness2")))[0] : null;
                        string employeeMobilePhone = registry.GetValue("phoneMobile") != null ? ((string[])(registry.GetValue("phoneMobile")))[0] : null;
                        string employeeFax = registry.GetValue("phonePrimaryFax") != null ? ((string[])(registry.GetValue("phonePrimaryFax")))[0] : null;
                        string employeeExtension = registry.GetValue("Extension") != null ? ((string[])(registry.GetValue("Extension")))[0] : null;
                        string employeeMailBoxUser = registry.GetValue("emailAddress") != null ? ((string[])(registry.GetValue("emailAddress")))[0] : null;
                        string employeePreferredLanguage = registry.GetValue("PreferredLanguage") != null ? ((string[])(registry.GetValue("PreferredLanguage")))[0] : null;
                        string employeeDefaultWorkStation = registry.GetValue("DefaultWorkstation") != null ? ((string[])(registry.GetValue("DefaultWorkstation")))[0] : null;
                        string employeeNTDomainUser = registry.GetValue("NTDomainUser") != null ? ((string[])(registry.GetValue("NTDomainUser")))[0] : null;

                        //SqlCommand myCommand = new SqlCommand("INSERT INTO dbo.EmployeeCurrent (EmployeeUserName, EmployeeFirstName,EmployeeLastName,EmployeeAddress,EmployeeCity,EmployeeZipCode,EmployeeCountry,EmployeeTitle,EmployeeDepartment,EmployeeCompany,EmployeeDisplayname,EmployeeHomePhone1,EmployeeHomePhone2,EmployeeBusinessPhone1,EmployeeBusinessPhone2,EmployeeMobilePhone,EmployeeFax,EmployeeExtension,EmployeeMailBoxUser,EmployeePreferredLanguage,EmployeeDefaultWorkStation,EmployeeNTDomainUser) Values ('" +
                        // user + "','" + givenName + "','" + surName + "','" + streetAddress + "','" + localityName + "','" + postalCode + "','" + country + "','" + title + "','" + departmentName + "','" + companyName + "','" + displayName + "','" + phoneHome1 + "','" + phoneHome2 + "','" + phoneBusiness1 + "','" + phoneBusiness2 + "','" + phoneMobile + "','" + phonePrimaryFax + "','" + Extension + "','" + emailAddress + "','" + PreferredLanguage + "','" + DefaultWorkstation + "','" + NTDomainUser + "')", mySqlConn);


                        String[] skills = (String[])registry.GetValue("Skills");

                        if (skills != null)
                        {
                            foreach (String g in skills)
                            {

                                SqlCommand myCommandSkill = new SqlCommand("dbo.insEmployeeSkillStage_Insert", mySqlConn);
                                myCommandSkill.CommandType = CommandType.StoredProcedure;

                                SqlParameter paramskill1 = new SqlParameter("@pEmployeeUserName", SqlDbType.NVarChar, 50);
                                paramskill1.Value = user;

                                SqlParameter paramskill2 = new SqlParameter("@pSkill", SqlDbType.NVarChar, 250);
                                paramskill2.Value = g;

                                myCommandSkill.Parameters.Add(paramskill1);
                                myCommandSkill.Parameters.Add(paramskill2);

                                mySqlConn.Open();

                                myCommandSkill.ExecuteNonQuery();

                                mySqlConn.Close();


                            }
                        }


                        SqlCommand myCommand = new SqlCommand("insEmployeeCurrent_Insert", mySqlConn);
                        myCommand.CommandType = CommandType.StoredProcedure;

                        // , , , @, @pEmployeeZipCode, @pEmployeeCountry, @pEmployeeTitle, @pEmployeeDepartment, @pEmployeeCompany, @pEmployeeUserName, @pEmployeeDisplayName, @pEmployeeHomePhone1, @pEmployeeHomePhone2, @pEmployeeBusinessPhone1, @pEmployeeBusinessPhone2, @pEmployeeMobilePhone, @pEmployeeFax, @pEmployeeExtension, @pEmployeeMailBoxUser, @pEmployeePreferredLanguage, @pEmployeeDefaultWorkStation, @pEmployeeNTDomainUser
                        SqlParameter param1 = new SqlParameter("@pEmployeeFirstName", SqlDbType.NVarChar, 80);
                        if (employeeFirstName != null)
                            param1.Value = employeeFirstName;
                        else
                            param1.Value = DBNull.Value;
                        myCommand.Parameters.Add(param1);

                        SqlParameter param2 = new SqlParameter("@pEmployeeLastName", SqlDbType.NVarChar, 80);
                        if (employeeLastName != null)
                            param2.Value = employeeLastName;
                        else
                            param2.Value = DBNull.Value;
                        myCommand.Parameters.Add(param2);

                        SqlParameter param3 = new SqlParameter("@pEmployeeAddress", SqlDbType.NVarChar, 80);
                        if (employeeAddress != null)
                            param3.Value = employeeAddress;
                        else
                            param3.Value = DBNull.Value;
                        myCommand.Parameters.Add(param3);

                        SqlParameter param4 = new SqlParameter("@pEmployeeCity", SqlDbType.NVarChar, 30);
                        if (employeeCity != null)
                            param4.Value = employeeCity;
                        else
                            param4.Value = DBNull.Value;
                        myCommand.Parameters.Add(param4);

                        SqlParameter param5 = new SqlParameter("@pEmployeeZipCode", SqlDbType.NVarChar, 15);
                        if (employeeZipCode != null)
                            param5.Value = employeeZipCode;
                        else
                            param5.Value = DBNull.Value;
                        myCommand.Parameters.Add(param5);

                        SqlParameter param6 = new SqlParameter("@pEmployeeCountry", SqlDbType.NVarChar, 50);
                        if (employeeCountry != null)
                            param6.Value = employeeCountry;
                        else
                            param6.Value = DBNull.Value;
                        myCommand.Parameters.Add(param6);

                        SqlParameter param7 = new SqlParameter("@pEmployeeTitle", SqlDbType.NVarChar, 20);
                        if (employeeTitle != null)
                            param7.Value = employeeTitle;
                        else
                            param7.Value = DBNull.Value;
                        myCommand.Parameters.Add(param7);

                        SqlParameter param8 = new SqlParameter("@pEmployeeDepartment", SqlDbType.NVarChar, 50);
                        if (employeeDepartment != null)
                            param8.Value = employeeDepartment;
                        else
                            param8.Value = DBNull.Value;
                        myCommand.Parameters.Add(param8);

                        SqlParameter param9 = new SqlParameter("@pEmployeeCompany", SqlDbType.NVarChar, 50);
                        if (employeeCompany != null)
                            param9.Value = employeeCompany;
                        else
                            param9.Value = DBNull.Value;
                        myCommand.Parameters.Add(param9);

                        SqlParameter param10 = new SqlParameter("@pEmployeeUserName", SqlDbType.NVarChar, 50);
                        if (user != null)
                            param10.Value = user;
                        else
                            param10.Value = DBNull.Value;
                        myCommand.Parameters.Add(param10);

                        SqlParameter param11 = new SqlParameter("@pEmployeeDisplayName", SqlDbType.NVarChar, 50);
                        if (employeeDisplayname != null)
                            param11.Value = employeeDisplayname;
                        else
                            param11.Value = DBNull.Value;
                        myCommand.Parameters.Add(param11);

                        SqlParameter param12 = new SqlParameter("@pEmployeeHomePhone1", SqlDbType.NVarChar, 25);
                        if (employeeHomePhone1 != null)
                            param12.Value = employeeHomePhone1;
                        else
                            param12.Value = DBNull.Value;
                        myCommand.Parameters.Add(param12);

                        SqlParameter param13 = new SqlParameter("@pEmployeeHomePhone2", SqlDbType.NVarChar, 25);
                        if (employeeHomePhone2 != null)
                            param13.Value = employeeHomePhone2;
                        else
                            param13.Value = DBNull.Value;
                        myCommand.Parameters.Add(param13);

                        SqlParameter param14 = new SqlParameter("@pEmployeeBusinessPhone1", SqlDbType.NVarChar, 25);
                        if (employeeBusinessPhone1 != null)
                            param14.Value = employeeBusinessPhone1;
                        else
                            param14.Value = DBNull.Value;
                        myCommand.Parameters.Add(param14);

                        SqlParameter param15 = new SqlParameter("@pEmployeeBusinessPhone2", SqlDbType.NVarChar, 25);
                        if (employeeBusinessPhone2 != null)
                            param15.Value = employeeBusinessPhone2;
                        else
                            param15.Value = DBNull.Value;
                        myCommand.Parameters.Add(param15);

                        SqlParameter param16 = new SqlParameter("@pEmployeeMobilePhone", SqlDbType.NVarChar, 30);
                        if (employeeMobilePhone != null)
                            param16.Value = employeeMobilePhone;
                        else
                            param16.Value = DBNull.Value;
                        myCommand.Parameters.Add(param16);

                        SqlParameter param17 = new SqlParameter("@pEmployeeFax", SqlDbType.NVarChar, 30);
                        if (employeeFax != null)
                            param17.Value = employeeFax;
                        else
                            param17.Value = DBNull.Value;
                        myCommand.Parameters.Add(param17);


                        SqlParameter param18 = new SqlParameter("@pEmployeeExtension", SqlDbType.NVarChar, 5);
                        if (employeeExtension != null)
                            param18.Value = employeeExtension;
                        else
                            param18.Value = DBNull.Value;
                        myCommand.Parameters.Add(param18);

                        SqlParameter param19 = new SqlParameter("@pEmployeeMailBoxUser", SqlDbType.NVarChar, 200);
                        if (employeeMailBoxUser != null)
                            param19.Value = employeeMailBoxUser;
                        else
                            param19.Value = DBNull.Value;
                        myCommand.Parameters.Add(param19);


                        SqlParameter param20 = new SqlParameter("@pEmployeePreferredLanguage", SqlDbType.NVarChar, 20);
                        if (employeePreferredLanguage != null)
                            param20.Value = employeePreferredLanguage;
                        else
                            param20.Value = DBNull.Value;
                        myCommand.Parameters.Add(param20);

                        SqlParameter param21 = new SqlParameter("@pEmployeeDefaultWorkStation", SqlDbType.NVarChar, 10);
                        if (employeeDefaultWorkStation != null)
                            param21.Value = employeeDefaultWorkStation;
                        else
                            param21.Value = DBNull.Value;
                        myCommand.Parameters.Add(param21);

                        SqlParameter param22 = new SqlParameter("@pEmployeeNTDomainUser", SqlDbType.NVarChar, 20);
                        if (employeeNTDomainUser != null)
                            param22.Value = employeeNTDomainUser;
                        else
                            param22.Value = DBNull.Value;
                        myCommand.Parameters.Add(param22);



                        Dts.Log("Pre otvaranja konekcije", 999, null);

                        mySqlConn.Open();

                        myCommand.ExecuteNonQuery();

                        mySqlConn.Close();
                    }
                }
            }
            catch (Exception e)
            {

                throw e;

            }

            Dts.TaskResult = (int)ScriptResults.Success;
        }
    }
}