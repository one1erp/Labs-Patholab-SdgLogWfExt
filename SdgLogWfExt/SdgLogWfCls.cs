using System;
using System.IO;
using System.Runtime.InteropServices;
using LSEXT;
using LSSERVICEPROVIDERLib;
using Oracle.ManagedDataAccess.Client;

namespace SdgLogWfExt
{
    [ComVisible(true)]
    [ProgId("SdgLogWfExt.SdgLogWfCls")]
    public class SdgLogWfCls : IWorkflowExtension
    {


        private bool debug;
        private string _recordId;
        private string _tableName;
        private string _sql;
        private double _sessionId;
        private string _eventName;

        OracleConnection oraCon = null;
        OracleCommand cmd = null;
        private string _entityName;


        public void Execute(ref LSExtensionParameters Parameters)
        {

            INautilusServiceProvider _sp;

            try
            {
                _sp = Parameters["SERVICE_PROVIDER"];
                _tableName = Parameters["TABLE_NAME"];
        

                string role = Parameters["ROLE_NAME"];
                debug = (role.ToUpper() == "DEBUG");
                long wnid = Parameters["WORKFLOW_NODE_ID"];



                 _sessionId = Parameters["SESSION_ID"];   	
                  






                //*******************************************************
              //  System.Windows.Forms.MessageBox.Show("Try to get eevnt name");


                INautilusDBConnection _ntlsCon = null;
                if (_sp != null)
                {
                    _ntlsCon = _sp.QueryServiceProvider("DBConnection") as NautilusDBConnection;
                }
                else
                {
                    return;
                }
                if (_ntlsCon != null)
                {

                    oraCon = GetConnection(_ntlsCon);

                    cmd = oraCon.CreateCommand();
                }


                var records = Parameters["RECORDS"];

                
                records.MoveLast();
           //     _sessionId = _ntlsCon.GetSessionId();
                
                //Get id of selected value
                _recordId = records.Fields[_tableName + "_ID"].Value.ToString();

                _entityName = records.Fields["NAME"].Value;

                _eventName = GetEventName(wnid);

                Add2Log();

             





            }

            catch (Exception ex)
            {
                MyLog("Error " + ex.Message);

            }
            finally
            {


                if (cmd != null)
                {
                    cmd.Dispose();
                    cmd = null;
                }
                if (oraCon != null)
                {
                    oraCon.Close();
                    oraCon = null;
                }


            }
        }

        private string GetEventName(long wnid)
        {
            _eventName = "";

            _sql = string.Format("SELECT name FROM lims_sys.workflow_node wp where wp.workflow_node_id=( SELECT wn.parent_id FROM lims_sys.workflow_node wn WHERE workflow_node_id={0} ) and wp.WORKFLOW_NODE_TYPE_ID=(select nt.WORKFLOW_NODE_TYPE_ID from lims_sys.WORKFLOW_NODE_TYPE nt where nt.name='Event')", wnid);
            MyLog(_sql);
            cmd = new OracleCommand(_sql, oraCon);
            var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                _eventName = reader[0].ToString();


            }

            return _eventName;


        }


        public void Add2Log()
        {
            string sdgId4Log = "";

            sdgId4Log = _tableName != "SDG" ? GetSdgId(_tableName, _recordId) : _recordId;

            int temp;

            if (sdgId4Log != null && int.TryParse(sdgId4Log, out temp))
            {

                try
                {
                    string sql = "begin lims.Insert_To_Sdg_Log  (" + sdgId4Log + ", '" + _eventName + "'," +
                                 _sessionId.ToString() + ", '" + _eventName + " "+ _entityName + "' ); end;";
                    cmd = new OracleCommand(sql, oraCon);

                    var res =
                        cmd.ExecuteNonQuery();

                }
                catch (Exception e)
                {
                    //Logger.WriteLogFile(e);


                }

            }
        }


        public string GetSdgId(string tableName, string entityId)
        {

            try
            {

                String sql = "";
                switch (_tableName)
                {
                    case "SAMPLE":
                        sql = "SELECT SDG_ID FROM  lims_sys.Sample where sample_id='" + entityId + "'";
                        break;
                    case "ALIQUOT":
                        sql = " SELECT Sample.SDG_ID FROM  lims_sys.Sample where lims_sys. sample.sample_id in(SELECT  lims_sys.aliquot.sample_id FROM  lims_sys.aliquot where  lims_sys.aliquot.aliquot_id='" + entityId + "')";
                        break;
                    case "TEST":
                        sql = "SELECT SDG_ID FROM  lims_sys.Sample where lims_sys.Sample.sample_id in(SELECT lims_sys.aliquot.sample_id FROM  lims_sys.aliquot where aliquot_id in (SELECT lims_sys.test.aliquot_id FROM  lims_sys.test where lims_sys.test.test_id ='" + entityId + "))'";
                        break;
                    default:
                        sql = "";
                        break;
                }
                if (!string.IsNullOrEmpty(sql))
                {
                    cmd = new OracleCommand(sql, oraCon);

                    var res =
                        cmd.ExecuteScalar();
                    if (res != null)
                    {

                        var id = res.ToString();
                        return id;
                    }
                }
            }
            catch (Exception e)
            {
                //   Logger.WriteLogFile(e);
                return null;
            }
            return null;

        }


        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {
            OracleConnection connection = null;
            if (ntlsCon != null)
            {
                //initialize variables
                string rolecommand;
                //try catch block
                try
                {

                    string connectionString;
                    string server = ntlsCon.GetServerName();
                    string user = ntlsCon.GetUsername();
                    string password = ntlsCon.GetPassword();

                    connectionString =
                        string.Format("Data Source={0};User ID={1};Password={2};Connection Timeout=60", server, user, password);

                    var username = ntlsCon.GetUsername();
                    if (string.IsNullOrEmpty(username))
                    {
                        var serverDetails = ntlsCon.GetServerDetails();
                        connectionString = "User Id=/;Data Source=" + serverDetails + ";";
                    }


                    //create connection
                    connection = new OracleConnection(connectionString);

                    //open the connection
                    connection.Open();

                    //get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    //set role lims user
                    if (limsUserPassword == "")
                    {
                        //lims_user is not password protected 
                        rolecommand = "set role lims_user";
                    }
                    else
                    {
                        //lims_user is password protected
                        rolecommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    //set the oracle user for this connection
                    OracleCommand command = new OracleCommand(rolecommand, connection);

                    //try/catch block
                    try
                    {
                        //execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        //throw the exeption
                        MyLog("Inconsistent role Security : " + f.Message);
                    }

                    //get session id
                    double sessionId = ntlsCon.GetSessionId();

                    //connect to the same session 
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    //Build the command 
                    command = new OracleCommand(sSql, connection);

                    //execute the command
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    //throw the exeption
                    MyLog("Err At GetConnection: " + e.Message);
                }
            }
            return connection;
        }

        public void MyLog(string sss)
        {
            try
            {
                using (FileStream file = new FileStream("C:\\logs\\sdglog.txt", FileMode.Append, FileAccess.Write))
                {
                    var streamWriter = new StreamWriter(file);
                    streamWriter.WriteLine(DateTime.Now);

                    streamWriter.WriteLine(sss);
                    streamWriter.WriteLine();
                    streamWriter.Close();
                }
            }
            catch
            {
            }


        }
    }
}