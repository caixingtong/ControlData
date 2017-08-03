using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace ControlData
{
    public class DbHelper
    {
        private readonly string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["Conn"].ConnectionString;
        private SqlConnection Connection = null; //数据库连接对象
        private SqlTransaction Trans = null;//事务对象
        private bool isUseTranMode = false;

        /// <summary>
        /// 构造方法
        /// </summary>
        public DbHelper():this(false)
        {
        }

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="isUseTranMode">是否启用事务模式</param>
        public DbHelper(bool isUseTranMode)
        {
            this.isUseTranMode = isUseTranMode;
            Connection = new SqlConnection(connectionString);
            Connection.Open();
            if (this.isUseTranMode) Trans = Connection.BeginTransaction();
        }

        /// <summary>
        /// 新建一个sqlCommand对象，并把事务对象添加进来
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="cmdType">是sql语句，还是存储过程</param>
        /// <returns>返回sqlCommand对象</returns>
        private SqlCommand BuildCommand(string cmdText, CommandType cmdType)
        {
            SqlCommand command;
            if(this.isUseTranMode)
            {
                command = new SqlCommand(cmdText, Connection) { CommandType = cmdType, CommandTimeout = 240, Transaction = Trans };
            }
            else
            {
                command = new SqlCommand(cmdText, Connection) { CommandType = cmdType, CommandTimeout = 240 };
            }
            return command;
        }

        /// <summary>
        /// 将sqlCommand对象的参数添加进来
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="cmdType">是sql语句，还是存储过程</param>
        /// <param name="parameters">传递进来的参数</param>
        /// <returns>返回sqlCommand对象</returns>
        private SqlCommand AddParamets(CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand command = BuildCommand(cmdText, cmdType);
            if (parameters == null) return command;
            foreach (SqlParameter parameter in parameters)
            {
                if ((parameter.Direction == ParameterDirection.InputOutput) && (parameter.Value == null))
                    parameter.Value = DBNull.Value;
                command.Parameters.Add(parameter);
            }
            return command;
        }

        /// <summary>
        /// 执行sql语句
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <returns></returns>
        public int ExecuteNonQuery(string cmdText)
        {
            return ExecuteNonQuery(CommandType.Text, cmdText, null);
        }
        
        /// <summary>
        ///  执行sql语句
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="parameters">传递进来的参数</param>
        /// <returns>返回影响的行数</returns>
        public int ExecuteNonQuery(string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = AddParamets(CommandType.Text, cmdText, parameters);
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        ///  执行sql语句
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="cmdType">是sql语句，还是存储过程</param>
        /// <param name="parameters">传递进来的参数</param>
        /// <returns>返回影响的行数</returns>
        public int ExecuteNonQuery(CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = AddParamets(cmdType, cmdText, parameters);
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 返回第一个对象
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <returns>返回第一个对象</returns>
        public object ExecuteScalar(string cmdText)
        {
            return ExecuteScalar(CommandType.Text, cmdText, null);
        }
        
        /// <summary>
        /// 返回第一个对象
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="parameters">传递进来的参数</param>
        /// <returns>返回第一个对象</returns>
        public object ExecuteScalar(string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = AddParamets(CommandType.Text, cmdText, parameters);
            return cmd.ExecuteScalar();
        }

        /// <summary>
        /// 返回第一个对象
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="cmdType">是sql语句，还是存储过程</param>
        /// <param name="parameters">传递进来的参数</param>
        /// <returns>返回第一个对象</returns>
        public object ExecuteScalar(CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = AddParamets(cmdType, cmdText, parameters);
            return cmd.ExecuteScalar();
        }

        /// <summary>
        /// 返回DataSet集
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <returns>返回DataSet集</returns>
        public DataSet ExecuteDataset(string cmdText)
        {
            return ExecuteDataset(CommandType.Text, cmdText, null);
        }

        /*该方法对原有的方法进行了扩展 2015-09-06 By JohnsonZhang*/
        /// <summary>
        /// 返回DataSet集
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="parameters">传递进来的参数</param>
        /// <returns>返回DataSet集</returns>
        public DataSet ExecuteDataset(string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = AddParamets(CommandType.Text, cmdText, parameters);
            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds);
            return ds;
        }

        /// <summary>
        /// 返回DataSet集
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="cmdType">是sql语句，还是存储过程</param>
        /// <param name="parameters">传递进来的参数</param>
        /// <returns>返回DataSet集</returns>
        public DataSet ExecuteDataset(CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = AddParamets(cmdType, cmdText, parameters);
            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds);
            return ds;
        }

        /// <summary>
        /// 返回DataTable
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <returns>返回DataTable</returns>
        public DataTable ExecuteDataTable(string cmdText)
        {
            return ExecuteDataTable(CommandType.Text, cmdText, null);
        }

        /*该方法对原有的方法进行了扩展 2015-09-06 By JohnsonZhang*/
        /// <summary>
        /// 返回DataTable
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="parameters">传递进来的参数</param>
        /// <returns>返回DataTable</returns>
        public DataTable ExecuteDataTable(string cmdText, params SqlParameter[] parameters)
        {
            DataSet ds = ExecuteDataset(CommandType.Text, cmdText, parameters);
            return ds == null || ds.Tables.Count == 0 ? null : ds.Tables[0];
        }

        /// <summary>
        /// 返回DataTable
        /// </summary>
        /// <param name="cmdText">要执行的sql语句</param>
        /// <param name="cmdType">是sql语句，还是存储过程</param>
        /// <param name="parameters">传递进来的参数</param>
        /// <returns>返回DataTable</returns>
        public DataTable ExecuteDataTable(CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            DataSet ds = ExecuteDataset(cmdType, cmdText, parameters);
            return ds == null || ds.Tables.Count == 0 ? null : ds.Tables[0];
        }

        /// <summary>
        /// 提交事务
        /// </summary>
        public void Commit()
        {
            this.Trans.Commit();
            this.Close();
        }

        /// <summary>
        /// 回滚事务
        /// </summary>
        public void Rollback()
        {
            this.Trans.Rollback();
            this.Close();
        }

        /// <summary>
        /// 关闭对象
        /// </summary>
        public void Close()
        {
            if (this.Connection.State != ConnectionState.Closed)
            {
                this.Connection.Close();
            }
        }
    }
}
