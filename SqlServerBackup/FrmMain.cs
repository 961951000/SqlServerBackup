using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SqlServerBackup
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dialog = new FolderBrowserDialog
            {
                Description = @"请选择备份文件保存路径"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var filename = Path.Combine(dialog.SelectedPath, $"{Regex.Match(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString, @"Catalog=([^;]+)").Groups[1].Value}_{DateTime.Now:yyyyMMdd}.bak");
                BakReductSql(filename, true);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var fileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = @"请选择文件",
                Filter = @"所有文件(*.*)|*.*"
            };
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                var filename = fileDialog.FileName;
                BakReductSql(filename, false);
            }
        }
        /// <summary>
        /// 对数据库的备份和恢复操作，Sql语句实现
        /// </summary>
        /// <param name="filename">保护路径的文件名</param>
        /// <param name="isBak">该操作是否为备份操作，是为true否，为false</param>
        private void BakReductSql(string filename, bool isBak)
        {
            var connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            var conn = new SqlConnection(connectionString);
            var cmdBakRst = new SqlCommand
            {
                Connection = conn,
                CommandType = CommandType.Text
            };
            var database = string.IsNullOrEmpty(conn.Database) ? Regex.Match(connectionString, @"Catalog=([^;]+)").Groups[1].Value : conn.Database;
            var cmdText = $"backup database {database} to disk='{filename}'";
            try
            {
                conn.Open();

                if (!isBak) //如果是恢复操作
                {
                    const string setOffline = "Alter database GroupMessage Set Offline With rollback immediate ";
                    const string setOnline = " Alter database GroupMessage Set Online With Rollback immediate";
                    cmdBakRst.CommandText = setOffline + cmdText + setOnline;
                }
                else
                {
                    cmdBakRst.CommandText = cmdText;
                }
                cmdBakRst.ExecuteNonQuery();
                MessageBox.Show(!isBak ? @"恭喜你，数据成功恢复为所选文档的状态！" : @"恭喜，你已经成功备份当前数据！", @"系统消息");
            }
            catch (SqlException sexc)
            {
                MessageBox.Show(@"失败，可能是对数据库操作失败，原因：" + sexc, @"数据库错误消息");
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"对不起，操作失败，可能原因：" + ex, @"系统消息");
            }
            finally
            {
                cmdBakRst.Dispose();
                conn.Close();
                conn.Dispose();
            }
        }
    }
}
