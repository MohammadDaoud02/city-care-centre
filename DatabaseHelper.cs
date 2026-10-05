using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace CityCareClinic
{
    public static class DatabaseHelper
    {
        // نص الاتصال الصحيح والمدعوم للـ LocalDB
        private static string connectionString = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=CityCareDB;Integrated Security=True;";

        // دالة تنفيذ أوامر الإضافة والتعديل والحذف
        public static bool ExecuteNonQuery(string query, SqlParameter[] parameters = null)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        con.Open();
                        cmd.ExecuteNonQuery();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطأ في الاتصال بقاعدة البيانات: " + ex.Message, "خطأ SQL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // دالة استرجاع البيانات للـ DataGridView
        public static DataTable ExecuteQuery(string query, SqlParameter[] parameters = null)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطأ أثناء جلب البيانات: " + ex.Message, "خطأ SQL", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return dt;
        }
    }
}