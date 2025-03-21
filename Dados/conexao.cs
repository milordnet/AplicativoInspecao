using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace Milord.Dados
{
    public class conexao
    {
     

        public static string caminhoconexao = ConfigurationManager.ConnectionStrings["milordpro"].ConnectionString;

        protected bool Executarprocedure(string nomeprocedure, List<SqlParameter> parametros)
        {
            SqlConnection conn = new SqlConnection(caminhoconexao);
            SqlCommand command = new SqlCommand(nomeprocedure, conn);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.AddRange(parametros.ToArray());
            conn.Open();

            try
            {
                command.ExecuteNonQuery();
                return true;
            }
            catch (Exception e)
            {
                throw new ArgumentException(e.Message);
            }
            finally
            {
                conn.Close();
            }

        }

        public DataTable retornadados(string sql)
        {
            SqlConnection conn = new SqlConnection(caminhoconexao);
            try
            {
                conn.Open();
                SqlDataAdapter adp = new SqlDataAdapter(sql, conn);
                DataTable dt = new DataTable();
                adp.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                throw new ArgumentException(e.Message);

            }
            finally
            {
                conn.Close();
            }

        }
    }

}