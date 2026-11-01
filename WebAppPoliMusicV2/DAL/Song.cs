using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Web;

namespace WebAppPoliMusicV2.DAL
{
    public class Song
    {
        private string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        public int Save(string name, string path, int plays)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "INSERT INTO [TBL_SONG]([SONG_NAME], [SONG_PATH], [PLAYS]) " +
                                      "VALUES(@SONG_NAME, @SONG_PATH, @PLAYS)";
                    cmd.Parameters.AddWithValue("@SONG_NAME", name);
                    cmd.Parameters.AddWithValue("@SONG_PATH", path);
                    cmd.Parameters.AddWithValue("@PLAYS", plays);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        public DataTable Read()
        {
            DataTable dtSong = new DataTable();
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT [ID_SONG], [SONG_NAME], [SONG_PATH], [PLAYS] FROM [TBL_SONG] ORDER BY [ID_SONG]", con))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dtSong);
                    }
                }
            }
            return dtSong;
        }

        public int Update(int id, string name, string path, int plays)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "UPDATE [TBL_SONG] SET [SONG_NAME] = @SONG_NAME, [SONG_PATH] = @SONG_PATH, [PLAYS] = @PLAYS " +
                                      "WHERE [ID_SONG] = @ID_SONG";
                    cmd.Parameters.AddWithValue("@SONG_NAME", name);
                    cmd.Parameters.AddWithValue("@SONG_PATH", path);
                    cmd.Parameters.AddWithValue("@PLAYS", plays);
                    cmd.Parameters.AddWithValue("@ID_SONG", id);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        public int Delete(int id)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "DELETE [TBL_SONG] WHERE [ID_SONG] = @ID_SONG";
                    cmd.Parameters.AddWithValue("@ID_SONG", id);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
