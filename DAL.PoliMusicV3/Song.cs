using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.PoliMusicV3
{
    public class Song
    {
        private string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        public DataTable Read()
        {
            DataTable dtSong = new DataTable();
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                string sentence =
                    "SELECT S.[ID_SONG], S.[SONG_NAME], S.[SONG_PATH], S.[PLAYS], " +
                    "ISNULL(STUFF((SELECT ', ' + A.[ARTIST_NAME] + ' ' + ISNULL(A.[ARTIST_LASTNAME], '') " +
                    "FROM [TBL_PLAYER] P INNER JOIN [TBL_ARTIST] A ON A.[ID_ARTIST] = P.[ID_ARTIST] " +
                    "WHERE P.[ID_SONG] = S.[ID_SONG] FOR XML PATH('')), 1, 2, ''), '') AS [ARTISTS], " +
                    "ISNULL((SELECT TOP 1 AL.[TITLE_ALBUM] FROM [TBL_TRACKLIST] T " +
                    "INNER JOIN [TBL_ALBUM] AL ON AL.[ID_ALBUM] = T.[ID_ALBUM] " +
                    "WHERE T.[ID_SONG] = S.[ID_SONG] ORDER BY T.[ID_ALBUM]), '') AS [ALBUM], " +
                    "'~/PlaylistImages/playlist' + CAST(S.[ID_SONG] AS VARCHAR(10)) + '.png' AS [COVER] " +
                    "FROM [TBL_SONG] S ORDER BY S.[ID_SONG]";

                using (SqlCommand cmd = new SqlCommand(sentence, con))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dtSong);
                        return dtSong;
                    }
                }
            }
        }
    }
}
