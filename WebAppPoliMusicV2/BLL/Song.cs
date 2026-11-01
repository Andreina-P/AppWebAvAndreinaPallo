using System;
using System.Collections.Generic;
using System.Data;
using System.Web;

namespace WebAppPoliMusicV2.BLL
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
            return new DAL.Song(strConnectionString).Save(name, path, plays);
        }

        public DataTable Read()
        {
            return new DAL.Song(strConnectionString).Read();
        }

        public int Update(int id, string name, string path, int plays)
        {
            return new DAL.Song(strConnectionString).Update(id, name, path, plays);
        }

        public int Delete(int id)
        {
            return new DAL.Song(strConnectionString).Delete(id);
        }
    }
}