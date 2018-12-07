using MarketingDataProcessing.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketingDataProcessing.Models
{

    /**
    * @author  LinhVT
    * @version 1.0
    * @since   2018/11/6
    */
    [SqlParameter("position")]
    class Position
    {
        private int _Position_Id;
        private string _Name;
        public Position()
        {

        }
        public Position(int id, string name)
        {
            Position_Id = id;
            Name = name;
        }
        #region properties
        [Required, AutoIncrement, PrimaryKey, SqlParameter("position_id")]
        public int Position_Id
        {
            get
            {
                return _Position_Id;
            }
            set
            {
                _Position_Id = value;
            }
        }
        [Required, Unique, ExcelColumn("G"), SqlParameter("name")]
        public string Name
        {
            get
            {
                return _Name;
            }
            set
            {
                _Name = value;
            }
        }
        #endregion
    }
}
