using MarketingDataProcessing.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketingDataProcessing.Models
{
    [SqlParameter("json")]
    class Synthesis:BaseMVVM
    {
        private int _Id = 0;
        private string _Json;
        [Required, PrimaryKey, AutoIncrement]
        public int Id { get; set; }

        [Required, Unique, ExcelColumn("all"), SqlParameter("json")]
        public string Json
        {
            get
            {
                return _Json;
            }

            set
            {
                _Json = value;
                RaisePropertyChanged(nameof(Json));
            }
        }
        public static Synthesis CreateNew(Utilities.DataSet data)
        {

            Synthesis temp = new Synthesis();
            try
            {
                Type type = typeof(Searching);
                temp.Json = data.Value(SqlParameterAttribute.GetNameOfParameterInSql(type, nameof(Json)));
            }
            catch (Exception)
            {

                return null;
            }
            return temp;
        }
    }
}
