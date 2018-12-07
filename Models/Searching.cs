using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MarketingDataProcessing.Attributes;

namespace MarketingDataProcessing.Models
{
    /*
        DROP TABLE IF EXISTS
           searching
        ;
        DROP TABLE IF EXISTS
           synthesis
        ;

        -- 2. Create a "party" table and its properties.
        CREATE TABLE "synthesis"(
	        id SERIAL, -- No
	        json text
        );

        -- 2. Create a "party" table and its properties.
        CREATE TABLE "searching"(
	        id SERIAL, -- No
	        construction_no varchar(200), -- D
            construction_name varchar(500), --E
	        fulltext  text, --G
	        json text
        );
     */
    [SqlParameter("searching")]
    class Searching:BaseMVVM
    {

        private int _Id = 0;
        private string _ConstructionNo;
        private string _ConstructionName;
        private string _Fulltext;

        [Required,PrimaryKey,AutoIncrement]
        public int Id { get; set; }

        [Required, Unique, ExcelColumn("A"),SqlParameter("construction_no")]
        public string ConstructionNo {
            get
            {
                return _ConstructionNo;
            }

            set
            {
                _ConstructionNo = value;
                RaisePropertyChanged(nameof(ConstructionNo));
            }
        }


        [Required, ExcelColumn("B"), SqlParameter("construction_name")]
        public string ConstructionName {
            get
            {
                return _ConstructionName;
            }

            set
            {
                _ConstructionName = value;
                RaisePropertyChanged(nameof(ConstructionName));
            }
        }


        [Required, ExcelColumn("all"), SqlParameter("fulltext")]
        public string Fulltext {
            get
            {
                return _Fulltext;
            }

            set
            {
                _Fulltext = value;
                RaisePropertyChanged(nameof(Fulltext));
            }
        }
        public static Searching CreateNew(Utilities.DataSet data)
        {

            Searching temp = new Searching();
            try
            {
                Type type = typeof(Searching);
                temp.ConstructionNo = data.Value(SqlParameterAttribute.GetNameOfParameterInSql(type, nameof(ConstructionNo)));
                temp.ConstructionName = data.Value(SqlParameterAttribute.GetNameOfParameterInSql(type, nameof(ConstructionName)));
                temp.Fulltext = data.Value(SqlParameterAttribute.GetNameOfParameterInSql(type, nameof(Fulltext)));
            }
            catch (Exception)
            {

                return null;
            }
            return temp;
        }

    }
}
