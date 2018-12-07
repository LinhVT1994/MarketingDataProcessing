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
           Execution
        ;

        -- 2. Create a "party" table and its properties.
        CREATE TABLE "Execution"(
	        id SERIAL, -- No
	        construction_no varchar(200), -- D
            construction_name varchar(500), --E
	        position  varchar(500), --G
	        owner varchar(500), -- H
	        partner varchar(500) -- O
        );
     */
     [SqlParameter("Execution")]
    class Execution
    {
 
        [Required,PrimaryKey,AutoIncrement]
        public int Id { get; set; }

        [Required, Unique, ExcelColumn("D"),SqlParameter("construction_no")]
        public string ConstructionNo { get; set; }


        [Required, ExcelColumn("E"), SqlParameter("construction_name")]
        public string ConstructionName { get; set; }


        [Required, ExcelColumn("G"), SqlParameter("position")]
        public string Position { get; set; }


        [Required, ExcelColumn("H"), SqlParameter("owner")]
        public string Owner { get; set; }


        [Required, ExcelColumn("O"), SqlParameter("partner")]
        public string Partner { get; set; }
    }
}
