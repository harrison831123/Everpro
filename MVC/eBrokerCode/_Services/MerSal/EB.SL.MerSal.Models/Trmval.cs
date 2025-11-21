using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.MerSal.Models
{
    public class Trmval : IModel
    {
        /// <summary>
        /// 類別名稱
        /// </summary>
        [DataMember]
        [Column(Name = "term_id", IsKey = true, IsTrimRight = true)]
        public string TermID { get; set; }

        /// <summary>
        /// 類別代碼
        /// </summary>
        [DataMember]
        [Column(Name = "term_code", IsKey = true, IsTrimRight = true)]
        public string TermCode { get; set; }

        /// <summary>
        /// 類別代碼
        /// </summary>
        [DataMember]
        [Column(Name = "term_show", IsKey = true)]
        public string TermShow { get; set; }

        /// <summary>
        /// 類別代碼名稱
        /// </summary>
        [DataMember]
        [Column(Name = "term_meaning", IsTrimRight = true)]
        public string TermMeaning { get; set; }

        /// <summary>
        /// 類別代碼顯示順序
        /// </summary>
        [DataMember]
        [Column(Name = "term_sequence")]
        public short TermSequence { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        /// [DataMember]
        [Column("remark")]
        public string Remark { get; set; }
    }
}
