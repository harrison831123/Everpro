//需求單號：20250707001 競賽計C商品清單。 2025/07/03 BY Harrison 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;
using Microsoft.CUF.Framework;

namespace EP.SD.Collections.PlanSet.Service
{
    [DataContract]
    public class PlanSetWarptSetCondition
    {
        //[DataMember]
        //[Display(Name = "商品清單")]
        //public int ProductList { get; set; }

        ///// <summary>
        ///// 商品清單(列舉)
        ///// </summary>
        //public ProductList Type
        //{
        //    get
        //    {
        //        return (ProductList)ProductList;
        //    }
        //}

        [DataMember]
        [Display(Name = "商品清單")]
        public ProductList SelectedProduct { get; set; }

        [DataMember]
        [Display(Name = "保險公司")]
        public string CompanyCode { get; set; }

        //[DataMember]
        //[Display(Name = "年期")]
        //public string PlanYear { get; set; }

    }
}
