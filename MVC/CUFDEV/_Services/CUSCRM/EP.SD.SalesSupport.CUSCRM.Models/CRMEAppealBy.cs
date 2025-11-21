using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
    /// <summary>
    /// 立案申訴通知對象
    /// </summary>
    [Table ("CRMEAppealBy")]
    public class CRMEAppealBy:IModel
    {
        /// <summary>
        /// 自動編號
        /// </summary>
        [Column("ID", IsIdentity = true)]
        [Display(Name = "自動編號")]
        public int ID { get; set; }

        /// <summary>
        /// 受理編號
        /// </summary>
        [Column("No")]
        [Display(Name = "受理編號")]
        public string No { get; set; }

        /// <summary>
        /// 申訴人
        /// </summary>
        [Column("AppealName")]
        [Display(Name = "申訴人")]
        public string AppealName { get; set; }

        /// <summary>
        /// 申訴人行動電話
        /// </summary>      
        [Column("AppealMobile")]
        [Display(Name = "申訴人行動電話")]
        [Required(ErrorMessage = "申訴人行動電話必輸入")]
        public string AppealMobile { get; set; }

        /// <summary>
        /// 申訴人簡訊內容
        /// </summary>
        [Column("AppealMobile_Content")]
        [Display(Name = "申訴人簡訊內容")]
        public string AppealMobile_Content { get; set; }

        /// <summary>
        /// 申訴人email
        /// </summary>
        [Column("AppealEmail")]
        [Display(Name = "申訴人email")]
        //[Required(ErrorMessage = "必須輸入Email")]
        //[DataType(DataType.EmailAddress, ErrorMessage = "請輸入正確的電子信箱")]
        public string AppealEmail { get; set; }

        /// <summary>
        /// 申訴人email內容
        /// </summary>
        [Column("AppealEmail_Content")]
        [Display(Name = "申訴人email內容")]
        public string AppealEmail_Content { get; set; }

        /// <summary>
        /// 受任人
        /// </summary>
        [Column("EntrustdName")]
        [Display(Name = "受任人姓名")]
        public string EntrustdName { get; set; }

        /// <summary>
        /// 受任人行動電話
        /// </summary>
        [Column("EntrustdMobile")]
        [Display(Name = "受任人行動電話")]
        public string EntrustdMobile { get; set; }

        /// <summary>
        /// 受任人簡訊內容
        /// </summary>
        [Column("EntrustdMobile_Content")]
        [Display(Name = "受任人簡訊內容")]
        public string EntrustdMobile_Content { get; set; }

        /// <summary>
        /// 受任人Email
        /// </summary>
        [Column("EntrustdEmail")]
        [Display(Name = "受任人Email")]
        [DataType(DataType.EmailAddress, ErrorMessage = "請輸入正確的電子信箱")]
        public string EntrustdEmail { get; set; }

        /// <summary>
        /// 受任人Email內容
        /// </summary>
        [Column("EntrustdEmail_Content")]
        [Display(Name = "受任人Email內容")]
        public string EntrustdEmail_Content { get; set; }

        /// <summary>
        /// 稱謂
        /// </summary>
        [Column("Title")]
        [Display(Name = "稱謂")]
        public string Title { get; set; }

        /// <summary>
        /// 經辦人員
        /// </summary>
        [Column("DoUser")]
        [Display(Name = "經辦人員")]
        public string DoUser { get; set; }

        /// <summary>
        /// 經辦人員姓氏
        /// </summary>
        [Column("DoUserFirstName")]
        [Display(Name = "經辦人員姓氏")]
        public string DoUserFirstName { get; set; }

        /// <summary>
        /// 經辦人員分機表
        /// </summary>
        [Column("DoUserTelExt")]
        [Display(Name = "經辦人員分機表")]
        public string DoUserTelExt { get; set; }

        /// <summary>
        /// 建立人員
        /// </summary>
        [Column("Creator")]
        [Display(Name = "建立人員")]
        public string Creator { get; set; }

        /// <summary>
        /// 建立時間
        /// </summary>
        [Column("CreateTime")]
        [Display(Name = "建立時間")]
        public DateTime CreateTime { get; set; }


    }
}
