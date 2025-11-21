namespace SACTAPI.Models
{
    /// <summary>
    /// 需求單號：
    /// 簽約資料
    /// </summary>
    public class IDownLoadAGData
    {
        public IDownLoadAGData()
        {
            AGID = string.Empty;
            Name = string.Empty;
            SignOffType = string.Empty;
            CenterName = string.Empty;
            CenterCode = string.Empty;
            AdminName = string.Empty;
            AdminID = string.Empty;
            WcCenterName = string.Empty;
            WcCenter = string.Empty;
            UmName = string.Empty;
            UmCode = string.Empty;
            LevelName = string.Empty;
            Sex = string.Empty;
            Birthday = string.Empty;
            ReNameYN = string.Empty;
            OldName = string.Empty;
            marriedYN = "N";
            wifeEPYN = string.Empty;
            wifeName = string.Empty;
            wifeCenterName = string.Empty;
            Education = string.Empty;
            Mobil = string.Empty;
            Tel = string.Empty;
            Email = string.Empty;
            Address1 = string.Empty;
            Address2 = string.Empty;
            BankCode = string.Empty;
            BranchCode = string.Empty;
            Account = string.Empty;
            DonateYN = string.Empty;
            DonateDate = string.Empty;
            IntroducerName = string.Empty;
            IntroducerID = string.Empty;
            DirectorName = string.Empty;
            DirectorID = string.Empty;
            LifeInsurance = string.Empty;
            LicenseNo1 = string.Empty;
            Fee1 = 0;
            ForeignCurrency = string.Empty;
            LicenseNo3 = string.Empty;
            Fee3 = 0;
            Investment = string.Empty;
            LicenseNo4 = string.Empty;
            Fee4 = 0;
            PropertyInsurance = string.Empty;
            LicenseNo2 = string.Empty;
            Fee2 = 0;
            FinancialMarkets = string.Empty;
            FMLicenseNo = string.Empty;
            OtherLicense = string.Empty;
            UniversalAccount = string.Empty;
            FeeTotal = 0;
        }

        public string AGID { get; set; }            //	簽約人身份證字號
        public string Name { get; set; }            //	姓名
        public string SignOffType { get; set; }     //	台灣身份證/居留證
        public string CenterName { get; set; }      //	處別
        public string CenterCode { get; set; }      //	處代號
        public string AdminName { get; set; }       //	業務協理姓名
        public string AdminID { get; set; }         //	業務協理ID
        public string WcCenterName { get; set; }    //	通訊處名稱
        public string WcCenter { get; set; }        //	通訊處代碼
        public string UmName { get; set; }          //	組織名稱
        public string UmCode { get; set; }          //	組織代號
        public string LevelName { get; set; }       //	推薦稱謂
        public string Sex { get; set; }             //	性別
        public string Birthday { get; set; }        //	出生年月日
        public string ReNameYN { get; set; }        //	曾更名
        public string OldName { get; set; }         //	舊名
        public string marriedYN { get; set; }       //	是否已婚
        public string wifeEPYN { get; set; }        //	配偶是否為永達保經業務員
        public string wifeName { get; set; }        //	配偶姓名
        public string wifeCenterName { get; set; }  //	配偶處別
        public string Education { get; set; }       //	畢業之最高學歷
        public string Mobil { get; set; }           //	行動電話
        public string Tel { get; set; }             //	住家電話
        public string Email { get; set; }           //	E-mail信箱
        public string Address1 { get; set; }        //	戶籍地址
        public string Address2 { get; set; }        //	目前住所地址
        public string BankCode { get; set; }        //	銀行代碼
        public string BranchCode { get; set; }      //	分行名稱
        public string Account { get; set; }         //	帳號
        public string DonateYN { get; set; }        //	同意基金會捐款
        public string DonateDate { get; set; }      //	扣款起始年月
        public string IntroducerName { get; set; }  //	推介人
        public string IntroducerID { get; set; }    //	推介人ID
        public string DirectorName { get; set; }    //	輔導人
        public string DirectorID { get; set; }      //	輔導人ID
        public string LifeInsurance { get; set; }   //	人身保險資格
        public string LicenseNo1 { get; set; }      //	人身合格證號
        public int Fee1 { get; set; }               //	人身登錄費
        public string ForeignCurrency { get; set; } //	外幣資格
        public string LicenseNo3 { get; set; }      //	外幣合格證號
        public int Fee3 { get; set; }               //	外幣登錄費
        public string Investment { get; set; }      //	投資型資格
        public string LicenseNo4 { get; set; }      //	投資型合格證號
        public int Fee4 { get; set; }               //	投資型登錄費
        public string PropertyInsurance { get; set; }   //	財產保險資格
        public string LicenseNo2 { get; set; }      //	財產合格證號
        public int Fee2 { get; set; }               //	財產登錄費
        public string FinancialMarkets { get; set; }    //	金融常識與職業道德資格
        public string FMLicenseNo { get; set; }     //	金融常識與職業道德合格證號
        public string OtherLicense { get; set; }    //	其他金融相關證照
        public string UniversalAccount { get; set; }    //	萬用帳號
        public int FeeTotal { get; set; }	        //	登錄費合計

    }
}