using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 受理保單資料
	/// </summary>
	[Table("CRMEInsurancePolicy")]
	public class CRMEInsurancePolicy : IModel
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
		/// 保單號碼
		/// </summary>
		[Column("PolicyNo")]
		[Display(Name = "保單號碼")]
		public string PolicyNo { get; set; }

		/// <summary>
		/// 保公保單號碼
		/// </summary>
		[Column("CoPolicyNo")]
		[Display(Name = "保公保單號碼")]
		public string CoPolicyNo { get; set; }

		/// <summary>
		/// 保險公司代碼
		/// </summary>
		[Column("CompanyCode")]
		[Display(Name = "保險公司代碼")]
		public string CompanyCode { get; set; }

		/// <summary>
		/// 保險公司名稱
		/// </summary>
		[Column("CompanyName")]
		[Display(Name = "保險公司名稱")]
		public string CompanyName { get; set; }

		/// <summary>
		/// 原幣別保費
		/// </summary>
		[Column("ModPrem")]
		[Display(Name = "原幣別保費")]
		public decimal ModPrem { get; set; }

		/// <summary>
		/// 保單狀態
		/// </summary>
		[Column("Status")]
		[Display(Name = "保單狀態")]
		public string Status { get; set; }

		/// <summary>
		/// 保單狀態名稱
		/// </summary>
		[Column("StatusName")]
		[Display(Name = "保單狀態名稱")]
		public string StatusName { get; set; }

		/// <summary>
		/// 繳別
		/// </summary>
		[Column("Modx")]
		[Display(Name = "繳別")]
		public string Modx { get; set; }

		/// <summary>
		/// 繳別名稱
		/// </summary>
		[Column("ModxName")]
		[Display(Name = "繳別名稱")]
		public string ModxName { get; set; }

		/// <summary>
		/// 要保人ID
		/// </summary>
		[Column("OwnerID")]
		[Display(Name = "要保人ID")]
		public string OwnerID { get; set; }

		/// <summary>
		/// 要保人
		/// </summary>
		[Column("Owner")]
		[Display(Name = "要保人")]
		public string Owner { get; set; }

		/// <summary>
		/// 要保人手機
		/// </summary>
		[Column("OwnerMobile")]
		[Display(Name = "要保人手機")]
		public string OwnerMobile { get; set; }

		/// <summary>
		/// 被保人ID
		/// </summary>
		[Column("InsuredID")]
		[Display(Name = "被保人ID")]
		public string InsuredID { get; set; }

		/// <summary>
		/// 被保人
		/// </summary>
		[Column("Insured")]
		[Display(Name = "被保人")]
		public string Insured { get; set; }

		/// <summary>
		/// 生效日
		/// </summary>
		[Column("IssDate")]
		[Display(Name = "生效日")]
		public string IssDate { get; set; }

		/// <summary>
		/// 原招攬人團隊
		/// </summary>
		[Column("VMCode")]
		[Display(Name = "原招攬人團隊")]
		public string VMCode { get; set; }

		/// <summary>
		/// 原招攬人團隊名稱
		/// </summary>
		[Column("VMName")]
		[Display(Name = "原招攬人團隊名稱")]
		public string VMName { get; set; }

		/// <summary>
		/// 原招攬人體系
		/// </summary>
		[Column("SMCode")]
		[Display(Name = "原招攬人體系")]
		public string SMCode { get; set; }

		/// <summary>
		/// 原招攬人體系名稱
		/// </summary>
		[Column("SMName")]
		[Display(Name = "原招攬人體系名稱")]
		public string SMName { get; set; }

		/// <summary>
		/// 原招攬人處
		/// </summary>
		[Column("CenterCode")]
		[Display(Name = "原招攬人處")]
		public string CenterCode { get; set; }

		/// <summary>
		/// 原招攬人處名稱
		/// </summary>
		[Column("CenterName")]
		[Display(Name = "原招攬人處名稱")]
		public string CenterName { get; set; }

		/// <summary>
		/// 原招攬人實駐
		/// </summary>
		[Column("WCCode")]
		[Display(Name = "原招攬人實駐")]
		public string WCCode { get; set; }

		/// <summary>
		/// 原招攬人實駐名稱
		/// </summary>
		[Column("WCName")]
		[Display(Name = "原招攬人實駐名稱")]
		public string WCName { get; set; }

		/// <summary>
		/// 原招攬人ID
		/// </summary>
		[Column("AgentCode")]
		[Display(Name = "原招攬人ID")]
		public string AgentCode { get; set; }

		/// <summary>
		/// 原招攬人
		/// </summary>
		[Column("AgentName")]
		[Display(Name = "原招攬人")]
		public string AgentName { get; set; }

		/// <summary>
		/// 原招攬人副總ID
		/// </summary>
		[Column("VMLeaderID")]
		[Display(Name = "原招攬人副總ID")]
		public string VMLeaderID { get; set; }

		/// <summary>
		/// 原招攬人副總
		/// </summary>
		[Column("VMLeader")]
		[Display(Name = "原招攬人副總")]
		public string VMLeader { get; set; }

		/// <summary>
		/// 原招攬人協理ID
		/// </summary>
		[Column("SMLeaderID")]
		[Display(Name = "原招攬人協理ID")]
		public string SMLeaderID { get; set; }

		/// <summary>
		/// 原招攬人協理
		/// </summary>
		[Column("SMLeader")]
		[Display(Name = "原招攬人協理")]
		public string SMLeader { get; set; }

		/// <summary>
		/// 原招攬人處主管ID
		/// </summary>
		[Column("CenterLeaderID")]
		[Display(Name = "原招攬人處主管ID")]
		public string CenterLeaderID { get; set; }

		/// <summary>
		/// 原招攬人處主管
		/// </summary>
		[Column("CenterLeader")]
		[Display(Name = "原招攬人處主管")]
		public string CenterLeader { get; set; }

		/// <summary>
		/// 接續服務人團隊
		/// </summary>
		[Column("SUVMCode")]
		[Display(Name = "接續服務人團隊")]
		public string SUVMCode { get; set; }

		/// <summary>
		/// 接續服務人團隊名稱
		/// </summary>
		[Column("SUVMName")]
		[Display(Name = "接續服務人團隊名稱")]
		public string SUVMName { get; set; }

		/// <summary>
		/// 接續服務人體系
		/// </summary>
		[Column("SUSMCode")]
		[Display(Name = "接續服務人體系")]
		public string SUSMCode { get; set; }

		/// <summary>
		/// 接續服務人體系名稱
		/// </summary>
		[Column("SUSMName")]
		[Display(Name = "接續服務人體系名稱")]
		public string SUSMName { get; set; }

		/// <summary>
		/// 接續服務人處
		/// </summary>
		[Column("SUCenterCode")]
		[Display(Name = "接續服務人處")]
		public string SUCenterCode { get; set; }

		/// <summary>
		/// 接續服務人處名稱
		/// </summary>
		[Column("SUCenterName")]
		[Display(Name = "接續服務人處名稱")]
		public string SUCenterName { get; set; }

		/// <summary>
		/// 接續服務人實駐
		/// </summary>
		[Column("SUWCCode")]
		[Display(Name = "接續服務人實駐")]
		public string SUWCCode { get; set; }

		/// <summary>
		/// 接續服務人實駐名稱
		/// </summary>
		[Column("SUWCName")]
		[Display(Name = "接續服務人實駐名稱")]
		public string SUWCName { get; set; }

		/// <summary>
		/// 接續服務人ID
		/// </summary>
		[Column("SUAgentCode")]
		[Display(Name = "接續服務人ID")]
		public string SUAgentCode { get; set; }

		/// <summary>
		/// 接續服務人
		/// </summary>
		[Column("SUAgentName")]
		[Display(Name = "接續服務人")]
		public string SUAgentName { get; set; }

		/// <summary>
		/// 接續服務人副總ID
		/// </summary>
		[Column("SUVMLeaderID")]
		[Display(Name = "接續服務人副總ID")]
		public string SUVMLeaderID { get; set; }

		/// <summary>
		/// 接續服務人副總
		/// </summary>
		[Column("SUVMLeader")]
		[Display(Name = "接續服務人副總")]
		public string SUVMLeader { get; set; }

		/// <summary>
		/// 接續服務人協理ID
		/// </summary>
		[Column("SUSMLeaderID")]
		[Display(Name = "接續服務人協理ID")]
		public string SUSMLeaderID { get; set; }

		/// <summary>
		/// 接續服務人協理
		/// </summary>
		[Column("SUSMLeader")]
		[Display(Name = "接續服務人協理")]
		public string SUSMLeader { get; set; }

		/// <summary>
		/// 接續服務人處主管ID
		/// </summary>
		[Column("SUCenterLeaderID")]
		[Display(Name = "接續服務人處主管ID")]
		public string SUCenterLeaderID { get; set; }

		/// <summary>
		/// 接續服務人處主管
		/// </summary>
		[Column("SUCenterLeader")]
		[Display(Name = "接續服務人處主管")]
		public string SUCenterLeader { get; set; }

		/// <summary>
		/// 有效件定義
		/// </summary>
		[Column("StatusType")]
		[Display(Name = "有效件定義")]
		public string StatusType { get; set; }

		/// <summary>
		/// 幣別
		/// </summary>
		[Column("MoneyType")]
		[Display(Name = "幣別")]
		public string MoneyType { get; set; }


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

	/// <summary>
	/// 受理保單資料的比對規則
	/// </summary>
	public class CRMEInsurancePolicyComparer : IEqualityComparer<CRMEInsurancePolicy>
	{
		public bool Equals(CRMEInsurancePolicy x, CRMEInsurancePolicy y)
		{
			if (ReferenceEquals(x, y))
				return true;

			if (ReferenceEquals(x, null) || ReferenceEquals(y, null))
				return false;

			return x.ID == y.ID && x.No == y.No && x.CompanyCode == y.CompanyCode && x.PolicyNo == y.PolicyNo && x.AgentCode == y.AgentCode && x.SUAgentCode == y.SUAgentCode;
		}

		public int GetHashCode(CRMEInsurancePolicy obj)
		{
			int hash = 17;
			hash = hash * 23 + obj.ID.GetHashCode();
			hash = hash * 23 + (obj.No != null ? obj.No.GetHashCode() : 0);
			hash = hash * 23 + (obj.CompanyCode != null ? obj.CompanyCode.GetHashCode() : 0);
			hash = hash * 23 + (obj.PolicyNo != null ? obj.PolicyNo.GetHashCode() : 0);
			hash = hash * 23 + (obj.AgentCode != null ? obj.AgentCode.GetHashCode() : 0);
			hash = hash * 23 + (obj.SUAgentCode != null ? obj.SUAgentCode.GetHashCode() : 0);
			return hash;
		}
	}
}
