using System;
using EP.VBEPModels;
using Microsoft.CUF;

/// <summary>
/// 202505 by Fion 20250527001_佣酬預估試算
/// </summary>
namespace EP.SD.Collections.PlanSet.Service
{
    public class AgentRewardPolicyInfo
    {
        public AgentRewardPolicyInfo()
        {
            if (AgentRewardPolicy == null)
                AgentRewardPolicy = new AgentRewardPolicy();
        }

        public AgentRewardPolicy AgentRewardPolicy { get; set; }
        public ExtraData ExtraDataInfo { get; set; }
    }
}
