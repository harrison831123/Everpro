using Microsoft.AspNet.SignalR;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace SignalRAPI
{
    public class SignalRHub : Hub
    {
        // ... 其他 Hub 方法

        // /// <summary>
        // /// 发送给指定组
        // /// </summary>
        // public void CallGroup(string fromname, string content)
        // {
        //     string groupname = Context.QueryString["groupname"]; // 获取客户端发送过来的用户名
        //     // 根据username获取对应的ConnectionId
        //     Clients.Group(groupname).show(fromname + ":" + content);
        // }

        /// <summary>
        /// 当客户端成功连接到 Hub 时触发
        /// </summary>
        public override Task OnConnected()
        {
            // 獲取已驗證的用戶名
            //string username = Context.User.Identity.Name;

            //if (string.IsNullOrEmpty(username))
            //{

            //    username = "TESTER_A001";
            //}

            //if (!string.IsNullOrEmpty(username))
            //{
            //    // 現在，無論是否登入，它都會使用 "TESTER_A001" 來加入群組
            //    Groups.Add(Context.ConnectionId, username);
            //    // 您可以在這裡加入 Console.WriteLine($"Connection ID: {Context.ConnectionId} joined Group: {username}");
            //}
            var username = Context.QueryString["userid"];

            if (!string.IsNullOrEmpty(username))
            {
                Groups.Add(Context.ConnectionId, username);
            }

            return base.OnConnected();
        }

        /// <summary>
        /// 当客户端断开连接时触发
        /// </summary>
        public override Task OnDisconnected(bool stopCalled)
        {
            // 1. 获取已验证的用户名
            string username = Context.User.Identity.Name;

            if (!string.IsNullOrEmpty(username))
            {
                // 2. 将此连接从用户群组中移除
                Groups.Remove(Context.ConnectionId, username);

                // 建议：在此处移除 Cache 中的 ConnectionId 记录
                // 例如：YourCacheService.Remove(username);
            }

            return base.OnDisconnected(stopCalled);
        }
    }
}
