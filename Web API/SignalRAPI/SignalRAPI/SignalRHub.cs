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
        /// <summary>
        /// 當客戶端成功連線到 Hub 時觸發
        /// </summary>
        public override Task OnConnected()
        {
            // 從 Query String 獲取用戶 ID
            var username = Context.QueryString["userid"];

            if (!string.IsNullOrEmpty(username))
            {
                // 將此 Connection ID 加入到以 username 為名的群組中
                Groups.Add(Context.ConnectionId, username);
            }

            return base.OnConnected();
        }

        /// <summary>
        /// 當客戶端斷開連線時觸發
        /// </summary>
        public override Task OnDisconnected(bool stopCalled)
        {
            // 1. 獲取已驗證的用戶名稱 
            //string username = Context.User.Identity.Name;
            var username = Context.QueryString["userid"];

            if (!string.IsNullOrEmpty(username))
            {
                // 2. 將此連線從用戶群組中移除
                Groups.Remove(Context.ConnectionId, username);
            }

            return base.OnDisconnected(stopCalled);
        }
    }
}
