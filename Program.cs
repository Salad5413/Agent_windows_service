using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using Azure.Messaging.ServiceBus;
using Newtonsoft.Json.Linq;

namespace WindowsService1
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>

        static void Main(string[] args)
        {
            if (Environment.UserInteractive) //将 Windows 服务作为控制台应用运行
            {
                agent service1 = new agent(args);
                service1.TestStartupAndStop(args);
            }
            else // 原Main方法
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                    new agent(args)
                };
                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}
