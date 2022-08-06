using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Security.Permissions;
using System.Security.Principal;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Azure.Messaging.ServiceBus;



namespace WindowsService1
{
    public partial class agent : ServiceBase
    {
        public agent(string[] args)
        {
            InitializeComponent();
        }

        ///////////////////////////////////////////////////////////
        /// Server Bus initialize 

        // connection string to your Service Bus namespace
        static string connectionString = "Endpoint=sb://agentservice.servicebus.windows.net/;SharedAccessKeyName=RootManageSharedAccessKey;SharedAccessKey=63qwPMHVfLnj82rHgCmFZaT9CsIwYhg1W8WGfZ3dym0=";

        // name of Service Bus queue
        static string queueSenderName = "sendAgentMessage";
        static string queueReceiverName = "receiveAgentMessage";

        // the client that owns the connection and can be used to create senders and receivers
        static ServiceBusClient client;

        // the sender used to publish messages to the queue
        static ServiceBusSender sender;

        // the sender used to receive messages from the queue
        static ServiceBusReceiver receiver;

        /*
        // the processor that reads and processes messages from the queue
        static ServiceBusProcessor processor;
        */

        protected override void OnStart(string[] args)
        {
            windowsService();
        }

        /* 
         * 被托管机器本地Agent服务
         * 
         * 主功能一览：
         * list: 列出本机所有进程及其Id
         * k +（进程Id）：结束该Id对应的进程
         * f + （进程名）：根据名称搜索进程
         * kill_all: 结束上一次搜索找到的所有进程
         * s + （进程名）：根据名称打开该应用程序
         * 
         * 需要手动启动/停止该Windows服务
         * 
         */
        public static async Task windowsService()
        {


            int num_process = Process.GetProcesses().Length; // 获取进程总数，若没有任何进程，停止循环
            Process[] searched_processes = new Process[1]; // 该array存储“f”指令的搜索结果

            /*
            await sendMessageToServerBus("\n***************************************\n" +
                                             "list: 列出本机所有进程及其Id\n" +
                                             "k + (进程Id)：结束该Id对应的进程\n" +
                                             "f + (进程名)：根据名称搜索进程\n" +
                                             "kill_all: 结束上一次搜索找到的所有进程" +
                                             "s + (程序名) : 打开该应用程序" +
                                             "\n***************************************\n");
            */

            // windows服务程序主体
            while (num_process > 0)
            {
                string command = await receiveMessageFromServerBus(); // 从 server bus queue 读取指令

                string com = "";
                string info = "";

                if (command.Contains(" "))
                {
                    com = command.Substring(0, command.IndexOf(" ")); // 提取指令
                    info = command.Substring(command.IndexOf(" ") + 1, command.Length - command.IndexOf(" ") - 1); // 提取后续信息
                }
                else
                {
                    com = command;
                }
                

                if (com.Equals("list")) // 列出本机所有进程及其Id
                {
                    await list_process();
                }
                else if (com.Equals("k")) // 结束该Id(info)对应的进程
                {
                    await kill_process(info);
                }
                else if (com.Equals("f")) // 根据名称(info)搜索进程
                {
                    searched_processes = await search_process(info);
                }
                else if (com.Equals("kill_all")) // 结束上一次搜索找到的所有进程
                {
                    if (searched_processes[0] != null) // 判断上一次是否搜索到结果
                    {
                        await kill_all(searched_processes); // 若有结果，结束所有进程
                    }
                    else
                    {
                        await sendMessageToServerBus("搜索结果为空，无法删除进程");
                        continue;
                    }
                }
                else if (com.Equals("s")) // 根据应用程序名称打开该程序进程
                {
                    Process.Start(info);
                }
                else // 若无法识别指令，弹出提醒并继续循环
                {
                    await sendMessageToServerBus($"Unknown Command: {com}");
                }

                // 更新进程总数
                num_process = Process.GetProcesses().Length;
            }
        }

        static async Task sendMessageToServerBus(string inputMessage)
        {
            // Create the clients that we'll use for sending and processing messages.
            client = new ServiceBusClient(connectionString);

            // Create a sender that we can use to send the message
            sender = client.CreateSender(queueReceiverName);


            JsonMessage jsonMessage = new JsonMessage { message = inputMessage };
            string jsonMessageBody = JsonConvert.SerializeObject(jsonMessage);


            // create a message that we can send. UTF-8 encoding is used when providing a string.
            ServiceBusMessage jsonMessageSend = new ServiceBusMessage(jsonMessageBody);

            try
            {
                // Use the producer client to send the batch of messages to the Service Bus queue
                await sender.SendMessageAsync(jsonMessageSend);
                Console.WriteLine($"messages \"{inputMessage}\" has been published to the queue.");
            }
            finally
            {
                // Calling DisposeAsync on client types is required to ensure that network
                // resources and other unmanaged objects are properly cleaned up.
                await sender.DisposeAsync();
                await client.DisposeAsync();
            }
        }

        static async Task<string> receiveMessageFromServerBus()
        {
            // Create the client object that will be used to create sender and receiver objects
            client = new ServiceBusClient(connectionString);

            // Create a receiver that we can use to receive the message
            receiver = client.CreateReceiver(queueSenderName);


            // Intialize ServiceBusReceivedMessage var to contain the received message
            ServiceBusReceivedMessage receivedMessage;
            

            while (true)
            {
                // Receive message
                receivedMessage = await receiver.ReceiveMessageAsync();

                if(receivedMessage != null)
                {
                    JsonMessage jsonMessage = JsonConvert.DeserializeObject<JsonMessage>(receivedMessage.Body.ToString());

                    // complete the message. message is deleted from the queue. 
                    await receiver.CompleteMessageAsync(receivedMessage);

                    // Calling DisposeAsync on client types is required to ensure that network
                    // resources and other unmanaged objects are properly cleaned up.
                    await receiver.DisposeAsync();
                    await client.DisposeAsync();

                    return jsonMessage.message;
                }

            }

        }

        // 遍历并列出本机所有进程
        static private async Task list_process()
        {
            // 存储所有进程到 local_processes
            Process[] local_processes = Process.GetProcesses();

            string allProcesses = "";
            // 将所有进程添加至一个string
            foreach (Process p in local_processes)
            {
                allProcesses += p.ProcessName;
                allProcesses += "\t\tId: ";
                allProcesses += p.Id.ToString();
                allProcesses += "\n";
            }
            allProcesses += $"本机共有 {local_processes.Length} 个进程";

            // send messages to server bus
            await sendMessageToServerBus(allProcesses);
        }


        /* 根据 Id 结束进程
        * @param 需结束的进程Id
        * @return true 成功结束进程，false 未成功结束进程
        */
        static private async Task kill_process(string id)
        {
            try // 尝试结束进程
            {
                using (Process chosen = Process.GetProcessById(Int32.Parse(id)))
                {
                    chosen.Kill();
                    chosen.WaitForExit();
                    await sendMessageToServerBus($"{id}进程已结束");
                }
            }
            catch (Exception) // 若进程结束失败，弹出提示并结束方法
            {
                await sendMessageToServerBus($"{id}进程结束失败");
            }
        }


        /* 根据名称搜索进程，列出在 console 上
        * @param 需搜索的进程名称
        * @return 搜索结果的 Array
        */
        static private async Task<Process[]> search_process(string name)
        {
            Process[] local_processes = Process.GetProcesses();
            Process[] result = new Process[local_processes.Length]; // 存储所有查询到的进程，并作为最后输出结果
            int i = 0; // 记录找到的结果个数

            string processesList = "";

            foreach (Process p in local_processes) // 遍历所有进程名称，将任何名称中包含"name"字符串、
                                                   // 的进程存储到Array“result”并将列表发送至server
            {
                if(p.ProcessName.ToLower().Contains(name.ToLower()))
                {
                    result[i] = p;
                    i++;
                    processesList += p.ProcessName;
                    processesList += "\t\tId: ";
                    processesList += p.Id.ToString();
                    processesList += "\n";
                }
            }

            // 输出结果
            processesList += $"\n共找到 {i} 个结果";

            // 将列表发送至server
            await sendMessageToServerBus(processesList);

            return result;
        }

        /* 结束一组进程
        * @param 需结束的进程 Array
        * @return true 成功结束所有进程，false 未成功结束所有进程
        */
        static private async Task kill_all(Process[] p_array)
        {
            // 二次确认是否结束所有进程

            string doubleCheckMessage = "是否结束以下进程？";

            foreach (Process p in p_array) 
            {   
                if(p != null)
                {
                    doubleCheckMessage += p.ProcessName;
                    doubleCheckMessage += "\t\tId: ";
                    doubleCheckMessage += p.Id.ToString();
                    doubleCheckMessage += "\n";
                }
            }
            doubleCheckMessage += "[Y/n]";  //输入“Y”确定，输入其他任何字符则取消

            // 发送二次确认消息至 server
            await sendMessageToServerBus(doubleCheckMessage);

            // 等待接收二次确认消息
            string choice = await receiveMessageFromServerBus();

            if (choice.Equals("Y")) // 确定并尝试结束所有上一次搜索结果中的进程
            {
                try
                {
                    foreach (Process p in p_array)
                    {
                        if(p != null)
                        {
                            p.Kill();
                            p.WaitForExit();
                        }
                    }
                    await sendMessageToServerBus("进程已结束"); // 所有进程成功结束
                }
                catch (Exception)
                {
                    await sendMessageToServerBus("结束进程失败"); // 所有（部分）进程未能结束
                }
            }
        }

        // string 转成 json
        static JObject toJson(string str)
        {
            JObject jsonResult = JObject.Parse(str);
            return jsonResult;
        }






        protected override void OnStop()
        {
        }

        private void process1_Exited(object sender, EventArgs e)
        {
            
        }

        //将 Windows 服务作为控制台应用运行的辅助方法
        internal void TestStartupAndStop(string[] args)
        {
            this.OnStart(args);
            Console.ReadLine();
            this.OnStop();
        }


    }
}
