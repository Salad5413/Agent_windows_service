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
using Newtonsoft.Json.Linq;

namespace WindowsService1
{
    public partial class agent : ServiceBase
    {
        public agent(string[] args)
        {
            InitializeComponent();
        }

        // 遍历并列出本机所有进程
        private void list_process()
        {
            // 存储所有进程到 local_processes
            Process[] local_processes = Process.GetProcesses();

            // 在console中列出所有进程
            foreach (Process p in local_processes)
            {
                Console.WriteLine("{0}\t\tId:{1}", p, p.Id);

            }
            Console.WriteLine("本机共有 {0} 个进程", local_processes.Length);
        }

        /* 根据 Id 结束进程
        * @param 需结束的进程Id
        * @return true 成功结束进程，false 未成功结束进程
        */
        private bool kill_process(string id)
        {
            try // 尝试结束进程
            {
                using (Process chosen = Process.GetProcessById(Int32.Parse(id)))
                {
                    chosen.Kill();
                    chosen.WaitForExit();
                    Console.WriteLine("{0}进程已结束", id);
                    return true;
                }
            }
            catch (Exception) // 若进程结束失败，弹出提示并结束方法
            {
                Console.WriteLine("{0}进程结束失败",id);
                return false;
            }
        }

        /* 根据名称搜索进程，列出在 console 上
        * @param 需搜索的进程名称
        * @return 搜索结果的 Array
        */
        private Process[] search_process(string name)
        {
            Process[] local_processes = Process.GetProcesses();
            Process[] result = new Process[local_processes.Length]; // 存储所有查询到的进程，并作为最后输出结果
            int i = 0; // 记录找到的结果个数

            Console.WriteLine("");

            foreach (Process p in local_processes) // 遍历所有进程名称，将任何名称中包含"name"字符串、
                                                   // 的进程存储到Array“result”
            {
                if(p.ProcessName.ToLower().Contains(name.ToLower()))
                {
                    result[i] = p;
                    i++;
                    Console.WriteLine("{0}\t\t\tId:{1}", p, p.Id);
                }
            }
            //输出结果
            Console.WriteLine("\n共找到 {0} 个结果", i);
            return result;
        }

        /* 结束一组进程
        * @param 需结束的进程 Array
        * @return true 成功结束所有进程，false 未成功结束所有进程
        */
        private bool kill_all(Process[] p_array)
        {
            // 二次确认是否结束所有进程
            Console.WriteLine("是否结束以下进程？");
            foreach (Process p in p_array) 
            {   
                if(p != null)
                {
                    Console.WriteLine("{0}\t\tId:{1}", p.ProcessName, p.Id);
                }
            }
            Console.WriteLine("[Y/n]");//输入“Y”确定，输入其他任何字符则取消
            string choice = Console.ReadLine();

            if(choice.Equals("Y")) // 确定并尝试结束所有上一次搜索结果中的进程
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
                    Console.WriteLine("进程已结束"); // 所有进程成功结束
                    return true;
                }
                catch (Exception)
                {
                    Console.WriteLine("结束进程失败"); // 所有（部分）进程未能结束
                    return false;
                }
            }
            else // 取消并结束方法
            { 
                return false; 
            }
        }


        protected override void OnStart(string[] args)
        {
            /* 
             * 第一周任务：被托管机器本地Agent服务
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

            int num_process = Process.GetProcesses().Length; // 获取进程总数，若没有任何进程，停止循环
            Process[] searched_processes = new Process[1]; // 该array存储“f”指令的搜索结果
            while (num_process > 0)
            {
                Console.WriteLine("\n***************************************\n" +
                                  "list : 列出本机所有进程及其Id\n" +
                                  "k + (进程Id)：结束该Id对应的进程\n" +
                                  "f + (进程名)：根据名称搜索进程\n" +
                                  "kill_all : 结束上一次搜索找到的所有进程" +
                                  "s + (程序名) : 打开该应用程序" +
                                  "\n***************************************\n");


                string command = Console.ReadLine(); // 读取用户输入
                string com = command.Substring(0, command.IndexOf(" ")); // 提取指令
                string info = command.Substring(command.IndexOf(" ") + 1, command.Length - command.IndexOf(" ") - 1); // 提取后续信息

                if (com.Equals("list")) // 列出本机所有进程及其Id
                {
                    list_process();
                }
                else if (com.Equals("k")) // 结束该Id(info)对应的进程
                {
                    kill_process(info);
                }
                else if (com.Equals("f")) // 根据名称(info)搜索进程
                {
                    searched_processes = search_process(info);
                }
                else if (com.Equals("kill_all")) // 结束上一次搜索找到的所有进程
                {
                    if (searched_processes[0] != null) // 判断上一次是否搜索到结果
                    {
                        kill_all(searched_processes); // 若有结果，结束所有进程
                    }
                    else
                    {
                        Console.WriteLine("搜索结果为空"); // 若结果为空，则不结束任何进程
                        continue;
                    }
                }
                else if (com.Equals("s")) // 根据应用程序名称打开该程序进程
                {
                    Process.Start(info);
                }
                else // 若无法识别指令，弹出提醒并继续循环
                {
                    Console.WriteLine("Unknown Command: {0}", com);
                }

                // 更新进程总数
                num_process = Process.GetProcesses().Length;
            }
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
