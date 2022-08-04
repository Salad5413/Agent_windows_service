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
        /// 

        // connection string to your Service Bus namespace
        static string connectionString = "Endpoint=sb://agentservice.servicebus.windows.net/;SharedAccessKeyName=RootManageSharedAccessKey;SharedAccessKey=63qwPMHVfLnj82rHgCmFZaT9CsIwYhg1W8WGfZ3dym0=";

        // name of your Service Bus queue
        static string queueName = "sendAgentMessage";

        // the client that owns the connection and can be used to create senders and receivers
        static ServiceBusClient client;

        // the sender used to publish messages to the queue
        static ServiceBusSender sender;

        // the processor that reads and processes messages from the queue
        static ServiceBusProcessor processor;

        // number of messages to be sent to the queue
        private const int numOfMessages = 3;

        // handle received messages
        static async Task MessageHandler(ProcessMessageEventArgs args)
        {
            string body = args.Message.Body.ToString();
            Console.WriteLine($"Received: {body}");
            
            // complete the message. message is deleted from the queue. 
            await args.CompleteMessageAsync(args.Message);
        }

        // handle any errors when receiving messages
        static Task ErrorHandler(ProcessErrorEventArgs args)
        {
            Console.WriteLine(args.Exception.ToString());
            return Task.CompletedTask;
        }

        static async Task Main(string[] args)
        {
            if (Environment.UserInteractive) //将 Windows 服务作为控制台应用运行
            {
                agent service1 = new agent(args);
                //service1.TestStartupAndStop(args);
                ////////////////////////////////////////////////
                // The Service Bus client types are safe to cache and use as a singleton for the lifetime
                // of the application, which is best practice when messages are being published or read
                // regularly.
                //
                // Create the clients that we'll use for sending and processing messages.
                client = new ServiceBusClient(connectionString);
                sender = client.CreateSender(queueName);

                // create a processor that we can use to process the messages
                processor = client.CreateProcessor(queueName, new ServiceBusProcessorOptions());

                //////////////////////////////////////////////////////////////////////////////////// 
                /// receive message
                try
                {
                    // add handler to process messages
                    processor.ProcessMessageAsync += MessageHandler;

                    // add handler to process any errors
                    processor.ProcessErrorAsync += ErrorHandler;

                    // start processing 
                    await processor.StartProcessingAsync();

                    Console.WriteLine("Wait for a minute and then press any key to end the processing");
                    Console.ReadKey();

                    // stop processing 
                    Console.WriteLine("\nStopping the receiver...");
                    await processor.StopProcessingAsync();
                    Console.WriteLine("Stopped receiving messages");
                }
                finally
                {
                    // Calling DisposeAsync on client types is required to ensure that network
                    // resources and other unmanaged objects are properly cleaned up.
                    await processor.DisposeAsync();
                    await client.DisposeAsync();
                }
                /// receive message
                //////////////////////////////////////////////////////////////////////////////////// 

                //////////////////////////////////////////////////////////////////////////////////// 
                ///send message
                // create a batch 
                /*
                using ServiceBusMessageBatch messageBatch = await sender.CreateMessageBatchAsync();

                for (int i = 1; i <= numOfMessages; i++)
                {
                    // try adding a message to the batch
                    if (!messageBatch.TryAddMessage(new ServiceBusMessage($"Message {i}")))
                    {
                        // if it is too large for the batch
                        throw new Exception($"The message {i} is too large to fit in the batch.");
                    }
                }

                try
                {
                    // Use the producer client to send the batch of messages to the Service Bus queue
                    await sender.SendMessagesAsync(messageBatch);
                    Console.WriteLine($"A batch of {numOfMessages} messages has been published to the queue.");
                }
                finally
                {
                    // Calling DisposeAsync on client types is required to ensure that network
                    // resources and other unmanaged objects are properly cleaned up.
                    await sender.DisposeAsync();
                    await client.DisposeAsync();
                }

                Console.WriteLine("Press any key to end the application");
                Console.ReadKey();
                */
                //send message
                //////////////////////////////////////////////////////////////////////////////////// 
                ////////////////////////////////////////////////
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
