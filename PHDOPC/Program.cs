using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Opc;
using OpcCom;
using Opc.Da;
using Opc.Hda;
using OpcRcw;
using OpcRcw.Da;
using OpcCom.Da.Wrapper;
using System.Data;

namespace PHDOPC
{
    class Program
    {

        private Opc.URL url;
        private Opc.Da.Server server;
        private OpcCom.Factory fact = new OpcCom.Factory();
        private Opc.Da.SubscriptionState groupState;
        private Opc.Da.Subscription groupRead;
        private Opc.Da.Subscription groupWrite;

        private Opc.Hda.Server serverHda;        

        public void Work()
        {
            int i = 5;
            if (i == 1)
            {
                url = new Opc.URL("opcda://192.168.0.4/OPC.PHDServerDA.1");
                server = new Opc.Da.Server(fact, null);
                server.Connect(url, new Opc.ConnectData(new System.Net.NetworkCredential()));

                Opc.Da.Item[] itemCollection = new Opc.Da.Item[1];
                itemCollection[0] = new Opc.Da.Item { ItemName = "RTOS.TEST.PV", MaxAge = -1 };
                Opc.Da.ItemValueResult[] result = server.Read(itemCollection);

                Console.WriteLine(result[0].Value);
                server.Disconnect();
            }
            else if (i == 2)
            {
                url = new Opc.URL("opcda://192.168.0.4/OPC.PHDServerDA.1");
                server = new Opc.Da.Server(fact, null);
                server.Connect(url, new Opc.ConnectData(new System.Net.NetworkCredential()));

                groupState = new Opc.Da.SubscriptionState();
                groupState.Name = "Group";
                groupState.ServerHandle = null;
                groupState.ClientHandle = Guid.NewGuid().ToString();
                groupState.Active = true;
                groupState.UpdateRate = 1000;
                groupState.Deadband = 0;
                groupState.Locale = null;

                groupRead = (Opc.Da.Subscription)server.CreateSubscription(groupState);


                string[] itemName = { "RTOS.TEST.PV", "RTOR.TI1237.DACA.PV" };

                Opc.Da.Item[] items = new Opc.Da.Item[2];

                for (int j = 0; j < items.Length; j++)
                {
                    items[j] = new Opc.Da.Item();
                    items[j].ClientHandle = Guid.NewGuid().ToString();
                    items[j].ItemPath = null;
                    items[j].ItemName = itemName[j];
                }

                groupRead.AddItems(items);
                //groupRead.DataChanged += new Opc.Da.DataChangedEventHandler(group_DataChanged);

                Opc.Da.ItemValueResult[] results = groupRead.Read(groupRead.Items);
                foreach (ItemValueResult result in results)
                {
                    Console.WriteLine("{0},{1},{2},{3}", result.ItemName, result.Value, result.Quality, result.Timestamp);
                }
                //groupRead.DataChanged -= new Opc.Da.DataChangedEventHandler(group_DataChanged);

                groupRead.RemoveItems(groupRead.Items);
                server.CancelSubscription(groupRead);
                groupRead.Dispose();
                server.Disconnect();
            }
            else if (i == 3)
            {
                url = new Opc.URL("opcda://192.168.0.4/OPC.PHDServerDA.1");
                server = new Opc.Da.Server(fact, null);
                server.Connect(url, new Opc.ConnectData(new System.Net.NetworkCredential()));

                groupState = new Opc.Da.SubscriptionState();
                groupState.Name = "Group";
                groupState.ServerHandle = null;
                groupState.ClientHandle = Guid.NewGuid().ToString();
                groupState.Active = true;
                groupState.UpdateRate = 1000;
                groupState.Deadband = 0;
                groupState.Locale = null;

                groupWrite = (Opc.Da.Subscription)server.CreateSubscription(groupState);

                string[] itemName = { "RTOS.TEST.PV", "RTOS.TEST2.PV" };

                Opc.Da.Item[] items = new Opc.Da.Item[2];

                for (int j = 0; j < items.Length; j++)
                {
                    items[j] = new Opc.Da.Item();
                    items[j].ClientHandle = Guid.NewGuid().ToString();
                    items[j].ItemPath = null;
                    items[j].ItemName = itemName[j];
                }

                groupWrite.AddItems(items);

                //groupRead.DataChanged += new Opc.Da.DataChangedEventHandler(group_DataChanged);

                Opc.Da.ItemValue[] writeValues = new Opc.Da.ItemValue[groupWrite.Items.Length];
                for (int k = 0; k < groupWrite.Items.Length; k++)
                {
                    writeValues[k] = new Opc.Da.ItemValue((ItemIdentifier)groupWrite.Items[k]);
                }

                writeValues[0].Value = 5;
                writeValues[1].Value = 6;

                groupWrite.Write(writeValues);

                Console.WriteLine("Press any key to close...");
                Console.ReadLine();

                //groupRead.DataChanged -= new Opc.Da.DataChangedEventHandler(group_DataChanged);

                groupWrite.RemoveItems(groupWrite.Items);
                server.CancelSubscription(groupWrite);
                groupWrite.Dispose();
                server.Disconnect();
            }
            else if (i == 4)
            {
                url = new Opc.URL("opchda://192.168.0.4/OPC.PHDServerHDA.1");
                serverHda = new Opc.Hda.Server(fact, null);
                serverHda.Connect(url, new Opc.ConnectData(new System.Net.NetworkCredential()));

                try
                {
                }
                catch (Opc.ConnectFailedException opcConnEx)
                {
                    Console.WriteLine(string.Format("Could not connect to server {0}", "OPC.PHDServerHDA.1"));
                    Console.WriteLine(opcConnEx.ToString());
                }

                Console.WriteLine("Are we connected? " + serverHda.IsConnected);

                string[] itemName = { "RTOS.TEST.PV", "RTOS.TEST2.PV" };

                Opc.Hda.Trend groupHda = new Trend(serverHda);

                groupHda.Name = "HDA";
                groupHda.AggregateID = AggregateID.NOAGGREGATE;
                DateTime startTime = DateTime.Now.AddHours(-1);
                DateTime endTime = DateTime.Now;
                groupHda.StartTime = new Opc.Hda.Time(startTime);
                groupHda.EndTime = new Opc.Hda.Time(endTime);
                groupHda.MaxValues = 0;
                groupHda.IncludeBounds = false;

                serverHda.Trends.Add(groupHda);

                Opc.Hda.Item[] items = new Opc.Hda.Item[2];
                for (int m = 0; m < 2; m++)
                {
                    items[m] = new Opc.Hda.Item();
                    items[m].ItemName = itemName[m];
                    items[m].ItemPath = null;
                    items[m].ClientHandle = Guid.NewGuid().ToString();
                }

                IdentifiedResult[] identifiedResult = serverHda.CreateItems(items);

                if (identifiedResult != null)
                {
                    foreach (IdentifiedResult item in identifiedResult)
                    {
                        if (item.ResultID.Succeeded())
                        {
                            groupHda.Items.Add(new Opc.Hda.Item(item));
                        }
                    }
                }

                Opc.Hda.ItemValueCollection[] results = groupHda.ReadRaw();

                Opc.Hda.ItemValueCollection result1 = results[0];

                foreach (Opc.Hda.ItemValue result in result1)
                {
                    Console.WriteLine("{0},{1},{2}", result.Value, result.Quality, result.Timestamp);
                }

                serverHda.Disconnect();
            }
            else if (i == 5)
            {
                url = new Opc.URL("opchda://192.168.0.4/OPC.PHDServerHDA.1");
                serverHda = new Opc.Hda.Server(fact, null);
                serverHda.Connect(url, new Opc.ConnectData(new System.Net.NetworkCredential()));

                try
                {
                }
                catch (Opc.ConnectFailedException opcConnEx)
                {
                    Console.WriteLine(string.Format("Could not connect to server {0}", "OPC.PHDServerHDA.1"));
                    Console.WriteLine(opcConnEx.ToString());
                }

                Console.WriteLine("Are we connected? " + serverHda.IsConnected);

                string[] itemName = { "RTOR.TI1237.DACA.PV", "RTOS.TEST2.PV" };

                Opc.Hda.Trend groupHda = new Trend(serverHda);

                groupHda.Name = "HDA";
                groupHda.AggregateID = AggregateID.AVERAGE;

                DateTime startTime = DateTime.Now.AddHours(-1);
                DateTime endTime = DateTime.Now;                
                //TimeSpan span = endTime.Subtract(startTime);                
                //int calcInterval = ((int)span.TotalSeconds);
                //groupHda.ResampleInterval = (decimal)calcInterval;

                //DateTime startTime = new DateTime(2020, 6, 12, 9, 0, 0, 0);
                //DateTime endTime = new DateTime(2020, 6, 12, 10, 0, 0, 0);
                int calcInterval = 300;
                groupHda.ResampleInterval = (decimal)calcInterval;
                

                groupHda.StartTime = new Opc.Hda.Time(startTime);
                groupHda.EndTime = new Opc.Hda.Time(endTime);
                groupHda.MaxValues = 0;
                groupHda.IncludeBounds = false;                

                serverHda.Trends.Add(groupHda);

                Opc.Hda.Item[] items = new Opc.Hda.Item[2];
                for (int m = 0; m < 2; m++)
                {
                    items[m] = new Opc.Hda.Item();
                    items[m].ItemName = itemName[m];
                    items[m].ItemPath = null;
                    items[m].ClientHandle = Guid.NewGuid().ToString();
                }

                IdentifiedResult[] identifiedResult = serverHda.CreateItems(items);

                if (identifiedResult != null)
                {
                    foreach (IdentifiedResult item in identifiedResult)
                    {
                        if (item.ResultID.Succeeded())
                        {
                            groupHda.Items.Add(new Opc.Hda.Item(item));
                        }
                    }
                }

                Opc.Hda.ItemValueCollection[] results = groupHda.ReadProcessed();

                Opc.Hda.ItemValueCollection result1 = results[0];

                foreach (Opc.Hda.ItemValue result in result1)
                {
                    Console.WriteLine("{0},{1},{2}", result.Value, result.Quality, result.Timestamp);
                }

                serverHda.Disconnect();
            }
        }

        public void group_DataChanged(object subscriptionHandle, object requestHandle, Opc.Da.ItemValueResult[] values)
        {
            Console.WriteLine("**********DataChanged**********");
            foreach (ItemValueResult item in values)
            {
                Console.WriteLine("DataChanged");
                Console.WriteLine("Item:{0},Value:{1}", item.ItemName, item.Value);
                Console.WriteLine("Quality:{0},Timestamp:{1}", item.Quality, item.Timestamp);
                Console.WriteLine("Handle:{0}", requestHandle);
            }
        }

        static void Main(string[] args)
        {
            Program pm = new Program();
            pm.Work();
        }
    }
}
