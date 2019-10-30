using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Neo4jClient;
using Microsoft.Office.Interop.Outlook;

namespace OutlookToNeo4j
{
    using Exception = Microsoft.Office.Interop.Outlook.Exception;


    class Program
    {
        
        static void Main(string[] args)
        {
           

            //New Bolt Graph Client
            var graphClient = new BoltGraphClient("bolt://localhost:7687/", "neo4j", "Myneo4j");
            graphClient.Connect();

            //Store the data from the DataProvider
            StoreData(graphClient, GetDataProvider());
            Console.WriteLine("Done!");
        }

        //This is the Cypher bit.
        //I've used 'nameof' to allow for some type safety - i.e. if someone changes the property name on a 'Person' the code will break, so it will be obvious. 
        static void StoreData(IGraphClient client, IDataProvider dataProvider)
        {
            foreach (var data in dataProvider.GetExchangeData())
            {
                if (data.From.Email != null)
                { 
                    client.Cypher
                        .Merge($"(from:{Person.Labels} {{email: $frm.{nameof(Person.Email)} }})")
                        .OnCreate().Set($"from.name = $frm.{nameof(Person.Name)}")
                        .Merge($"(e:{Email.Labels} {{id: $data.{nameof(FromExchange.Id)}}})")
                        .OnCreate().Set($"e.subject = $data.{nameof(FromExchange.Subject)},e.conversation = $data.{nameof(FromExchange.Conversation)}, e.sentOn = $data.{nameof(FromExchange.SentOn)}")
                        .Merge($"(from)-[:{RelationshipTypes.Sent}]->(e)")
                        .With("e" )
                        .Unwind(data.To, "to")
                        
                        .Merge($"(t:{Person.Labels} {{email: coalesce(to.{nameof(Person.Email)},'none')}})")
                        .OnCreate().Set($"t.name = to.{nameof(Person.Name)}")
                        .Merge($"(e)-[:{RelationshipTypes.Received}]->(t)")
                        .WithParams(new { data, frm = data.From })
                        .ExecuteWithoutResults();
                }
            }
        }

        static IDataProvider GetDataProvider()
        {
            //Normally this would go to Exchange and get the data.
             //return new TestDataProvider(100);
            return new DataProvider();
        }
    }

    // I use this to 'Type-safe' the types - to prevent accidents!
    public static class RelationshipTypes
    {
        public const string Sent = "SENT";
        public const string Received = "RECEIVED";
    }

    // To allow us to switch real / test data providers
    public interface IDataProvider
    {
        IEnumerable<FromExchange> GetExchangeData();
    }

    // A Test provider - randomly makes users and emails.
    public class TestDataProvider : IDataProvider
    {
        private readonly Random _random = new Random((int)DateTime.Now.Ticks);
        private readonly List<FromExchange> _data = new List<FromExchange>();

        public TestDataProvider(int count)
        {
            for (int i = 0; i < count; i++)
            {
                var fi = _random.Next(count);
                var ft = _random.Next(count);
                if (ft == fi) ft = _random.Next(count);
                var from = new Person { Email = $"user{fi}@testplace.com", Name = $"Test Person_{fi}" };
                var to = new Person { Email = $"user{ft}@testplace.com", Name = $"Test Person_{ft}" };
                _data.Add(new FromExchange { From = from, To = new List<Person> { to }, Id =  Guid.NewGuid().ToString(), Subject = $"subject {_random.Next(count)}" });
            }
        }

        public IEnumerable<FromExchange> GetExchangeData()
        {
            foreach (var t in _data)
                yield return t;
        }
    }


    /// <summary>
    /// Implement this one for actual data
    /// </summary>
    public class DataProvider : IDataProvider
    {
        Application outlookApplication = null;
        NameSpace outlookNamespace = null;
        MAPIFolder inboxFolder = null;
        Items mailItems = null;
        private readonly List<FromExchange> _data = new List<FromExchange>();

        public IEnumerable<FromExchange> GetExchangeData()
        {
            // throw new NotImplementedException();
            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;
                Int32 ii = 0;
                foreach (var inboxItem in mailItems)
                {
                    if (!(inboxItem is MailItem item))
                        continue;

                    var stringBuilder = new StringBuilder();
                    stringBuilder.AppendLine($"From: {item.SenderEmailAddress}");
                    stringBuilder.AppendLine($"To: {item.To}");
                    stringBuilder.AppendLine($"Conversation: {item.ConversationID}");
                    //stringBuilder.AppendLine($"CC: {item.CC}");
                    //stringBuilder.AppendLine($"BCC: {item.BCC}");
                    stringBuilder.AppendLine("");
                    stringBuilder.AppendLine($"Subject: {item.Subject}");
                    
                    Console.WriteLine(stringBuilder);
                    var email = new Email { SentOn = item.SentOn };
                    var from = new Person {Email = item.SenderEmailAddress, Name = item.SenderName};
                    var cc = GetPersonsFromEmailString(item.CC);
                    var to = GetPersonsFromEmailString(item.To);
                    var bcc = GetPersonsFromEmailString(item.BCC);

                    _data.Add(new FromExchange {From = @from, To = to, CC = cc, BCC = bcc, Id = item.EntryID, Conversation = item.ConversationID, Subject = item.Subject, SentOn=item.SentOn});
                    ii++;
                    if (ii > 1000) break;
                    //Marshal.ReleaseComObject(item);
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("Press ENTER to carry on processing.");
                Console.ReadLine();
            }
            finally
            {
               /* ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);*/

            }
            foreach (var t in _data)
                yield return t;
        }
       /* private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }*/

       private IEnumerable<Person> GetPersonsFromEmailString(string emails)
       {
           if (string.IsNullOrWhiteSpace(emails))
               return new Person[0];
           return emails.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(s => new Person {Email = s});
       }
    }
}

    //The output from the 'IDataProvider' GetExchangeData method. Obviously changes to this would mean changes to the Cypher etc.
    public class FromExchange
    {
        public Person From { get; set; }
        public IEnumerable<Person> To { get; set; }
        public IEnumerable<Person> CC { get; set; }
        public IEnumerable<Person> BCC { get; set; }
        public string Conversation { get; set; }

        public string Subject { get; set; }
        public string Id { get; set; }
        public DateTime SentOn { get; set; }
    }

    #region What we're storing
    public class Email
    {
        public const string Labels = "Email";
        public string Id { get; set; }
        public string Conversation { get; set; }
        public string Subject { get; set; }
        public DateTime SentOn { get; set; }
    }

    public class Person
    {
        public const string Labels = "Person";
        public string Name { get; set; }
        public string Email { get; set; }

    }
    #endregion What we're storing
    
   
