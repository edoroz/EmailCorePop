
using System;
using System.Collections.Generic;
using System.IO;
using OpenPop.Mime;
using OpenPop.Pop3;

namespace EmailCorePop {
    
    public class Program {
        static void Main(string[] args) {
            ConnectPop3 oC= new ConnectPop3();
            List<OpenPop.Mime.Message> lstMessages= oC.GetMessages();

            if (lstMessages.Count > 0)
            foreach (var oMessage in lstMessages) {
                Console.WriteLine(oMessage.Headers.Subject + "\t" + oMessage.Headers.From);
                
                    // Get the message body
                MessagePart body = oMessage.FindFirstHtmlVersion();
                if (body == null) {
                    body = oMessage.FindFirstPlainTextVersion();
                }

                if (body != null) {
                    Console.WriteLine("Body: " + body.GetBodyAsText());
                } else {
                    Console.WriteLine("No body found for this message.");
                }
                SaveAttachments(oMessage);

            }
            //Console.ReadLine();
        } //-main

        private static void SaveAttachments(Message message) {
            List<MessagePart> attachments = message.FindAllAttachments();
            foreach (var attachment in attachments) {
                string filePath = Path.Combine("attachments", attachment.FileName);
                Directory.CreateDirectory("attachments");

                using (FileStream stream = new FileStream(filePath, FileMode.Create)) {
                    byte[] bytes = attachment.Body;
                    stream.Write(bytes, 0, bytes.Length);
                }

                Console.WriteLine("Attachment saved: " + filePath);
            }
        }

    } //-cls

    public class ConnectPop3 {
        public string   HostName= "pop.gmail.com";
        public int      port    = 995;
        public bool     useSSL  = true;
        public string   UserName= "playotec@gmail.com";
        public string   Pass    = "peom zebt nmzo vcpm";

        public List<OpenPop.Mime.Message> GetMessages() {
            using (Pop3Client oClient = new Pop3Client()) {
                try {
                    oClient.Connect(HostName, port, useSSL);
                    oClient.Authenticate(UserName, Pass);

                    int messageCount = oClient.GetMessageCount();
                    List<OpenPop.Mime.Message> lstMessages = new List<OpenPop.Mime.Message>();

                    for (int i = messageCount; i > 0; i--) {
                        lstMessages.Add(oClient.GetMessage(i));

                    }
                    return lstMessages;
                } catch (Exception ex) {
                    Console.WriteLine("An error occurred: "+ex + " " + ex.Message);
                    return null;
                }
            }
        } //-mth
    } //-cls
}
