using Microsoft.ML.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailClassificationAddin
{
    public class MailDto
    {
        [LoadColumn(0)]
        public String EntryID { get; set; }
        
        [LoadColumn(1)]
        public String Subject { get; set; }
        
        [LoadColumn(2)]
        public String Body { get; set; }

        [LoadColumn(3)] 
        public String CC { get; set; }

        [LoadColumn(4)] 
        public String BCC { get; set; }

        [LoadColumn(5)] 
        public String TO { get; set; }

        [LoadColumn(6)] 
        public String SenderEmailAddress { get; set; }

        [LoadColumn(7)] 
        public String ConversationID { get; set; }

        [LoadColumn(8)] 
        public Boolean Unread { get; set; }

        [LoadColumn(9)] 
        public Boolean IsMarkedAsTask { get; set; }

        [LoadColumn(10)] 
        public Boolean Replied { get; set; }

        [LoadColumn(11)] 
        public DateTime SentOn { get; set; }

        [LoadColumn(12)] 
        public DateTime LastModificationTime { get; set; }
    }
}
