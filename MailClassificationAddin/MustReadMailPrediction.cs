using Microsoft.ML.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailClassificationAddin
{
    public class MustReadMailPrediction
    {
        [ColumnName("PredictedLabel")]
        public Boolean Unread;
    }
}
