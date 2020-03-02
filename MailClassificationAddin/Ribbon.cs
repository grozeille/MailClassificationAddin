using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.ML;
using Microsoft.ML.Data;
using Microsoft.ML.Transforms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailClassificationAddin
{
    public partial class Ribbon
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private IList<MailDto> mailsCache = new List<MailDto>();

        private IDictionary<String, MailDto> mailsCacheByEntryID = new Dictionary<String, MailDto>();

        private MLContext ctx;

        private ITransformer model = null;

        private PredictionEngine<MailDto, MustReadMailPrediction> mustReadPreditionEngine = null;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            // load the existing saved model
            string dataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MailClassificationAddin");
            if (!Directory.Exists(dataFolder))
            {
                Directory.CreateDirectory(dataFolder);
            }

            ctx = new MLContext();
            string modelFile = Path.Combine(dataFolder, "mustReadModel.zip");
            if (File.Exists(modelFile))
            {
                DataViewSchema schema;
                model = ctx.Model.Load(modelFile, out schema);
                mustReadPreditionEngine = ctx.Model.CreatePredictionEngine<MailDto, MustReadMailPrediction>(model);
            }
        }

        private void buttonClassify_Click(object sender, RibbonControlEventArgs e)
        {
            if(mustReadPreditionEngine == null)
            {
                MessageBox.Show("Please train the model first", "Classification", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var context = e.Control.Context;
            if(context != null && context.Selection != null && context.Selection.Count > 0)
            {
                object item = context.Selection[1];
                if (item is MailItem)
                {
                    MailItem mailItem = item as MailItem;
                    MailDto mailDto = toMailDto(mailItem);

                    MustReadMailPrediction mustReadMailPrediction = mustReadPreditionEngine.Predict(mailDto);

                    if(mustReadMailPrediction.Unread)
                    {
                        MessageBox.Show("Nah, just skip it...", "Classification", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("You must read it !", "Classification", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private async void buttonTrain_Click(object sender, RibbonControlEventArgs e)
        {
            var currentFolder = e.Control.Context.CurrentFolder as Outlook.MAPIFolder;
            this.buttonTrain.Enabled = false;
            await Task.Run(() =>
            {
                Training(currentFolder);
            });
            this.buttonTrain.Enabled = true;
            MessageBox.Show("Training completed", "Classification", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Training(Outlook.MAPIFolder folder)
        {
            // To read: https://www.oreilly.com/library/view/machine-learning-for/9781449314835/ch04.html

            // load the cache
            string dataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MailClassificationAddin");
            if(!Directory.Exists(dataFolder))
            {
                Directory.CreateDirectory(dataFolder);
            }
            string dataFile = Path.Combine(dataFolder, "mails.json");
            if(File.Exists(dataFile))
            {
                string json = File.ReadAllText(dataFile);
                mailsCache = JsonConvert.DeserializeObject<List<MailDto>>(json);
                foreach(var m in mailsCache)
                {
                    mailsCacheByEntryID.Remove(m.EntryID);
                    mailsCacheByEntryID.Add(m.EntryID, m);
                }
            }

            // scan for new mails
            string currentAddress = Globals.ThisAddIn.Application.Session.CurrentUser.Address;
            var items = folder.Items;
            items.Sort("sentOn", true);

            Outlook.MAPIFolder sentMails = Globals.ThisAddIn.Application.Session.DefaultStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);

            var now = DateTime.Now;
            var maxDate = now.AddMonths(-1);

            // extract all mails to learn which one is important
            log.Info("Start scanning mails until " + maxDate);
            Stopwatch stopwatch = Stopwatch.StartNew();
            foreach (var item in items)
            {
                var mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    // use the mail item
                    String entryId = mail.EntryID;
                    MailDto mailDto = null;
                    Boolean needToLoadMail = true;
                    if (mailsCacheByEntryID.TryGetValue(entryId, out mailDto))
                    {
                        DateTime lastModificationTime = mail.LastModificationTime;
                        if (lastModificationTime == mailDto.LastModificationTime)
                        {
                            needToLoadMail = false;                            
                        }
                        else
                        {
                            mailsCacheByEntryID.Remove(entryId);
                            mailsCache.Remove(mailDto);
                            mailDto = null;
                        }
                    }

                    if(needToLoadMail)
                    {
                        mailDto = toMailDto(mail);

                        mailsCache.Add(mailDto);
                        mailsCacheByEntryID.Add(mailDto.EntryID, mailDto);
                    }


                    // try if we have any reply to this mail
                    if (!mailDto.Replied)
                    {
                        Boolean replied = false;
                        Conversation conversation = mail.GetConversation();
                        if (conversation != null)
                        {
                            SimpleItems rootItems = conversation.GetRootItems();
                            foreach (var r in rootItems)
                            {
                                SimpleItems childrenItems = conversation.GetChildren(r);
                                foreach (var c in childrenItems)
                                {
                                    var childrenMail = c as Outlook.MailItem;
                                    if (childrenMail.SenderEmailAddress.Equals(currentAddress, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        replied = true;
                                        break;
                                    }
                                }

                                if (replied)
                                {
                                    break;
                                }
                            }
                        }
                        mailDto.Replied = replied;
                    }
                    


                    if (mailDto.SentOn <= maxDate)
                    {
                        break;
                    }                    
                }
            }
            stopwatch.Stop();
            log.Info("Mail scanned, elapsed time: " + stopwatch.ElapsedMilliseconds);

            // update the cache
            string output = JsonConvert.SerializeObject(mailsCache);
            File.WriteAllText(dataFile, output);

            //string trainFile = Path.Combine(dataFolder, "train.tsv");
            //ctx.Data.LoadFromTextFile<MailDto>(trainFile, hasHeader: true);

            // TODO train classification 
            log.Info("Start training");
            
            IDataView trainingDataView = ctx.Data.LoadFromEnumerable<MailDto>(mailsCache);

            if(mailsCache.Count < 10)
            {
                log.Warn("Invalid data for training, not enough mails (<10), skip");
                return;
            }

            int unreadNumberOfValues = mailsCache.Select(_ => _.Unread).Distinct().Count();
            if(unreadNumberOfValues != 2)
            {
                log.Warn("Invalid data for training, they are all read or unread, skip");
                return;
            }

            var pipeline = ctx.Transforms.Conversion.MapValueToKey(inputColumnName: "Unread", outputColumnName: "Label")
                            .Append(ctx.Transforms.Text.FeaturizeText(inputColumnName: "Subject", outputColumnName: "SubjectFeaturized"))
                            .Append(ctx.Transforms.Text.FeaturizeText(inputColumnName: "SenderEmailAddress", outputColumnName: "SenderEmailAddressFeaturized"))
                            .Append(ctx.Transforms.Text.FeaturizeText("BodyFeaturized", new Microsoft.ML.Transforms.Text.TextFeaturizingEstimator.Options
                            {
                                WordFeatureExtractor = new Microsoft.ML.Transforms.Text.WordBagEstimator.Options { NgramLength = 2, UseAllLengths = true },
                                CharFeatureExtractor = new Microsoft.ML.Transforms.Text.WordBagEstimator.Options { NgramLength = 3, UseAllLengths = false },
                            }, "Body"))
                            .Append(ctx.Transforms.Text.FeaturizeText(inputColumnName: "TO", outputColumnName: "TOFeaturized"))
                            .Append(ctx.Transforms.Text.FeaturizeText(inputColumnName: "CC", outputColumnName: "CCFeaturized"))
                            .Append(ctx.Transforms.Concatenate("Features", 
                                "SubjectFeaturized", 
                                "SenderEmailAddressFeaturized", 
                                "BodyFeaturized",
                                "TOFeaturized",
                                "CCFeaturized"))
                            .Append(ctx.Transforms.NormalizeLpNorm("Features", "Features"))
                            .AppendCacheCheckpoint(ctx);

            var trainingPipeline = pipeline
                .Append(ctx.MulticlassClassification.Trainers.SdcaMaximumEntropy("Label", "Features"))
                .Append(ctx.Transforms.Conversion.MapKeyToValue("PredictedLabel"));

            model = trainingPipeline.Fit(trainingDataView);
            string modelFile = Path.Combine(dataFolder, "mustReadModel.zip");
            ctx.Model.Save(model, trainingDataView.Schema, modelFile);

            mustReadPreditionEngine = ctx.Model.CreatePredictionEngine<MailDto, MustReadMailPrediction>(model);
        }

        private static MailDto toMailDto(MailItem mail)
        {
            MailDto mailDto;
            DateTime lastModificationTime = mail.LastModificationTime;
            String subject = mail.Subject;
            String cc = mail.CC;
            String bcc = mail.BCC;
            String senderEmailAddress = mail.SenderEmailAddress;
            String to = mail.To;
            String body = mail.Body;
            Boolean unread = mail.UnRead;
            String conversationID = mail.ConversationID;
            DateTime sentOn = mail.SentOn;
            Boolean isMarkedAsTask = mail.IsMarkedAsTask;
            String entryId = mail.EntryID;

            mailDto = new MailDto();
            mailDto.EntryID = entryId;
            mailDto.BCC = bcc;
            mailDto.CC = cc;
            mailDto.TO = to;
            mailDto.Body = body;
            mailDto.SenderEmailAddress = senderEmailAddress;
            mailDto.Subject = subject;
            mailDto.ConversationID = conversationID;

            mailDto.Unread = unread;
            mailDto.IsMarkedAsTask = isMarkedAsTask;


            mailDto.SentOn = sentOn;
            mailDto.LastModificationTime = lastModificationTime;
            return mailDto;
        }
    }
}
