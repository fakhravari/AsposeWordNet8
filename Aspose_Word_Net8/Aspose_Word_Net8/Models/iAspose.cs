using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.MailMerging;
using System.Drawing;
using System.Security.Cryptography.X509Certificates;

namespace Aspose_Word_Net8.Models
{
    public class LicenseSetupService : IHostedService
    {
        private readonly ILogger<LicenseSetupService> _logger;

        public LicenseSetupService(ILogger<LicenseSetupService> logger)
        {
            _logger = logger;
        }

        public Task StartAsync(CancellationToken cancellationToken)
        {
            string licensePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "PigisFileServer", "Aspose.Total.NET.lic");

            try
            {
                new Aspose.Words.License().SetLicense(licensePath);
            }
            catch (Exception ex)
            {

            }
            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            return Task.CompletedTask;
        }
    }

    public class iAspose
    {
        private readonly IConfiguration _configuration;

        public iAspose(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public static string templateurl = "";
        public static string BuildPrintLetter(string[] field, object[] data)
        {
            string template = "LT_HavooRush.docx";

            string binPath = AppContext.BaseDirectory;
            string projectRoot = Directory.GetParent(Directory.GetParent(Directory.GetParent(Directory.GetParent(binPath).FullName).FullName).FullName).FullName;
            string templatePath = Path.Combine(projectRoot, "wwwroot", "PigisFileServer", template);

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException("The template file was not found.", templatePath);
            }

            string saveDocPath = Path.Combine(projectRoot, "wwwroot", "PigisFileServer", "Temp");
            if (!Directory.Exists(saveDocPath))
            {
                Directory.CreateDirectory(saveDocPath);
            }

            Document my = new Document(templatePath)
            {
                MailMerge =
                {
                    FieldMergingCallback = new HandleMergeFieldAlternatingRows()
                }
            };
            my.MailMerge.Execute(field, data);

            string savedFilePathName = "PIGIS_Doc_" + DateTime.Now.Millisecond + ".docx";
            string fullPath = Path.Combine(saveDocPath, savedFilePathName);
            my.Save(fullPath);

            return savedFilePathName;
        }

        private class HandleMergeFieldAlternatingRows : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                try
                {
                    if (mBuilder == null)
                        mBuilder = new DocumentBuilder(e.Document);

                    if (e.FieldName.Equals("Shomareh"))
                    {
                        mBuilder.MoveToMergeField(e.FieldName);
                        mBuilder.Font.StyleIdentifier = StyleIdentifier.DefaultParagraphFont;

                        Run run = new Run(mBuilder.Document);
                        run.Text = e.FieldValue.ToString();

                        Aspose.Words.Font font = run.Font;
                        font.Bidi = false;
                        if (!templateurl.Contains("English"))
                        {
                            font.Bold = true;
                            font.Name = "B Nazanin";
                            font.Size = 9;
                            font.Color = Color.Black;
                        }
                        else
                        {
                            font.Name = "Arial";
                            font.Size = 9;
                            font.Color = Color.Black;
                        }
                        mBuilder.InsertNode(run);
                    }
                    if (e.FieldName.Equals("Tarikh"))
                    {
                        mBuilder.MoveToMergeField(e.FieldName);
                        mBuilder.Font.StyleIdentifier = StyleIdentifier.DefaultParagraphFont;

                        Run run = new Run(mBuilder.Document);
                        run.Text = e.FieldValue.ToString();

                        Aspose.Words.Font font = run.Font;

                        if (!templateurl.Contains("English"))
                        {
                            font.Bold = true;
                            font.Name = "B Nazanin";
                            font.Size = 12;
                            font.Color = Color.Black;
                        }
                        else
                        {

                            font.Name = "Arial";
                            font.Size = 11;
                            font.Color = Color.Black;
                        }
                        mBuilder.InsertNode(run);
                    }
                    if (e.FieldName.Equals("Semat"))
                    {
                        mBuilder.MoveToMergeField(e.FieldName);
                        mBuilder.Font.StyleIdentifier = StyleIdentifier.DefaultParagraphFont;

                        Run run = new Run(mBuilder.Document);
                        run.Text = e.FieldValue.ToString();
                        Aspose.Words.Font font = run.Font;
                        font.Bold = true;
                        font.Name = "B Nazanin";
                        font.Size = 12;
                        font.Color = Color.Black;
                        mBuilder.InsertNode(run);
                    }
                    if (e.DocumentFieldName.StartsWith("Html"))
                    {
                        DocumentBuilder builder = new DocumentBuilder(e.Document);
                        builder.MoveToMergeField(e.DocumentFieldName);
                        builder.InsertHtml((string)e.FieldValue);
                        e.Text = "";
                    }
                    if (e.DocumentFieldName.StartsWith("Image"))
                    {
                        DocumentBuilder builder = new DocumentBuilder(e.Document);
                        builder.MoveToMergeField(e.DocumentFieldName);
                        Shape shape = builder.InsertImage(e.FieldValue.ToString());
                        shape.WrapType = WrapType.None;
                        shape.BehindText = true;
                        shape.Width = 135;
                        shape.Height = 110;
                    }
                }
                catch (Exception ex)
                {
                    string Field = e.FieldName;
                    string Value = e.FieldValue.ToString();
                }
            }
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // throw new NotImplementedException();
            }

            public DocumentBuilder mBuilder { set; get; }
        }
    }
}
