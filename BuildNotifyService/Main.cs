using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using System;
using System.DirectoryServices;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;

namespace BuildNotifyService
{
	class Program
	{
		public static void Main(string[] args)
		{
            // Lade Exchange Service Informationen aus Config-Datei
            Config config = LoadConfigInformation();

            Uri uri = new Uri(config.Url);
			var tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(uri);
			IBuildServer buildService = tfs.GetService(typeof(IBuildServer)) as IBuildServer;

			if (buildService != null)
			{

                // Erstelle Build Detail Spezification, um die Antwort der Builds-Abfrage einzuschränken
				// Maximal einen Build per Definition
				// Sortiert nach FinishTime
				// Keine zusätzlichen Informationen laden (z.B. Error Details, etc.)
				// Nur Builds aus den letzten 24 Stunden
				IBuildDetailSpec buildDetailSpec = buildService.CreateBuildDetailSpec(config.TeamProject, "*");
				{
					buildDetailSpec.MaxBuildsPerDefinition = 1;
					buildDetailSpec.QueryOrder = BuildQueryOrder.FinishTimeDescending;
					buildDetailSpec.InformationTypes = null;
					buildDetailSpec.MinFinishTime = DateTime.Now.AddHours(-24.0);
				}

				var htmlBody = string.Empty;
				var buildQueryResult = buildService.QueryBuilds(buildDetailSpec);				
				foreach (var build in buildQueryResult.Builds)
				{
					// Filtere alle NICHT scheduled Builds aus
                    if ((build.BuildDefinition.ContinuousIntegrationType == ContinuousIntegrationType.Schedule || 
                        build.BuildDefinition.ContinuousIntegrationType == ContinuousIntegrationType.ScheduleForced))
                    {
						// Erstelle HTML Body mit Build informationen
						htmlBody += CreateHtmlBody(build, tfs);
                    }
				}

				var htmlFile = CreateFullHtml(htmlBody);
                
				// Sende Information-Mail an alle Entwickler
				SendMail(htmlFile, config);
			}
		}

		private static string CreateHtmlBody(IBuildDetail build, TfsTeamProjectCollection tfs)
		{
			// Lade Html Vorlage aus den Ressourcen
			var assembly = Assembly.GetExecutingAssembly();
			var htmlStream = assembly.GetManifestResourceStream("BuildNotifyService.ItemTemplate.html");
			var reader = new StreamReader(htmlStream);
			var htmlTemplate = reader.ReadToEnd();

			// Lade Build Details URL des Web Access, um einen Hyperlink erzeugen zu können
			Uri detailsUrl = null;
			TswaClientHyperlinkService service = tfs.GetService<TswaClientHyperlinkService>();
			if (service != null)
			{
				detailsUrl = service.GetViewBuildDetailsUrl(build.Uri);
			}

			return string.Format(
				htmlTemplate, 
				SetStatusImagePlaceholder(build.Status), 
				build.BuildDefinition.Name, 
				build.StartTime.ToString("dd.MMM.yyyy HH:mm"), 
				build.FinishTime.ToString("dd.MMM.yyyy HH:mm"), 
				detailsUrl
				);
		}

		private static string SetStatusImagePlaceholder(BuildStatus buildStatus)
		{
            if(buildStatus == BuildStatus.Succeeded)
            {
                return "cid:picOK";
            }
			else if (buildStatus == BuildStatus.Failed)
			{
				return "cid:picNotOK";
			}
			else if (buildStatus == BuildStatus.PartiallySucceeded)
			{
				return "cid:picPartially";
			}
            else if (buildStatus == BuildStatus.Stopped)
            {
                return "cid:picStopped";
            }
			else
			{
				return "cid:picUnknown";
			}
		}

		private static string CreateFullHtml(string htmlBody)
		{
			if (string.IsNullOrWhiteSpace(htmlBody))
				return "<html><body>Error: can not create the eMail HTML-Body</body></html>";
			else
			{
				StringBuilder htmlBuilder = new StringBuilder();
				htmlBuilder.Append("<html>");
				htmlBuilder.Append("<head>");
				htmlBuilder.Append("<title>");
				htmlBuilder.Append("Page-");
				htmlBuilder.Append(Guid.NewGuid().ToString());
				htmlBuilder.Append("</title>");
				htmlBuilder.Append("</head>");
				htmlBuilder.Append("<body>");
				htmlBuilder.Append(@"<span style=""color: #0000ff""><span style=""font-size: 24px""><strong>Der Build Status der letzten Nacht!</strong></span></span> <br> <br>");
				htmlBuilder.Append(@"<table border=""1px"" cellpadding=""5"" cellspacing=""0"" >");
				htmlBuilder.Append(@"<style=""border: solid 1px Black; font-size: small;"">");

				//Create Header Row
				htmlBuilder.Append(@"<tr align=""left"" valign=""top"">");
				htmlBuilder.Append(@"<td>" + htmlBody + "</td>");
				htmlBuilder.Append("</tr>");

				//Create Bottom Portion of HTML Document
				htmlBuilder.Append("</table>");
				htmlBuilder.Append("<br><br><br>");
				htmlBuilder.Append("</body>");
				htmlBuilder.Append("</html>");

				//Create String to be Returned
				return htmlBuilder.ToString();
			}
		}

		private static Config LoadConfigInformation()
		{
			// Lade die Mail Informationen aus dem Config-XML-File
			Config conf = new Config();
			XmlDocument xmlDoc = new XmlDocument(); 
			xmlDoc.Load("config.xml"); 

			conf.FromMail = xmlDoc.GetElementsByTagName("FromMail")[0].InnerText;
			conf.ToMail = xmlDoc.GetElementsByTagName("ToMail")[0].InnerText;
			conf.OutputPath = xmlDoc.GetElementsByTagName("OutputPath")[0].InnerText;
			conf.SuccessImgPath = xmlDoc.GetElementsByTagName("SuccessImgPath")[0].InnerText;
			conf.FailedImgPath = xmlDoc.GetElementsByTagName("FailedImgPath")[0].InnerText;
			conf.PartiallyImgPath = xmlDoc.GetElementsByTagName("PartiallyImgPath")[0].InnerText;
            conf.DefaultImgPath = xmlDoc.GetElementsByTagName("DefaultImgPath")[0].InnerText;
            conf.StoppedImgPath = xmlDoc.GetElementsByTagName("StoppedImgPath")[0].InnerText;
            conf.TeamProject = xmlDoc.GetElementsByTagName("TeamProject")[0].InnerText;
            conf.Url = xmlDoc.GetElementsByTagName("TFSUrl")[0].InnerText;

			return conf;
		}

		private static void SendMail(string html, Config config)
		{
			// Initialisiere Exchange Service
			ExchangeService ews = InitializeExchangeService();
			  
			// Erzeuge Nachricht
			EmailMessage mail = new EmailMessage(ews);
			mail.Subject = "Nightly Build-Status";
			mail.Body = html;
			mail.Body.BodyType = BodyType.HTML;
			mail.From = new EmailAddress(config.FromMail);

            if (config.ToMail.Contains(";"))
            {
                string[] splitter = {";"};
                string[] recipients = config.ToMail.Split(splitter, 5, StringSplitOptions.None);
                foreach (var item in recipients)
                {
                    mail.ToRecipients.Add(item);
                }
            }
            else
			    mail.ToRecipients.Add(config.ToMail);

			// Füge Bilder als Attachement an
			AddImagesAsAttachments(mail);

			// Sende Email
			mail.SendAndSaveCopy();
		}

		private static void AddImagesAsAttachments(EmailMessage mail)
		{
			var appPath = Environment.CurrentDirectory;
			var picOk = Path.Combine(appPath, "BuildSuccess.png");
			var picNotOk = Path.Combine(appPath, "BuildFailed.png");
			var picUnknown = Path.Combine(appPath, "BuildDefault.png");
            var picPartially = Path.Combine(appPath, "BuildPartiallySucceeded.png");
            var picStopped = Path.Combine(appPath, "BuildStopped.png");

			var html = mail.Body.Text;
			if (html.IndexOf("cid:picOK") > -1)
			{
				byte[] theBytes = File.ReadAllBytes(picOk);
                var attachment = mail.Attachments.AddFileAttachment("BuildSuccess.png", theBytes);
				attachment.ContentId = "picOk";
			}

			if (html.IndexOf("cid:picNotOK") > -1)
			{
				byte[] theBytes = File.ReadAllBytes(picNotOk);
                var attachment = mail.Attachments.AddFileAttachment("BuildFailed.png", theBytes);
				attachment.ContentId = "picNotOK";
			}

            if (html.IndexOf("cid:picPartially") > -1)
            {
                byte[] theBytes = File.ReadAllBytes(picPartially);
                var attachment = mail.Attachments.AddFileAttachment("BuildPartiallySucceeded.png", theBytes);
                attachment.ContentId = "picPartially";
            }

            if (html.IndexOf("cid:picStopped") > -1)
            {
                byte[] theBytes = File.ReadAllBytes(picStopped);
                var attachment = mail.Attachments.AddFileAttachment("BuildStopped.png", theBytes);
                attachment.ContentId = "picStopped";
            }

			if (html.IndexOf("cid:picUnknown") > -1)
			{
				byte[] theBytes = File.ReadAllBytes(picUnknown);
                var attachment = mail.Attachments.AddFileAttachment("BuildDefault.png", theBytes);
				attachment.ContentId = "picUnknown";
			}
		}

        private static ExchangeService InitializeExchangeService()
        {
            // Initialisiere den ExchangeService 
            ExchangeService eService = new ExchangeService();
            
            // Benutze default Credentials
            eService.UseDefaultCredentials = true;
            
            eService.AutodiscoverUrl("john.doe@almsports.net", RedirectionURLValidationCallback);

            return eService;
        }

        private static bool RedirectionURLValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;

        }

	}
}
