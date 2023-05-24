using ClassLibraryNetCore;
using Docentric.Documents.ObjectModel;
using Docentric.Documents.Reporting;
using Docentric.Drawing;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocentricConsoleApplicationNetCore
{
    public class ImageGeneration : ReportBase
	{
		public ImageGeneration()
		{
			Name = "Image";
		}
		public SaveOptions SaveOptions
		{
			get;
			set;
		}

		public override string GenerateReport()
		{
			ImageData imageData = new();
			string content = "<p style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Main risk to the production plan is the sensitivity of the schedule and the requirement to deliver the planned waste volumes month on month.  Significant Rainfall Events occurred across the site in 2022.</p>\r\n<p style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Assumptions for loss calculations are as follows:</p>\r\n<ul>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> With the completion of Loders pit (late 2020) all production now comes from the North and West pits.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> Total waste removal for 2023 is 113.85Mbcm.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> Total coal mined (ROM) is 16.7Mt and total product coal is 11.2Mt.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> Dragline prime is forecast for 2023 at 12.1Mbcm, truck / shovel is 33.4Mbcm, excavator /truck is 48.8Mbcm, dozer push / cast blasting at 5.4Mbcm for a total of 99.7Mbcm total prime in 2023.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> Rehandle forecast for 2023 - Dragline is 10.0Mbcm, truck / shovel is 0.6Mbcm, excavator /truck is 1.2Mbcm for a total of 11.9Mbcm.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> Total Plant capacity is 12.5Mt of product coal, depending on the coal types available and some bypass is possible.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">The North CHPP handles approximately 60% of MTW product. The South CHPP handles approximately 40% of MTW product. If product from the pits is suitable, more coal can be sent to the south plant where the 2 products can be processed.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Coking coal makes up approximately 15% of the total product and thermal coal makes up the remainder.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> The South Plant can be ramped up to handle 50% of site production if required.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> The operation now has two (2) pits being mined (North, and West) at any one time providing some flexibility on where coal can be drawn.</li>\r\n</ul>\r\n<ul>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Dragline prime movements for 2023 are forecast as follows:\r\n<ol>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> Marion 8200 (#101): 5.5Mbcm.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> BE 1370 (#102): Parked up - nil.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> P&amp;H 9020 (#103): 6.6Mbcm.</li>\r\n</ol>\r\n</li>\r\n</ul>\r\n<ul>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Shovel prime movements for 2023 are forecast as follows:\r\n<ol>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> P&amp;H 4100 (S342): 0.4Mbcm.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> P&amp;H 4100XPC (S345): 19.4Mbcm.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> P&amp;H 4100XPC (S344): 13.6Mbcm.</li>\r\n</ol>\r\n</li>\r\n</ul>\r\n<ul>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Excavator prime movements for 2023 are forecast as follows:\r\n<ol>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Liebherr R9800 (EX321): 12.7Mbcm.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\"> Liebherr R9800 (EX322): 13.7Mbcm.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Hitachi EX3600 x 3 = 10.0Mbcm prime waste.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Liebherr 9400 x 2 = 6.5Mbcm prime waste</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">2 x Hitachi EX5500/5600 = 4.5Mbcm total for 2 machines.</li>\r\n</ol>\r\n</li>\r\n</ul>\r\n<ul>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Dragline prime movements for 2023 are forecast as follows:\r\n<ol>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Dragline costs are an average of USD2.12 per bcm while shovel/truck costs are an average of USD3.37 per bcm and Excavators are $4.10 per bcm.</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">9800  $25M 18 Month lead time</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">Hitachi 5600  $14M  12 month lead time</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">4100 XPC $52M EST  18 -24 month lead time</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">NHL   $8M   24 month lead time</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">830E  $7.5M  12 month lead time</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">800Ton Digger Hire $1m   ( $2000 hr x 500hrs ) a month</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">600Ton   Digger Hire   $750K per month</li>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">330Ton Truck  Hire $250K per month ($500hr x 500).</li>\r\n</ol>\r\n</li>\r\n</ul>\r\n<ul>\r\n<li style=\"font-family:Calibri;font-size:15px;text-align:justify;\">There is limited pump and pipes capacity at MTW, a new dam (NOOP - North Out Of Pit)  is being developed and has the permit to discharge.</li>\r\n</ul>\r\n";

            HtmlDocument document = new() { OptionWriteEmptyNodes = true };
            //string htmlDocpath = Path.Combine(ApplicationPath, "Resources\\" + Name + "\\HtmlFragment.html");
			document.LoadHtml(content);
			var htmlData = document.DocumentNode.OuterHtml;
            imageData.Data = Encoding.UTF8.GetBytes(htmlData);

            //imageData.Data = File.ReadAllBytes(System.IO.Path.Combine(ApplicationPath, "Resources\\" + Name + "\\HtmlFragment.html"));
            imageData.IsDataAvailable = true;
			// Create a temporary file for the generated report document.
			string reportDocumentFileName = System.IO.Path.GetTempPath() + "GeneratedReport_" + Guid.NewGuid().ToString() + OutputDocumentFileExtension;
			using (Stream reportDocumentStream = File.Create(reportDocumentFileName))
			{
				// Open the report template file.
				using (Stream reportTemplateStream = GetReportTemplate())
				{
					// Generate the report document using 'DocumentGenerator'.
					DocumentGenerator dg = new DocumentGenerator(imageData);
					//dg.SetNamedDataSourceValue("HtmlFragment", imageData.Data);
					FontManager.SetFontsSources(new FolderFontSource("C:\\Windows\\Fonts"));
					DocumentGenerationResult result = dg.GenerateDocument(reportTemplateStream, reportDocumentStream, SaveOptions.Word);
					if (result.HasErrors)
					{
						foreach (Error error in result.Errors) Console.Out.WriteLine(error.Message);
					}
				}
			}
			return reportDocumentFileName;
		}
	}
}

