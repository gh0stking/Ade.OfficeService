using Ade.OfficeService.Word;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace Ade.OfficeService.UnitTest
{
    public class WordTest
    {
        [Fact]
        public void TemplateRenderTest()
        {
            string curDir = Environment.CurrentDirectory;
            string pic1 = Path.Combine(curDir, "files", "1.jpg");
            string pic2 = Path.Combine(curDir, "files", "2.jpg");
            string pic3 = Path.Combine(curDir, "files", "3.jpg");
            string templateurl = Path.Combine(curDir, "files", "CarWord.docx");

            WordCar car = new WordCar()
            {
                OwnerName = "龚英韬",
                CarType = "豪华型宾利",
                CarPictures = new List<string> { pic1, pic2 },
                CarLicense = new List<string> { pic3 }
            };

            XWPFDocument doc = WordExportService.ExportFromTemplate(templateurl, car);


            File.WriteAllBytes(@"c:\test\test.docx", doc.ToBytes());
        }
       
    }
}
