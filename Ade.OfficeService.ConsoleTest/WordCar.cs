using Ade.OfficeService.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.ConsoleTest
{
    public class WordCar : IWordExport
    {
        [PlaceHolder(PlaceHolderEnum.A)]
        public string OwnerName { get; set; }

        [PlaceHolder(PlaceHolderEnum.B)]
        public string CarType { get; set; }

        [PicturePlaceHolder(PlaceHolderEnum.C,"车辆照片")]
        public List<string> CarPictures { get; set; }

        [PicturePlaceHolder(PlaceHolderEnum.D,"车辆证件")]
        public List<string> CarLicense { get; set; }
    }
}
