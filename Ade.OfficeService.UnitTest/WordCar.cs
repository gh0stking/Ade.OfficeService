using Ade.OfficeService.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.UnitTest
{
    public class WordCar
    {
        [PlaceHolder(PlaceHolderEnum.A)]
        public string OwnerName { get; set; }

        [PlaceHolder(PlaceHolderEnum.B)]
        public string CarType { get; set; }

        //图片占位的属性类型必须为List<string>,存放图片的绝对全地址
        [PicturePlaceHolder(PlaceHolderEnum.C,"车辆照片")]
        public List<string> CarPictures { get; set; }

        [PicturePlaceHolder(PlaceHolderEnum.D,"车辆证件")]
        public List<string> CarLicense { get; set; }
    }
}
