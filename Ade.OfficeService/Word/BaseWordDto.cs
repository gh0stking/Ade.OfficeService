using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace Ade.OfficeService.Word
{
    public class BaseWordDto
    {
        /// <summary>
        /// 获取渲染后的Word
        /// </summary>
        /// <param name="contentRootPath"></param>
        /// <returns></returns>
        public XWPFDocument GetRenderedWord(string templateUrl)
        {
            XWPFDocument word = GetTemplateWord(templateUrl);

            Render(word);

            CustomOperation(word);

            return word;
        }

        /// <summary>
        /// 图片Id
        /// </summary>
        protected uint PicId
        {
            get
            {
                picId++;
                return picId;
            }
        }

        private uint picId = 0;

        /// <summary>
        /// 子类在渲染模板之后的自定义操作
        /// </summary>
        /// <param name="word"></param>
        protected virtual void CustomOperation(XWPFDocument word)
        {
        }

        /// <summary>
        /// 获取插入图片配置信息
        /// </summary>
        /// <returns></returns>
        protected virtual List<AddPictureOptions> GetAddPictureOptionsList()
        {
            List<AddPictureOptions> listAddPictureOptions = new List<AddPictureOptions>();
            Type type = this.GetType();
            PropertyInfo[] props = type.GetProperties();

            List<string> listPictureUrl;
            string picName;
            foreach (PropertyInfo prop in props)
            {
                if (prop.IsDefined(typeof(PicturePlaceHolderAttribute)))
                {
                    try
                    {
                        listPictureUrl = (List<string>)prop.GetValue(this);
                    }
                    catch (Exception)
                    {
                        throw new Exception("图片占位符必须为字符串集合类型");
                    }

                    picName = prop.GetCustomAttribute<PicturePlaceHolderAttribute>().ImageType.ToString();
                    for (int i = 0; i < listPictureUrl.Count; i++)
                    {
                        string picUrl = listPictureUrl[i];
                        listAddPictureOptions.Add(new AddPictureOptions()
                        {
                            PicId = PicId,
                            PictureName = $"{picName}_{i + 1}",
                            LocalPictureUrl = picUrl,
                            PlaceHolder = prop.GetCustomAttribute<PicturePlaceHolderAttribute>().PlaceHolder,
                            Extension = WordHelper.GetRemoteFileExtention(listPictureUrl[i]),
                            ImageType = prop.GetCustomAttribute<PicturePlaceHolderAttribute>().ImageType
                        });
                    }
                }
            }

            return listAddPictureOptions;
        }

        /// <summary>
        /// 替换占位符钩子
        /// </summary>
        /// <param name="run"></param>
        protected virtual void HookOfRenderTextRun(XWPFRun run)
        {
        }

        /// <summary>
        /// 获取模板文件
        /// </summary>
        /// <param name="contentRootPath"></param>
        /// <returns></returns>
        private XWPFDocument GetTemplateWord(string templateUrl)
        {
            XWPFDocument word;

            Type type = this.GetType();
           
            if (!File.Exists(templateUrl))
            {
                throw new Exception("template not found");
            }

            try
            {
                using (FileStream fs = File.OpenRead(templateUrl))
                {
                    word = new XWPFDocument(fs);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("fail to open template");
            }

            return word;
        }

        /// <summary>
        /// 渲染Word文件
        /// </summary>
        /// <param name="word"></param>
        private void Render(XWPFDocument word)
        {
            if (word == null)
            {
                throw new ArgumentNullException("word");
            }

            Dictionary<string, string> placeHolderAndValueDict = GetPlaceHolderAndValueDict();

            List<AddPictureOptions> listAddPictureOptions = GetAddPictureOptionsList();

            List<string> listAllPlaceHolder = GetAllPlaceHolder();

            WordHelper.ReplacePlaceHolderInWord(word, placeHolderAndValueDict, listAddPictureOptions, listAllPlaceHolder, HookOfRenderTextRun);
        }

        private List<string> GetAllPlaceHolder()
        {
            List<string> listAllPlaceHolder = new List<string>();
            Type type = this.GetType();
            PropertyInfo[] props = type.GetProperties();

            foreach (PropertyInfo prop in props)
            {
                if (prop.IsDefined(typeof(PlaceHolderAttribute)))
                {
                    listAllPlaceHolder.Add(prop.GetCustomAttribute<PlaceHolderAttribute>().PlaceHolder.ToString());
                }

                if (prop.IsDefined(typeof(PicturePlaceHolderAttribute)))
                {
                    listAllPlaceHolder.Add(prop.GetCustomAttribute<PicturePlaceHolderAttribute>().PlaceHolder.ToString());
                }
            }

            return listAllPlaceHolder;
        }

        /// <summary>
        /// 获取展位符和值字典
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, string> GetPlaceHolderAndValueDict()
        {
            Dictionary<string, string> placeHolderAndValueDict = new Dictionary<string, string>();
            Type type = this.GetType();
            PropertyInfo[] props = type.GetProperties();

            foreach (PropertyInfo prop in props)
            {
                if (prop.IsDefined(typeof(PlaceHolderAttribute)))
                {
                    placeHolderAndValueDict.Add(prop.GetCustomAttribute<PlaceHolderAttribute>().PlaceHolder.ToString(), prop.GetValue(this)?.ToString());
                }
            }

            return placeHolderAndValueDict;
        }
    }
}
