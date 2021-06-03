using Quick.Excel.Model;
using System;
using System.Collections.Generic;

namespace Quick.Excel
{
    public class td
    {
        /// <summary>
        /// 是否是粗体
        /// </summary>
        public bool bold { get; set; } = false;

        public int colspan { get; set; } = 1;

        public int rowspan { get; set; } = 1;

        public text_align text_align { get; set; } = text_align.initial;
        /// <summary>
        /// 字体颜色
        /// </summary>
        public string color { get; set; }
        /// <summary>
        /// 获取字体颜色
        /// </summary>
        public byte[] color_bytes
        {
            get
            {
                if (string.IsNullOrEmpty(color))
                    return null;
                return convertColorString(color);
            }
        }
        /// <summary>
        /// 背景颜色
        /// </summary>
        public string background_color { get; set; }
        /// <summary>
        /// 获取背景颜色
        /// </summary>
        public byte[] background_color_bytes
        {
            get
            {
                if (string.IsNullOrEmpty(background_color))
                    return null;
                return convertColorString(background_color);
            }
        }

        /// <summary>
        /// 边框颜色
        /// </summary>
        public string border_color { get; set; }
        /// <summary>
        /// 获取边框颜色
        /// </summary>
        public byte[] border_color_bytes
        {
            get
            {
                if (string.IsNullOrEmpty(border_color))
                    return null;
                return convertColorString(border_color);
            }
        }

        private byte[] convertColorString(string color)
        {
            if (color.StartsWith("#"))
            {
                var str = color.Substring(1);
                if (str.Length == 6)
                {
                    List<byte> list = new List<byte>();
                    for (var i = 0; i < str.Length; i += 2)
                    {
                        list.Add(byte.Parse(str.Substring(i, 2), System.Globalization.NumberStyles.HexNumber));
                    }
                    return list.ToArray();
                }
            }
            return null;
        }

        /// <summary>
        /// 值
        /// </summary>
        public string value { get; set; }

        public td(string value)
        {
            this.value = value;
        }
    }
}
