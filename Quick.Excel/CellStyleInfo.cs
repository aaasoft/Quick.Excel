using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Quick.Excel
{
    internal static class ModelUtils
    {
        /// <summary>
        /// 判断两个对象是否相等
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="self"></param>
        /// <param name="obj"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static bool Equals<T>(this T self, object obj, params Func<T, object>[] parameters)
            where T : class
        {
            T toCompare = obj as T;
            if (toCompare == null)
            {
                return false;
            }
            foreach (var parameter in parameters)
                if (!Object.Equals(parameter(self), parameter(toCompare)))
                    return false;
            return true;
        }

        /// <summary>
        /// 计算哈希值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="self"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static int GetHashCode<T>(this T self, params Func<T, object>[] parameters)
        {
            return GetHashCode(parameters.Select(parameter => parameter(self)).ToArray());
        }

        /// <summary>
        /// 计算哈希值
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static int GetHashCode(params object[] parameters)
        {
            int hashCode = 13;
            if (parameters != null && parameters.Length > 0)
                foreach (var parameter in parameters)
                    if (parameter != null)
                        hashCode = (hashCode * 7) + parameter.GetHashCode();
            return hashCode;
        }
    }

    public class CellStyleInfo
    {
        public text_align text_align { get; set; }
        public bool bold { get; set; }
        public string color { get; set; }
        public string background_color { get; set; }
        public string border_color { get; set; }
        public CellStyleInfo() { }
        public CellStyleInfo(td td)
        {
            text_align = td.text_align;
            bold = td.bold;
            color = td.color;
            background_color = td.background_color;
            border_color = td.border_color;
        }

        public override int GetHashCode()
        {
            return this.GetHashCode(
    t => t.text_align,
    t => t.bold,
    t => t.color,
    t => t.background_color,
    t => t.border_color);
        }

        public override bool Equals(object obj)
        {
            return this.Equals(obj,
    t => t.text_align,
    t => t.bold,
    t => t.color,
    t => t.background_color,
    t => t.border_color);
        }
    }
}
