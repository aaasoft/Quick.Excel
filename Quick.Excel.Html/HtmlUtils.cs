using HtmlAgilityPack;
using Quick.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Quick.Excel.Html
{
    public class HtmlUtils
    {
        /// <summary>
        /// 将表格的html字符串转换为table对象
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        public static table Parse(string html)
        {
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            var tableElement = htmlDoc.DocumentNode;
            List<HtmlNode> trList = new List<HtmlNode>();

            Action<HtmlNodeCollection> trElementHandler = t =>
            {
                if (t == null)
                    return;
                foreach (HtmlNode trElement in t)
                    trList.Add(trElement);
            };

            //table下面的thead
            var nodes_thead = tableElement.SelectNodes("//thead");
            if (nodes_thead != null)
                foreach (HtmlNode theadElement in nodes_thead)
                    trElementHandler(theadElement.SelectNodes("./tr"));
            //table下面的tbody
            var nodes_tbody = tableElement.SelectNodes("//tbody");
            if (nodes_tbody != null)
                foreach (HtmlNode tbodyElement in nodes_tbody)
                    trElementHandler(tbodyElement.SelectNodes("./tr"));
            //table下面的tfoot
            var nodes_tfoot = tableElement.SelectNodes("//tfoot");
            if (nodes_tfoot != null)
                foreach (HtmlNode tbodyElement in nodes_tfoot)
                    trElementHandler(tbodyElement.SelectNodes("./tr"));
            //如果thead,tbody,tfoot里面都没有tr，则全局找tr
            if (trList.Count == 0)
                trElementHandler(tableElement.SelectNodes("//tr"));

            table table = new table();
            foreach (var trElement in trList)
            {
                tr tr = new tr();
                foreach (HtmlNode node in trElement.ChildNodes)
                {
                    if (node.HasChildNodes)
                    {
                        var inputNode = node.SelectNodes("./input")?.FirstOrDefault();
                        if (inputNode != null)
                        {
                            node.InnerHtml = inputNode.GetAttributeValue("value", "");
                        }
                    }
                    td td = null;

                    var nodeValue = System.Web.HttpUtility.HtmlDecode(node.InnerText.Trim() ?? String.Empty);
                    switch (node.Name.ToLower())
                    {
                        case "th":
                            td = new th(nodeValue);
                            break;
                        case "td":
                            td = new td(nodeValue);
                            break;
                    }
                    if (td == null)
                        continue;
                    //读取colspan
                    var tmpStr = node.GetAttributeValue("colspan", null);
                    if (!string.IsNullOrEmpty(tmpStr))
                        td.colspan = int.Parse(tmpStr);
                    //读取rowspan
                    tmpStr = node.GetAttributeValue("rowspan", null);
                    if (!string.IsNullOrEmpty(tmpStr))
                        td.rowspan = int.Parse(tmpStr);
                    //读取bgcolor
                    tmpStr = node.GetAttributeValue("bgcolor", null);
                    if (!string.IsNullOrEmpty(tmpStr))
                        td.background_color = tmpStr;
                    //读取style
                    tmpStr = node.GetAttributeValue("style", null);
                    if (!string.IsNullOrEmpty(tmpStr))
                    {
                        foreach (var line in tmpStr.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            var strArray = line.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                            if (strArray.Length < 2)
                                continue;
                            var key = strArray[0].Trim();
                            var value = strArray[1].Trim();

                            switch (key)
                            {
                                case "text-align":
                                    td.text_align = (text_align)Enum.Parse(typeof(text_align), value);
                                    break;
                                case "color":
                                    td.color = value;
                                    break;
                                case "background-color":
                                    td.background_color = value;
                                    break;
                                case "border-color":
                                    td.border_color = value;
                                    break;
                            }
                        }
                    }
                    tr.Add(td);
                }
                table.Add(tr);
            }
            return table;
        }
    }
}
