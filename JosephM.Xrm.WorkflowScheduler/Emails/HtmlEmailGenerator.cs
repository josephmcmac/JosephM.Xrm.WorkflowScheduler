using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using JosephM.Xrm.WorkflowScheduler.Extentions;
using JosephM.Xrm.WorkflowScheduler.Services;
using System;

namespace JosephM.Xrm.WorkflowScheduler.Emails
{
    public class HtmlEmailGenerator
    {
        public XrmService XrmService { get; set; }
        public string AppendAppIdToUrls { get; }
        public string WebUrl { get; set; }
        private bool IncludeHyperlinks
        {
            get { return !string.IsNullOrWhiteSpace(WebUrl); }
        }
        private StringBuilder _content = new StringBuilder();
        private const string pStyle = "style='font-family: Segoe UI;font-size: 13px;color: #444444;'";
        private const string thStyle = "style='font-family: Segoe UI;font-size: 13px;color: #444444;vertical-align: top;text-align: left; border-style: solid; border-width: 1px; padding: 5px;'";
        private const string tdStyle = "style='font-family: Segoe UI;font-size: 13px;color: #444444;vertical-align: top;text-align: left; border-style: solid; border-width: 1px; padding: 5px;'";
        private const string aStyle = "style='font-family: Segoe UI;font-size: 13px;'";

        public HtmlEmailGenerator(XrmService xrmService, string webUrl, string appendAppIdToUrls)
        {
            XrmService = xrmService;
            AppendAppIdToUrls = appendAppIdToUrls;
            WebUrl = webUrl != null && webUrl.EndsWith("/")
                ? webUrl.Substring(0, webUrl.Length - 1)
                : webUrl;
        }

        public void AppendParagraph(string text)
        {
            _content.AppendLine(string.Format("<p {0}>{1}</p>", pStyle, text));
        }

        //don't change to negative may break queries
        public static int MaximumNumberOfEntitiesToList
        {
            get
            {
                return 100;
            }
        }

        public void AppendTable(IEnumerable<Entity> take, LocalisationService localisationService, IEnumerable<string> fields = null, IDictionary<string, string> aliasTypeMaps = null, bool noHyperLinks = false, string appId = null)
        {
            if (!take.Any())
                return;

            var table = new StringBuilder();
            table.AppendLine("<table style=\"border-collapse: collapse\">");
            table.AppendLine("<thead><tr>");
            if (IncludeHyperlinks && !noHyperLinks)
            {
                table.AppendLine(string.Format("<th {0}></th>", thStyle));
            }
            var firstItem = take.First();
            if (fields == null)
                fields = new[] { XrmService.GetPrimaryNameField(firstItem.LogicalName) };
            foreach (var field in fields)
                AppendThForField(table, firstItem, field, aliasTypeMaps);
            table.AppendLine("</tr></thead>");
            foreach (var item in take.Take(MaximumNumberOfEntitiesToList))
            {
                table.AppendLine("<tr>");
                if (IncludeHyperlinks && !noHyperLinks)
                {
                    table.AppendLine(string.Format("<td {0}>", tdStyle));
                    var url = CreateUrl(item);
                    AppendUrl(url, "View", table);
                    table.AppendLine("</td>");
                }

                foreach (var field in fields)
                    AppendTdForField(table, item, field, localisationService, aliasTypeMaps);
                table.AppendLine("</tr>");
            }
            table.AppendLine("</table>");
            _content.AppendLine(table.ToString());

            if(take.Count() > MaximumNumberOfEntitiesToList)
            {
                AppendParagraph(string.Format("Note this list is incomplete as the maximum of {0} items has been listed", MaximumNumberOfEntitiesToList));
            }
        }

        public string CreateUrl(Entity item)
        {
            return string.Format("{0}/main.aspx?{1}pagetype=entityrecord&etn={2}&id={3}", WebUrl,
                AppendAppIdToUrls == null ? null : ("appid=" + AppendAppIdToUrls + "&"),
                item.LogicalName, item.Id);
        }

        private void AppendThForField(StringBuilder table, Entity firstItem, string field, IDictionary<string, string> aliasTypeMaps)
        {
            var tf = GetTypeAndField(firstItem, field, aliasTypeMaps);
            table.AppendLine(string.Format("<th {0}>{1}</th>", thStyle, XrmService.GetFieldLabel(tf.Field, tf.Type)));
        }

        private TypeAndField GetTypeAndField(Entity entity, string field, IDictionary<string, string> aliasTypeMaps)
        {
            var entityType = entity.LogicalName;
            if (field.Contains('.'))
            {
                var split = field.Split('.');
                var alias = split.First();
                if (aliasTypeMaps == null || !aliasTypeMaps.ContainsKey(alias))
                {
                    throw new Exception(string.Format("Cannot Determine Field's Entity Type. The Field {0} Has An Alias Prefix But The Method Argument {1} Does Not Contain A Map To The Entity Type", field, nameof(aliasTypeMaps)));
                }
                entityType = aliasTypeMaps[alias];
                field = split.ElementAt(1);
            }
            return new TypeAndField(entityType, field);
        }

        private class TypeAndField
        {
            public string Type { get; set; }
            public string Field { get; set; }

            public TypeAndField(string type, string field)
            {
                Type = type;
                Field = field;
            }
        }

        private void AppendTdForField(StringBuilder table, Entity item, string field, LocalisationService localisationService, IDictionary<string, string> aliasTypeMaps)
        {
            var tf = GetTypeAndField(item, field, aliasTypeMaps);
            table.AppendLine(string.Format("<td {0}>", tdStyle));
            table.Append(XrmService.GetFieldAsDisplayString(tf.Type, tf.Field, item.GetFieldValue(field), localisationService.TimeZonename));
            table.AppendLine("</td>");
        }

        private void AppendUrl(string url, string label, StringBuilder sb)
        {
            sb.Append(CreateHyperlink(url, label));
        }

        public string CreateHyperlink(string url, string label)
        {
            return string.Format("<a {0} href='{1}' >{2}</a>", aStyle, url, label);
        }

        public string GetContent()
        {
            return _content.ToString();
        }
    }
}