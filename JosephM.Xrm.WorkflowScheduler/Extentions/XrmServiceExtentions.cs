using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using System.Linq;

namespace JosephM.Xrm.WorkflowScheduler.Extentions
{
    public static class XrmServiceExtentions
    {
        /// <summary>
        /// Returns list of key values giving the types and field name parsed for the given string of field joins
        /// key = type, value = field
        /// </summary>
        /// <param name="xrmService"></param>
        /// <param name="fieldPath"></param>
        /// <param name="sourceType"></param>
        /// <returns></returns>
        public static IEnumerable<KeyValuePair<string,string>> GetTypeFieldPath(this XrmService xrmService, string fieldPath, string sourceType)
        {

            var list = new List<KeyValuePair<string, string>>();
            var splitOutFunction = fieldPath.Split(':');
            if (splitOutFunction.Count() > 1)
                fieldPath = splitOutFunction.ElementAt(1);
            var split = fieldPath.Split('.');
            var currentType = sourceType;
            list.Add(new KeyValuePair<string, string>(currentType, split.ElementAt(0).Split('|').First()));
            var i = 1;
            if (split.Length > 1)
            {
                foreach (var item in split.Skip(1).Take(split.Length - 1))
                {
                    var fieldName = item.Split('|').First();
                    if (split.ElementAt(i - 1).Contains("|"))
                    {
                        var targetType = split.ElementAt(i - 1).Split('|').Last();
                        list.Add(new KeyValuePair<string, string>(targetType, fieldName));
                        currentType = targetType;
                    }
                    else
                    {
                        var targetType = xrmService.GetLookupTargetEntity(list.ElementAt(i - 1).Value, currentType);
                        list.Add(new KeyValuePair<string, string>(targetType, fieldName));
                        currentType = targetType;
                    }
                    i++;
                }
            }
            return list;
        }

        /// <summary>
        /// Returns a query containing all the fields, and required joins for all the given fields
        /// field examples are "did_contactid.firstname" or "customerid|contact.lastname"
        public static QueryExpression BuildSourceQuery(this XrmService xrmService, string sourceType, IEnumerable<string> fields)
        {
            var query = XrmService.BuildQuery(sourceType, new string[0], null, null);
            foreach (var field in fields)
            {
                xrmService.AddRequiredQueryJoins(query, field);
            }
            return query;
        }

        public static void AddRequiredQueryJoins(this XrmService xrmService, QueryExpression query, string source)
        {
            var typeFieldPaths = xrmService.GetTypeFieldPath(source, query.EntityName);
            var splitOutFunction = source.Split(':');
            if (splitOutFunction.Count() > 1)
                source = splitOutFunction.ElementAt(1);
            var splitTokens = source.Split('.');
            if (typeFieldPaths.Count() == 1)
                query.ColumnSet.AddColumn(typeFieldPaths.First().Value);
            else
            {
                LinkEntity thisLink = null;

                for (var i = 0; i < typeFieldPaths.Count() - 1; i++)
                {
                    var lookupField = typeFieldPaths.ElementAt(i).Value;
                    var path = string.Join(".", splitTokens.Take(i + 1)).Replace("|","_");
                    if (i == 0)
                    {
                        var targetType = typeFieldPaths.ElementAt(i + 1).Key;
                        var matchingLinks = query.LinkEntities.Where(le => le.EntityAlias == path);

                        if (matchingLinks.Any())
                            thisLink = matchingLinks.First();
                        else
                        {
                            thisLink = query.AddLink(targetType, lookupField, xrmService.GetPrimaryKeyField(targetType), JoinOperator.LeftOuter);
                            thisLink.EntityAlias = path;
                            thisLink.Columns = xrmService.CreateColumnSet(new string[0]);
                        }
                    }
                    else
                    {
                        var targetType = xrmService.GetLookupTargetEntity(lookupField, thisLink.LinkToEntityName);
                        var matchingLinks = thisLink.LinkEntities.Where(le => le.EntityAlias == path);
                        if (matchingLinks.Any())
                            thisLink = matchingLinks.First();
                        else
                        {
                            thisLink = thisLink.AddLink(targetType, lookupField, xrmService.GetPrimaryKeyField(targetType), JoinOperator.LeftOuter);
                            thisLink.EntityAlias = path;
                            thisLink.Columns = xrmService.CreateColumnSet(new string[0]);

                        }

                    }
                }
                thisLink.Columns.AddColumn(typeFieldPaths.ElementAt(typeFieldPaths.Count() - 1).Value);
            }
        }


        public static string GetDisplayLabel(this XrmService xrmService, Entity targetObject, string token)
        {
            var fieldPaths = xrmService.GetTypeFieldPath(token, targetObject.LogicalName);
            var thisFieldType = fieldPaths.Last().Key;
            var thisFieldName = fieldPaths.Last().Value;
            var displayString = xrmService.GetFieldLabel(thisFieldName, thisFieldType);
            return displayString;
        }
        public static string GetDisplayString(this XrmService xrmService, Entity targetObject, string token, bool isHtml = false)
        {
            var fieldPaths = xrmService.GetTypeFieldPath(token, targetObject.LogicalName);
            var thisFieldType = fieldPaths.Last().Key;
            var thisFieldName = fieldPaths.Last().Value;
            string func = null;
            var getFieldString = token.Replace("|", "_");
            var splitFunc = getFieldString.Split(':');
            if (splitFunc.Count() > 1)
            {
                func = splitFunc.First();
                getFieldString = splitFunc.ElementAt(1);
            }
            var displayString = xrmService.GetFieldAsDisplayString(thisFieldType, thisFieldName, targetObject.GetFieldValue(getFieldString));
            return displayString;
        }
    }
}
