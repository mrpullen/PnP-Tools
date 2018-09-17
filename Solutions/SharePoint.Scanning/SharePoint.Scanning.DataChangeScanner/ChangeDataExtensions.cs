using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.Scanning.DataChangeScanner
{
    public static class ChangeDataExtensions
    {
        public static List<Change> GetAllSiteChanges(this ClientContext context, ChangeQuery query)
        {
            List<Change> allChanges = new List<Change>();
            var changes = context.Site.GetChanges(query);
            context.Load(changes);
            context.ExecuteQueryRetry();

            string lastToken = null;
            while (changes.Count > 0) {
                foreach (var change in changes)
                {
                    allChanges.Add(change);
                    lastToken = change.ChangeToken.StringValue;
                }

                query.ChangeTokenStart.StringValue = lastToken;
                changes = context.Site.GetChanges(query);
                context.Load(changes);
                context.ExecuteQueryRetry();

            }

            return allChanges;
        }
        public static bool GreaterThan(this ChangeToken mainToken, ChangeToken compareToken)
        {
            if(compareToken == null)
            {
                return true;
            }

            var mainValue = mainToken.StringValue.Split(';')[4];
            long mainLong = Convert.ToInt64(Convert.ToDecimal(mainValue));

            var compValue = compareToken.StringValue.Split(';')[4];
            long compLong = Convert.ToInt64(Convert.ToDecimal(compValue));

            return (mainLong > compLong);
                
                
        }
    }
}
