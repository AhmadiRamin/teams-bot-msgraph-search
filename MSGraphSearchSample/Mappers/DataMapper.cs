
using Microsoft.Graph;
using MSGraphSearchSample.Constants.Search;
using MSGraphSearchSample.Helpers;
using MSGraphSearchSample.Models.Search;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Event = MSGraphSearchSample.Models.Search.Event;

namespace MSGraphSearchSample.Mappers
{
    public static class DataMapper
    {
        public static List<Event> GetEvents(List<SearchHit> hits)
        {
            var items = new List<Event>();
            CultureInfo culture = CultureInfo.InvariantCulture;
            foreach (var hit in hits)
            {
                var graphEvent = hit.Resource as Microsoft.Graph.Event;
                var item = new Models.Search.Event()
                {
                    Subject = graphEvent.Subject,
                    End = DateTime.Parse(graphEvent.End.DateTime).ToString(),
                    Start = DateTime.Parse(graphEvent.Start.DateTime).ToString(),
                    IsAllDay = graphEvent.IsAllDay==false ? "No" : "Yes"
                };
                items.Add(item);
            }
            return items;
        }
        public static List<Models.Search.DriveItem> GetFiles(List<SearchHit> hits)
        {
            var items = new List<Models.Search.DriveItem>();
            foreach (var hit in hits)
            {
                var graphDriveItem = hit.Resource as Microsoft.Graph.DriveItem;
                var fileExtension = Regex.Match(graphDriveItem.Name, "[^.]+$").Value;
                var icon = FileExtensionIcons.getIconByFileType(fileExtension);
                var item = new Models.Search.DriveItem()
                {
                    Name = graphDriveItem.Name,
                    CreatedDateTime = graphDriveItem.CreatedDateTime.Value.DateTime.ToString(),
                    CreatedBy= graphDriveItem.CreatedBy.User.DisplayName,
                    Size= graphDriveItem.Size.Value.ToString(),
                    WebUrl = graphDriveItem.WebUrl,
                    Icon = icon
                };
                items.Add(item);
            }
            return items;
        }

        public static List<Models.Search.Message> GetMessages(List<SearchHit> hits)
        {
            var items = new List<Models.Search.Message>();
            foreach (var hit in hits)
            {
                var graphMessage = hit.Resource as Microsoft.Graph.Message;
                var item = new Models.Search.Message()
                {
                    Subject = graphMessage.Subject,
                    From = graphMessage.From.EmailAddress.Name,
                    CreatedDatedTime = graphMessage.CreatedDateTime.Value.DateTime.ToString(),
                    HasAttachments = graphMessage.HasAttachments == true ? "Yes" : "No",
                    WebLink = graphMessage.WebLink
                };
                items.Add(item);
            }
            return items;
        }

        public static List<Models.Search.ListItem> GetListItems(List<SearchHit> hits)
        {
            var items = new List<Models.Search.ListItem>();
            foreach (var hit in hits)
            {
                var item = new Models.Search.ListItem()
                {
                    Title = EntityHelper.GetValue(hit.Resource, "title"),
                    URL = EntityHelper.GetValue(hit.Resource, "url"),
                    Created = EntityHelper.GetValue(hit.Resource, "created"),
                    CreatedBy = EntityHelper.GetValue(hit.Resource, "createdby"),
                    LastModifiedTime = EntityHelper.GetValue(hit.Resource, "lastmodifiedtime")
                };
                items.Add(item);
            }
            return items;
        }
    }
}
