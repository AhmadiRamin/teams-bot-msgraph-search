using AdaptiveCards;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using MSGraphSearchSample.Constants;
using MSGraphSearchSample.Constants.MessagingExtension;
using MSGraphSearchSample.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace MSGraphSearchSample.Bots
{
    public partial class Bot<T> : TeamsActivityHandler
        where T : Dialog
    {
        // Handle queries
        // More info: https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/search-commands/define-search-command
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
        {
            // get query object
            query = query ?? throw new ArgumentNullException(nameof(query));
            // query string value
            // Multi-parameter is only supported for command type set to action. For query we only support 1 parameter currently.
            var text = query.Parameters.FirstOrDefault()?.Value as string ?? string.Empty;
            var parameterName = query.Parameters.FirstOrDefault()?.Name;

            // Setup Single Sign-On (obtaining access token)
            var userConfigSettings = await UserConfigProperty.GetAsync(turnContext, () => string.Empty);
            var tokenResponse = await GetTokenResponse(query.State, turnContext, cancellationToken);
            if (tokenResponse != null)
                return tokenResponse;

            // Return some items on initial run (if it's enabled in manifest)
            if (parameterName == "initialRun")
                return await GetPreviewItems();

            // Use your service (Graph, SharePoint, etc) to generate the MessagingExtensionAttachment
            switch (query.CommandId)
            {
                case CommandIds.SearchByName:
                    return await GetByName(text.ToLower());
                default:
                    return new MessagingExtensionResponse();
            }
        }

        // Method to handle link unfurling
        // more info: https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/link-unfurling?tabs=dotnet
        protected override async Task<MessagingExtensionResponse> OnTeamsAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
        {
            var url = query.Url.ToLower();

            // SSO integration, you can remove this if you don't need users to sign in
            var tokenResponse = await GetTokenResponse(query.State, turnContext, cancellationToken);
            if (tokenResponse != null)
                return tokenResponse;

            // In this sample we pull id from the query string, you can just use the url
            var queryStrings = HttpUtility.ParseQueryString(url.Substring(url.IndexOf('?')));
            var itemId = queryStrings.Get("id");

            if (itemId != null)
            {
                // Call your service here to get your results, and bind it to an adaptive card, or any supported cards

                var card = _fileService.GetCard("UnfurlingCard");
                var cardAttachment = new Attachment() { ContentType = AdaptiveCard.ContentType, Content = JsonConvert.DeserializeObject(card) };

                // Create attachments object
                var attachments = new List<MessagingExtensionAttachment>(){ new MessagingExtensionAttachment
                {
                    ContentType = cardAttachment.ContentType,
                    Content = cardAttachment.Content,
                    Preview = new ThumbnailCard
                    {
                        Title = "Title",
                        Subtitle = "Sub Title"
                    }.ToAttachment()
                } };

                // To boost the performance, you can add Cache info to the following object
                // more info: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.schema.teams.cacheinfo?view=botbuilder-dotnet-stable
                return new MessagingExtensionResponse
                {
                    // more info: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.schema.teams.messagingextensionresult?view=botbuilder-dotnet-stable#properties
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Type = "result", // possible values: result, auth, config
                        AttachmentLayout = "list", // possible values: list, grid
                        Attachments = attachments
                    }
                };
            }

            // Returns empty response
            return new MessagingExtensionResponse();

        }

        // The Preview card's Tap should have a Value property assigned, this will be returned to the bot in this event
        protected override Task<MessagingExtensionResponse> OnTeamsMessagingExtensionSelectItemAsync(ITurnContext<IInvokeActivity> turnContext, JObject query, CancellationToken cancellationToken)
        {
            // IMPORTANT: is not triggered in mobile teams application

            // We take every row of the results and wrap them in cards wrapped in MessagingExtensionAttachment objects.
            // The Preview is optional, if it includes a Tap, that will trigger the OnTeamsMessagingExtensionSelectItemAsync event back on this bot.

            return null;
        }

        // Methods to handle messaging extension responses
        protected async Task<MessagingExtensionResponse> GetByName(string text)
        {
            var attachments = new List<MessagingExtensionAttachment>();
            // Get the attachment items from your Graph or any other services and add them to the attachments collection

            return new MessagingExtensionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "result",
                    AttachmentLayout = "list",
                    Attachments = attachments
                }
            };
        }
        // Return initial run results
        protected async Task<MessagingExtensionResponse> GetPreviewItems()
        {
            var attachments = new List<MessagingExtensionAttachment>();
            // Get the attachment items from your Graph or any other services and add them to the attachments collection

            return new MessagingExtensionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "result",
                    AttachmentLayout = "list",
                    Attachments = attachments
                }
            };
        }

    }
}
