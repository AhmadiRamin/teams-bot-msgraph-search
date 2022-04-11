using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MSGraphSearchSample.Interfaces
{
    public interface IGraphHelper
    {
        GraphServiceClient GetApplicationServiceClient();
        GraphServiceClient GetDelegatedServiceClient(string _token);
        Task<string> GetOnBehalfOfAccessToken(string _token, string resourceUri);
    }
}
