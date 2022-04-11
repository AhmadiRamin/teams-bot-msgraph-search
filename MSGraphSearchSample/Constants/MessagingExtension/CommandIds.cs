using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MSGraphSearchSample.Constants.MessagingExtension
{
    public static class CommandIds
    {
        // Search commands
        public const string SearchByName = "SearchByName";
        public const string SearchByType = "SearchByType";
        public const string SearchByDate = "SearchByDate";

        // Action commands
        public const string CreateCard = "CreateCard";

        // SSO Commands
        public const string ShowProfile = "SHOWPROFILE";
        public const string SignOutCommand = "SIGNOUTCOMMAND";


    }
}
