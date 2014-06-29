using Aspose.Email.Mail;
using Aspose.Words;
using Aspose.Words.Loading;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose {
    class Program {
        private const string inFile = @"c:\devel\data\mhtml\simple_link2.mhtml";
        private const string outFile = @"c:\devel\data\mhtml\simple_link2.pdf";

        static void Main(string[] args) {
            Convert();
        }

        private static void Convert() {
            var options = new LoadOptions();
            options.ResourceLoadingCallback = new MhtmlResourceLoadingCallback(inFile);

            var document = new Document(inFile, options);
            document.Save(outFile, Aspose.Words.SaveFormat.Pdf);
        }
    }


    public class MhtmlResourceLoadingCallback : IResourceLoadingCallback
    {
         private readonly HashSet<string> resourceNames;
 
        public MhtmlResourceLoadingCallback(string fullFileName) {
             var message = MailMessage.Load(fullFileName);
            resourceNames = new HashSet<string>(message.LinkedResources.Select(resource => resource.ContentId));
        }
 
        public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
        {
            var resourceName = args.OriginalUri.StartsWith("cid:") ? args.OriginalUri.Substring(4) : args.OriginalUri;
            return resourceNames.Contains(resourceName) ? ResourceLoadingAction.Default : ResourceLoadingAction.Skip;
        }
     }

    /*
     * public class MhtmlResourceLoadingCallback : IResourceLoadingCallback {
        private readonly HashSet<string> resourceNames;

        public MhtmlResourceLoadingCallback(string fullFileName) {
            var message = MailMessage.Load(fullFileName);
            resourceNames = new HashSet<string>();
            foreach (var resource in message.LinkedResources) {
                resourceNames.Add(resource.ContentId);
                if (resource.ContentLink != null && !resourceNames.Contains(resource.ContentLink.AbsoluteUri))
                    resourceNames.Add(resource.ContentLink.AbsoluteUri);
            }
        }

        public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args) {
            var resourceName = args.Uri.StartsWith("cid:") ? args.Uri.Substring(4) : args.Uri;
            return resourceNames.Contains(resourceName) ? ResourceLoadingAction.Default : ResourceLoadingAction.Skip;
        }
    }
     * */

}
