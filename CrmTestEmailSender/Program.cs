using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.ServiceModel.Description;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace CrmTestEmailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var organizationServiceUrl = ConfigurationManager.AppSettings["CrmOrganizationService"];
                var userName = ConfigurationManager.AppSettings["UserName"];
                var password = ConfigurationManager.AppSettings["Password"];
                var fromId = ConfigurationManager.AppSettings["EmailFromCrmUserId"];
                var toIds = ConfigurationManager.AppSettings["EmailToCrmUsersId"];

                var uri = new Uri(organizationServiceUrl);
                var clientCredentials = new ClientCredentials();
                if (string.IsNullOrWhiteSpace(userName))
                {
                    clientCredentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;
                }
                else
                {
                    clientCredentials.Windows.ClientCredential.UserName = userName;
                    clientCredentials.Windows.ClientCredential.Password = password;
                }
                var service = new OrganizationServiceProxy(uri, null, clientCredentials, null);

                var toIdsArray = toIds.Split(';');

                var email = new Entity("email");
                email["subject"] = "testing " + DateTime.Now;
                email["description"] = "test message" + DateTime.Now;

                var toList = new List<Entity>();
                foreach (var toId in toIdsArray)
                {
                    var to = new Entity("activityparty");
                    to["partyid"] = new EntityReference("systemuser", new Guid(toId));
                    toList.Add(to);
                }
                email["to"] = toList.ToArray();

                var from = new Entity("activityparty");
                from["partyid"] = new EntityReference("systemuser", new Guid(fromId));
                email["from"] = new Entity[] { from };

                var createdEmail = service.Create(email);

                var request = new SendEmailRequest();
                request.EmailId = createdEmail;
                request.TrackingToken = string.Empty;
                request.IssueSend = true;
                service.Execute(request);

                Console.WriteLine("Success. Email Id=" + createdEmail);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            Console.ReadKey();
        }
    }
}