using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JosephM.Core.Attributes;
using JosephM.Core.Service;
using JosephM.Xrm.WorkflowScheduler;
using Microsoft.Xrm.Sdk.Client;

namespace JosephM.Xrm.Settings.Test
{
    public class TestXrmConfiguration : XrmConfiguration, IValidatableObject
    {
        [DisplayOrder(20)]
        [RequiredProperty]
        public override AuthenticationProviderType AuthenticationProviderType { get; set; }
        [DisplayOrder(30)]
        [RequiredProperty]
        [GridWidth(400)]
        public override string DiscoveryServiceAddress { get; set; }
        [DisplayOrder(40)]
        [RequiredProperty]
        public override string OrganizationUniqueName { get; set; }
        [DisplayOrder(50)]
        [RequiredProperty]
        [GridWidth(100)]
        [PropertyInContextByPropertyValues("AuthenticationProviderType",
            new object[]
            {
                AuthenticationProviderType.ActiveDirectory, AuthenticationProviderType.Federation
                , AuthenticationProviderType.OnlineFederation
            })]
        public override string Domain { get; set; }
        [DisplayOrder(60)]
        [RequiredProperty]
        [GridWidth(100)]
        [PropertyInContextByPropertyValues("AuthenticationProviderType",
            new object[]
            {
                AuthenticationProviderType.ActiveDirectory, AuthenticationProviderType.Federation
                , AuthenticationProviderType.OnlineFederation,
                AuthenticationProviderType.LiveId
            })]
        public override string Username { get; set; }
        [DisplayOrder(70)]
        [RequiredProperty]
        [GridWidth(100)]
        [PropertyInContextByPropertyValues("AuthenticationProviderType",
            new object[]
            {
                AuthenticationProviderType.ActiveDirectory, AuthenticationProviderType.Federation
                , AuthenticationProviderType.OnlineFederation,
                AuthenticationProviderType.LiveId
            })]
        public override string Password { get; set; }

        public IsValidResponse Validate()
        {
            var service = new XrmService(this);
            var response = service.VerifyConnection();
            var actualResponse = new IsValidResponse();
            foreach (var item in response.InvalidReasons)
                actualResponse.AddInvalidReason(item);
            return actualResponse;
        }
    }
}
