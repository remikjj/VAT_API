using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Vat_API4
{
    public class AuthorizedClerk
    {
        [JsonProperty("companyName")]
        public object CompanyName { get; set; }

        [JsonProperty("firstName")]
        public string FirstName { get; set; }

        [JsonProperty("lastName")]
        public string LastName { get; set; }

        [JsonProperty("nip")]
        public object Nip { get; set; }

        [JsonProperty("pesel")]
        public object Pesel { get; set; }
    }

    public class Subject
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("nip")]
        public string Nip { get; set; }

        [JsonProperty("statusVat")]
        public string StatusVat { get; set; }

        [JsonProperty("regon")]
        public string Regon { get; set; }

        [JsonProperty("pesel")]
        public object Pesel { get; set; }

        [JsonProperty("krs")]
        public string Krs { get; set; }

        [JsonProperty("residenceAddress")]
        public object ResidenceAddress { get; set; }

        [JsonProperty("workingAddress")]
        public string WorkingAddress { get; set; }

        [JsonProperty("representatives")]
        public List<object> Representatives { get; set; }

        [JsonProperty("authorizedClerks")]
        public List<AuthorizedClerk> AuthorizedClerks { get; set; }

        [JsonProperty("partners")]
        public List<object> Partners { get; set; }

        [JsonProperty("registrationLegalDate")]
        public string RegistrationLegalDate { get; set; }

        [JsonProperty("registrationDenialBasis")]
        public object RegistrationDenialBasis { get; set; }

        [JsonProperty("registrationDenialDate")]
        public object RegistrationDenialDate { get; set; }

        [JsonProperty("restorationBasis")]
        public object RestorationBasis { get; set; }

        [JsonProperty("restorationDate")]
        public object RestorationDate { get; set; }

        [JsonProperty("removalBasis")]
        public string RemovalBasis { get; set; }

        [JsonProperty("removalDate")]
        public string RemovalDate { get; set; }

        [JsonProperty("accountNumbers")]
        public List<object> AccountNumbers { get; set; }

        [JsonProperty("hasVirtualAccounts")]
        public bool HasVirtualAccounts { get; set; }
    }

    public class Result
    {
        [JsonProperty("subject")]
        public Subject Subject { get; set; }

        [JsonProperty("requestId")]
        public string RequestId { get; set; }

        [JsonProperty("requestDateTime")]
        public string RequestDateTime { get; set; }
    }

    public class Root
    {
        [JsonProperty("result")]
        public Result Result { get; set; }
    }




}
