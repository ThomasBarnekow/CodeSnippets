using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace CodeSnippets.Tests
{
    public class CsvTests
    {
        private const string Xml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<soap:Envelope xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">
<soap:Body>
<GetAccountsResponse xmlns=""http://Weiland.Ba.retrieval/"">
<GetAccountsResult>
<AccountRecords>3071</AccountRecords>
<Accounts>
<Account>
<AccountId>6</AccountId>
<AccountNumber>000091016138</AccountNumber>
<AccountTitle>3946231 Canada Inc.</AccountTitle>
<BankId>70</BankId>
<BankName>Royal Bank of Canada</BankName>
<BankRtn>ROYCCAT2</BankRtn>
<BranchId xsi:nil=""true""/>
<CompanyId>86</CompanyId>
<CurrencyName>Canada, Dollar</CurrencyName>
<CostCenter>4603</CostCenter>
<ActiveDate>1900-01-01T00:00:00</ActiveDate>
</Account>
</Accounts>
<AsOfDate>2019-10-31 06:04:36.059 -04:00</AsOfDate>
<BankRecords>0</BankRecords>
<CompanyRecords>0</CompanyRecords>
<ReturnCode>0</ReturnCode>
<Version>1.03</Version>
</GetAccountsResult>
</GetAccountsResponse>
</soap:Body>
</soap:Envelope>";

        private static readonly XNamespace Ns = "http://Weiland.Ba.retrieval/";

        private static readonly XName Accounts = Ns + "Accounts";
        private static readonly XName Account = Ns + "Account";
        private static readonly XName AccountId = Ns + "AccountId";
        private static readonly XName AccountTitle = Ns + "AccountTitle";
        private static readonly XName BankId = Ns + "BankId";
        private static readonly XName BankName = Ns + "BankName";
        private static readonly XName BankRtn = Ns + "BankRtn";
        private static readonly XName BranchId = Ns + "BranchId";
        private static readonly XName CompanyId = Ns + "CompanyId";
        private static readonly XName CurrencyName = Ns + "CurrencyName";
        private static readonly XName CostCenter = Ns + "CostCenter";
        private static readonly XName ActiveDate = Ns + "ActiveDate";

        [Fact]
        public void CanTransformXmlToCsv()
        {
            XElement envelope = XElement.Parse(Xml);
            XElement accounts = envelope.Descendants(Accounts).First();
            WriteAccountsCsv(accounts);
        }

        private static void WriteAccountsCsv(XElement accounts)
        {
            string separator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            string header = CreateCsvHeader(separator);
            string body = CreateCsvBody(accounts, separator);
            string csv = header + body;
            File.WriteAllText("Accounts.csv", csv);
        }

        private static string CreateCsvHeader(string separator)
        {
            var sb = new StringBuilder();

            sb.Append(AccountId.LocalName).Append(separator);
            sb.Append(AccountTitle.LocalName).Append(separator);
            sb.Append(BankId.LocalName).Append(separator);
            sb.Append(BankName.LocalName).Append(separator);
            sb.Append(BankRtn.LocalName).Append(separator);
            sb.Append(BranchId.LocalName).Append(separator);
            sb.Append(CompanyId.LocalName).Append(separator);
            sb.Append(CurrencyName.LocalName).Append(separator);
            sb.Append(CostCenter.LocalName).Append(separator);
            sb.AppendLine(ActiveDate.LocalName);

            return sb.ToString();
        }

        private static string CreateCsvBody(XElement accounts, string separator)
        {
            return accounts.Elements(Account)
                .Aggregate(
                    new StringBuilder(),
                    (sb, account) => sb.AppendLine(CreateCsvLineItem(account, separator)),
                    sb => sb.ToString());
        }

        private static string CreateCsvLineItem(XElement account, string separator)
        {
            var sb = new StringBuilder();

            AppendUnquoted(sb, account.Element(AccountId), separator);
            AppendQuoted(sb, account.Element(AccountTitle), separator);
            AppendUnquoted(sb, account.Element(BankId), separator);
            AppendQuoted(sb, account.Element(BankName), separator);
            AppendQuoted(sb, account.Element(BankRtn), separator);
            AppendUnquoted(sb, account.Element(BranchId), separator);
            AppendQuoted(sb, account.Element(CompanyId), separator);
            AppendQuoted(sb, account.Element(CurrencyName), separator);
            AppendUnquoted(sb, account.Element(CostCenter), separator);
            AppendQuoted(sb, account.Element(ActiveDate), null);

            return sb.ToString();
        }

        private static void AppendUnquoted(StringBuilder sb, XElement element, string separator)
        {
            sb.Append($"{element?.Value}");
            if (separator != null)
            {
                sb.Append(separator);
            }
        }

        private static void AppendQuoted(StringBuilder sb, XElement element, string separator)
        {
            sb.Append($"\"{element?.Value}\"");
            if (separator != null)
            {
                sb.Append(separator);
            }
        }
    }
}
