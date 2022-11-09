using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class ReportServerCredentials : IReportServerCredentials
{
    public System.Security.Principal.WindowsIdentity ImpersonationUser
    {
        get
        {
            return null;
        }
    }

    public System.Net.ICredentials NetworkCredentials
    {
        get
        {
            return new System.Net.NetworkCredential(_username, _password, _domain);
        }
    }

    string _username;
    string _password;
    string _domain;

    public ReportServerCredentials()
    {
        _username = "Administrator";
        _password = "Jakarta01";
        _domain = null;
    }

    public ReportServerCredentials(string userName, string password, string domain)
    {
        _username = userName;
        _password = password;
        _domain = domain;
    }

    public bool GetFormsCredentials(out System.Net.Cookie authCookie,
            out string user, out string password, out string authority)
    {
        authCookie = null;
        user = password = authority = null;
        return false;
    }
}
