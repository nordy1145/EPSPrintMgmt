# EPSPrintMgmt

##UPDATE:
-v1.0.9.8 If you want to clone a EPS Print queue, you need the following update installed on the web server in order to do this
https://support.microsoft.com/en-us/help/2954953/some-apis-do-not-work-when-they-are-called-in-services-in-windows

-v1.0.9.5 requires the new appsettings.config file to be populated with your custom config.  Please see the changelog file for more information.


-The lastest update pushed 9/13/2016 requires an updated web.config file.  Please copy your appSettings section in your web.config and replace the section in the web.config.template file.  Then rename it to just web.config.
###If you have issues with the drop down for looking up printers, refresh your browser cache by pressing ctrl+F5.


Download the "Compiled (date)".zip found in the Downloads folder if you would like the precompiled version. Replace all content except the web.config file.

Some requirements to run this project.
Requires PowerShell 3+ installed on all servers, including web and print servers.
WinRM must be enabled on EPS and Print Servers.
On the web server, IIS is required with the following features:
  
  Security - URL Authorization
  
  Security - Windows Authentication
  
  Application Development - .NET Extensibility 4.5
  
  Application Development - ASP.NET 4.5
  
Copy the extracted compiled version to the wwwroot directory or a virtual application on the IIS server.
Either create a new App Pool or use the DefaultAppPool and change the Identity to a service account that has admin access to the EPS/Print Servers.
  App Pools -> Advanced Settings -> Identity -> change to service account
Under the Default Web Site or custom Virtual Application go to Basic Settings and change the Connect As option to the service account defined above.  May have to change to the custom Application Pool created above.
  Under the IIS section open Authenication Disable Anonymous Authentication and Enable Windows Authentication.
  
Either at the Top IIS server level or the Web Site level, go to the ASP.NET section -> .NET Authorization Rules.  Remove the Allow Any and add an Allow Rule.   Add the AD Group you want to give access to the actual website.

Can bind a certificate to port 443 if you'd like.

Review the EPSPrintMgmt/ChangeLog.MD file for changes that have happened in each version.

[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=JYESRLRN7N2WC&lc=US&item_name=EPS%20Print%20Management&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donateCC_LG%2egif%3aNonHosted)
