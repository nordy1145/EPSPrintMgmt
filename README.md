# EPSPrintMgmt

UPDATE:
The lastest update pushed 9/13/2016 requires an updated web.config file.  Please copy your appSettings section in your web.config and replace the section in the web.config.template file.  Then rename it to just web.config.
If you have issues with the drop down for looking up printers, refresh your browser cache by pressing ctrl+F5.


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
