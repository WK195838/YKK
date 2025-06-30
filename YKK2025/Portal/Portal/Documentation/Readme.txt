DotNetNuke

For more details please see the DotNetNuke Installation Guide (in the Documentation\Public folder)

Clean Installation

- .NET Framework 1.1 must be installed ( 1.0 no longer supported )
- unzip package into C:\DotNetNuke
- create Virtual Directory in IIS called DotNetNuke which points to the directory where the DotNetNuke.vbproj file exists
- make sure you have default.aspx specified as a Default Document for your Virtual Directory
- rename release.config -> web.config
- SQL Server 2000 or MSDE 2000 ( MSDE can be acquired from http://www.asp.net/Tools/redir.aspx?path=msde )
  - manually create SQL Server database named "DotNetNuke" ( using Enterprise Manager or your tool of choice )
  - set the connection string in web.config ( ie. <add key="SiteSqlServer" value="Server=(local);Database=DotNetNuke;uid=;pwd=;" /> )
- enter username/password for the SuperUser and Administrator for the default portal in dotnetnuke.install template file.  These are currently set to host/host and admin/admin respectively
- browse to localhost/DotNetNuke in your web browser
- the application will automatically execute the necessary database scripts

Application Upgrades

- make sure you always backup your files/database before upgrading to a new version
- unzip the code over top of your existing application
- BACKUP web.config (web.backup.resources)
- follow procedure in Installation Guide to ensure web.config is set up right
- browse to localhost/DotNetNuke in your web browser
- the application will automatically execute the necessary database scripts

Security ( details in /403-3.htm file )

If using Windows 2000 - IIS5
- the {Server}/ASPNET user account must have Read, Write, and Change Control of the root application directory ( this allows the application to create files/folders ) 
If using Windows 2003 - IIS6
- the {Server}/NetworkService user account must have Read, Write, and Change Control of the root application directory ( this allows the application to create files/folders ) 
