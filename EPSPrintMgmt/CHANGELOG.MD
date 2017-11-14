# Change Log
All notable changes to this project will be documented in this file.

## [Unreleased]
### Added
- Additional Print Driver management along with label printers.
- Background jobs for adding/remove printers

### Changed
- Change the search option to choose either a EPS or Enterprise option, versus having to choose an individual server.

## [1.0.9.10] - 2017-11-13
### Added
- Added ability to copy print driver and device settings from a source Print Queue
### Changed
- Ability to manage print trays as this should be handled by cloning a print queue


## [1.0.9.9] - 2017-10-30
### Added
- Added support to add EPS and Enterprise print queues at the same time
- Added additional appsettings.config items
### Changed
- A potential broken Roslyn package that prevented the application from running properly after an upgrade.

## [1.0.9.8] - 2017-10-17
### Added
- Added the ability to disable bidirectional support when adding Enterprise Printers
- Added the ability to clone a EPS print queue when creating a new one.
- Added seperate email notification for Adding Enterprise Print Queues.
### Fixed
- Wording in appsettings.config file to better indicate the EditEnterprisePrinters key.

## [1.0.9.7] - 2016-10-05
### Fixed
- Fixed issue with Enterprise Printer details would not edit/delete correct printer.  Moved buttons to the row of the print server.

## [1.0.9.6] - 2016-10-05
### Added
- Email notification for purging print jobs from an entire print queue.
- Added additional debugging information for Print job Purging. 

### Changed
- Changed Print Server page to include the Server IP address and count of printers, if it's under a few thousand.

## [1.0.9.5] - 2016-10-05
### Added
- Added config options for additional security to delete, edit and purge print queues.  Includes new fields within the appsettings section.
- Added the initial build for adding/editing/deleting Enterprise Printers.  Will automatically Share the printer with the printer name.
- Added section in appsettings config file for Enterprise Print Drivers.
- Added a new tab to create Enterprise Printer if config is set to allow it.

### Changed
- Moved the appsettings section of web.config to it's own file.
- Modified Details view to handle both EPS and Enterprise Print Servers
- Modified the search to search all Enterprise Print Servers when an EPS server is not selected.

## [1.0.9.4] - 2016-09-29
### Added
- Added Delete Printer option from Printer details section.  Does not delete the printer port.
- Added Jquery UI package for dialog box integration.  Currently only being used for Deleting a Printer.
- Added new config option in Web.Config.  <add key="AllowEPSPrinterDeletion" value="false" />

### Changed
- The delete method for removing printers.  Now uses the PrintServer class instead of WMI.
Has improved performance for removing/changing printers.

## [1.0.9.3] - 2016-09-28
### Added
- Started initial logging to local log file.  Includes new section of web.config file for logging settings.
Using log4net library for logging
- Change log
- HTML Header to force IE to use the Edge engine.
- Created a support class to be used across all controllers.

### Changed
- The details of a printer no longer try to retrieve tray if the Driver contains "Zdesigner"
or if the UsePrintTrays config option is set to false in the web.config file.
- Changed some methods to use the new support class.

### Fixed
- Print Server name not showing up while editing a printer in the web page response or email notification.

## [1.0.9.2] - 2016-09-19
### Added
- Print test page option from details section of printer.
- Added version info on the bottom of the page.
- Ability to delete multiple print jobs from the Print Jobs tab.
- Added PowerShell script to automatically clear out old print jobs.

## [1.0.9.1] - 2016-09-13
### Added
- Added a search for a printer option.  This returns the details view of the printer.

### Changed
- Updated Jquery and other core components of the application.

## [1.0.9.0] - 2016-09-08
### Changed
- Changed the list of printers from print server to not show the number of jobs in tray.  This decreased load
