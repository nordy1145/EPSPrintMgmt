# Change Log
All notable changes to this project will be documented in this file.

## [Unreleased]
### Added
- Looking into Enterprise Printer Management.
- Additional Print Driver management along with label printers.

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