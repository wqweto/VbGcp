## VbGcp
VB6 Google Cloud Print proxy

VbGcp is a set of classes that support printing documents (PDF, TXT, JPG, GIF, PNG, DOC, XLS and more) through [Google Cloud Print](http://www.google.com/cloudprint/learn/) devices from VB6 projects.

GCP is available as a REST service on `www.google.com/cloudprint` address. There are two parts of the printing process: submitting jobs by client applications (this proxy) and receiving jobs by registered printers (firmware or software proxies).

GCP documentation used to develop VbGcp is available on [Submitting Print Jobs](https://developers.google.com/cloud-print/docs/sendJobs) on Google Developers. You can check out the [Google Cloud Print API Simulation](http://www.google.com/cloudprint/simulate.html) page too.

The easiest way to test GCP service is to [enable the Google Cloud Print connector in Google Chrome](http://support.google.com/cloudprint/bin/answer.py?&answer=1686197) which will register all your local printers as cloud printers. This works both with Windows printers and with CUPS printers on Linux.

### Sample included

The [sample application](https://github.com/wqweto/VbGcp/raw/master/Sample/GCPSample.exe) exercises the classes by rendering a prototype of GCP settings dialog (OAuth2 login and consent), print dialog (select printer, copies, collation, etc.) and printer properties dialog (more printer settings). 

![Login and Print dialog](https://github.com/wqweto/VbGcp/raw/master/Doc/ss_gcp_1.png)
![Printer Setup dialog](https://github.com/wqweto/VbGcp/raw/master/Doc/ss_gcp_2.png)

The sample can submit a file to GCP service and tracks its print job progress as GCP service processes the file. This is done by continually listing current printer jobs by status. All of the operations are implemented in async mode.

### Using the classes

`cGcpService` is the main class that supports calling methods on the GCP service and parsing JSON response. When `ASYNC_SUPPORT` is set to `False` the class can be used as a standalone class. Class methods map to service calls as follows:

 - `PrintDocument` calls [/submit](https://developers.google.com/cloud-print/docs/appInterfaces#submit) method
 - `GetJobs` calls [/jobs](https://developers.google.com/cloud-print/docs/appInterfaces#jobs) method
 - `DeleteJob` calls [/deletejob](https://developers.google.com/cloud-print/docs/appInterfaces#deletejob) method
 - `GetPrinterInfo` calls [/printer](https://developers.google.com/cloud-print/docs/appInterfaces#printer) method
 - `GetPrinters` calls [/search](https://developers.google.com/cloud-print/docs/appInterfaces#search) method

For asynchronous service requests leave `ASYNC_SUPPORT` set to `True` and include in your project `cGcpCallback` class too.

`cGcpPrinterCaps` is a helper class for easier setting of job capabilities -- number of copies, paper size, page orientation, etc. Print job settings are submitted to GCP service in JSON format and can be constructed manually. This wrapper class just parses printer capabilities as returned by the GCP service (both XPS and PPD formats) and exposes these in a consistent way as settable properties. (This is a work in progress from Google, results are not very consistent at present.)

Finally `cGcpOAuth` is a helper class for OAuth2 user authentication in [Installed Application](https://developers.google.com/accounts/docs/OAuth2#installed) mode. It uses Google Accounts to validate user login and to acquire user consent on accessing GCP service. Then it retrieves OAuth2 `refresh_token` from `accounts.google.com/o/oauth2/token` service which can be stored and later used to retrieve OAuth `access_token` for GCP service without showing another login screen.

## [API](https://github.com/wqweto/VbGcp/blob/master/Doc/API.md)
