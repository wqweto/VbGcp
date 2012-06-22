## VbGcp
VB6 Google Cloud Print proxy

VbGcp is a set of classes that support printing documents (PDF, TXT, JPG, GIF, PNG, DOC, XLS and more) through [Google Cloud Print](http://www.google.com/cloudprint/learn/) devices from VB6 projects.

GCP is available as a REST service on `www.google.com/cloudprint` address. There are two parts of the printing process: submitting jobs by client applications (this proxy) and receiving jobs by registered printers (firmware or software proxies).

GCP documentation used to develop VbGcp is available on [Submitting Print Jobs](https://developers.google.com/cloud-print/docs/sendJobs) on Google Developers.

The easiest way to test GCP service is to [enable the Google Cloud Print connector in Google Chrome](http://support.google.com/cloudprint/bin/answer.py?&answer=1686197) which will register all your local printers as cloud printers. This works both with Windows printers and with CUPS printers on Linux.

### Sample included

The [sample application](https://github.com/wqweto/VbGcp/raw/master/Sample/GCPSample.exe) exercises the classes by rendering a prototype of GCP settings dialog (OAuth2 login and consent), print dialog (select printer, copies, collation, etc.) and printer properties dialog (more printer settings). 

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

## API
(Work in progress)

### `cGcpService` class

GCP service wrapper class. Calls REST service methods and parses returned JSON result.

#### `Init(sHost As String, sCredentials As String, ByVal eType As GcpCredentialsTypeEnum) As Boolean`

Used to initialize service location and credentials used to access GCP service. The `sHost` parameter is either `https://www.google.com` for SSL or `http://www.google.com` for unencrypted access to the service. Parameter `sCredentials` contains the credentials used to access GCP service. Its format depends on `eType` parameter. 

 - `gcpCrtGoogleLogin`: `sCredentials` format is `google_account:password`
 - `gcpCrtOAuthRefreshToken`: `sCredentials` format is `refresh_token:client_id:client_secret`
 
The `refresh_token` is retrieved from OAuth2 service on user login/consent. The `client_id` and `client_secret` should be hard-coded in the application. These are acquired by registering your application with Google's OAuth2 service through their [APIs Console](https://code.google.com/apis/console#access).

#### `Property AsyncOperations As Boolean`

Determines if service calls (see methods table above) block or return immediately. In synchronous mode service methods return parsed JSON object. In async mode service methods return `cGcpCallback` object that raises its `Complete` event on request completion.

#### `PrintDocument(sPrinterId As String, sFile As String, [Title As String], [ContentType As String], [Capabilities As String]) As Object`

Prints a file to a particular cloud printer.

#### `GetPrinters([Pattern As String]) As Object`

Retrieves list of available printers. GCP service orders printers from most recently used to least used.

#### `GetPrinterInfo(sPrinterId As String) As Object`

Retrieves printer info including printer capabilities.

#### `GetJobs([PrinterId As String], [ByVal Limit As Long = 10]) As Object`

Retrieves (last 10) jobs submitted to any or a particular cloud printer.

#### `DeleteJob(sJobId As String) As Object`

Deletes a job by id.

#### `Property Timeout As Long`

Determines timeout (in seconds) on service calls. `PrintDocument` method triples the value. Default is 10 seconds.

#### `Property LastError As String`

Retrieves last error message.

### `cGcpCallback` class

Used for async service calls by `cGcpService` methods.

#### `Event Complete(oResult As Object)`

Occurs when service call is finished. `oResult` contains parsed JSON result.

### `cGcpOAuth` class

OAuth2 helper class. Wraps `WebBrowser` control navigation and retrieves result of user logon/consent.

#### `Init(oCtl As WebBrowser, sClientId As String, sClientSecret As String, [sScope As String]) As Boolean`

Initializes extension class. `sClientId` and `sClientSecret` usually are hard-coded in the application. These have to be acquired through Google's [APIs Console](https://code.google.com/apis/console#access).

#### `Event Complete(ByVal Allowed As Boolean, DenyReason As String)`

Occurs when user has granted or denied access to GCP service.

#### `Property RefreshToken As String`

On user consent contains OAuth2 `refresh_token` to be used with `cGcpService`.

