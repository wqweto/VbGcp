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

