# VulDB API VBA Examples

[![License: GPL-3.0](https://img.shields.io/badge/License-GPL--3.0-blue.svg)](LICENSE)
[![Language: VBA](https://img.shields.io/badge/Language-VBA-green.svg)](https://learn.microsoft.com/en-us/office/vba/api/overview/)
[![VulDB API](https://img.shields.io/badge/VulDB-API-blue.svg)](https://vuldb.com/kb/api)

Microsoft Excel and Visual Basic for Applications (VBA) examples for accessing the [VulDB API](https://vuldb.com/kb/api).

[VulDB](https://vuldb.com/) provides vulnerability information and cyber threat intelligence (CTI) for vulnerability management, security research, risk assessment and automated security workflows. This repository provides an Excel-based starting point for integrating VulDB data into Microsoft Office applications without requiring a separate API client application.

## Table of Contents

- [Repository Contents](#repository-contents)
- [Requirements](#requirements)
- [Getting Started](#getting-started)
- [API Authentication](#api-authentication)
- [Standalone VBA Example](#standalone-vba-example)
- [API Query Examples](#api-query-examples)
- [Response Handling](#response-handling)
- [API Credits and Limits](#api-credits-and-limits)
- [Troubleshooting](#troubleshooting)
- [Security Considerations](#security-considerations)
- [Documentation and Support](#documentation-and-support)
- [License](#license)

## Repository Contents

| File | Description |
| --- | --- |
| [`vuldb-api-example.xlsm`](vuldb-api-example.xlsm) | Macro-enabled Microsoft Excel workbook containing the VBA API example. |
| [`README.md`](README.md) | Repository documentation and additional integration examples. |
| [`LICENSE`](LICENSE) | GNU General Public License v3.0. |

## Requirements

- Microsoft Excel desktop with VBA support and permission to execute trusted macros.
- A [VulDB account](https://vuldb.com/) with a valid personal API key and sufficient API credits.
- Outbound HTTPS connectivity to `vuldb.com` on port 443.
- For the standalone VBA example below: Microsoft Windows with the WinHTTP COM component available.

The standalone example uses late binding and does not require an additional VBA library. Compatibility with Excel for macOS or Excel for the web is not established; Excel for the web does not execute VBA macros.

## Getting Started

1. Download [`vuldb-api-example.xlsm`](https://github.com/vuldb/vuldb-api-vba-examples/blob/main/vuldb-api-example.xlsm) using GitHub's download option, or clone the repository:

   ```bash
   git clone https://github.com/vuldb/vuldb-api-vba-examples.git
   ```

2. Open the workbook in Microsoft Excel desktop.
3. Review the workbook and its VBA code before allowing macro execution. Follow your organization's macro security policy.
4. Sign in to [VulDB](https://vuldb.com/) and obtain your personal API key from your user profile.
5. Open the VBA editor (**Alt+F11** on Windows), inspect the supplied modules and locate the API configuration and entry points.
6. Configure authentication as required by the supplied macro and execute the applicable workbook action or macro. To inspect available runnable macros, open **Developer > Macros** (or press **Alt+F8** on Windows).

> [!IMPORTANT]
> Keep your personal API key confidential. Never commit a workbook containing an active API key to a public repository.

## API Authentication

The VulDB API uses HTTPS POST requests to the following endpoint:

```text
https://vuldb.com/?api
```

Supply the API key using either the `X-VulDB-ApiKey` HTTP header or the `apikey` POST field. The standalone example uses the header to keep credentials separate from query parameters.

| Setting | Value |
| --- | --- |
| HTTP method | `POST` |
| Endpoint | `https://vuldb.com/?api` |
| Content type | `application/x-www-form-urlencoded` |
| Authentication header | `X-VulDB-ApiKey` |
| Default response format | JSON |

## Standalone VBA Example

The following additional example demonstrates a direct API request from a standard Excel VBA module. It is independent of the workbook's existing macros and can also be used in a new macro-enabled workbook on Windows.

1. Set the `VULDB_API_KEY` environment variable securely for the Windows user or process running Excel. Restart Excel if necessary so that it inherits the variable.
2. Open the VBA editor and select **Insert > Module**.
3. Paste the following code into the module.
4. Run `VulDBExample` and open the Immediate Window (**Ctrl+G** in the VBA editor) to view the response.

```vb
Option Explicit

Public Sub VulDBExample()
    On Error GoTo RequestFailed

    Dim apiKey As String
    Dim response As String

    apiKey = Environ$("VULDB_API_KEY")

    If Len(Trim$(apiKey)) = 0 Then
        MsgBox "VULDB_API_KEY is not configured.", vbExclamation
        Exit Sub
    End If

    response = VulDBRequest(apiKey, "id=67685&details=0")
    Debug.Print response
    Exit Sub

RequestFailed:
    MsgBox Err.Description, vbExclamation, "VulDB API"
End Sub

Private Function VulDBRequest(ByVal apiKey As String, _
                             ByVal query As String) As String
    Dim http As Object

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    With http
        .SetTimeouts 10000, 10000, 30000, 30000
        .Open "POST", "https://vuldb.com/?api", False
        .SetRequestHeader "Content-Type", _
                          "application/x-www-form-urlencoded"
        .SetRequestHeader "X-VulDB-ApiKey", apiKey
        .SetRequestHeader "Accept", "application/json"
        .Send query & "&version=3&format=json"

        If .Status < 200 Or .Status >= 300 Then
            Err.Raise vbObjectError + 1000, "VulDBRequest", _
                      "HTTP error " & CStr(.Status)
        End If

        VulDBRequest = .ResponseText
    End With
End Function
```

This example retrieves a vulnerability by its VulDB ID and returns the raw JSON response. It does not parse the response or write records into worksheets. For automated processing, add a JSON parser and check the API-level status before using the result.

## API Query Examples

The following examples illustrate API capabilities. They can be used as the `query` argument in the standalone VBA example; they are not a list of implemented workbook functions.

| Use case | POST parameters |
| --- | --- |
| Retrieve a vulnerability by VulDB ID | `id=67685` |
| Retrieve multiple entries | `id=67685,5` |
| Search for a CVE identifier | `search=CVE-2014-6271` |
| Search for a product | `search=Microsoft%20Windows` |
| Retrieve ten recent vulnerabilities | `recent=10` |
| Retrieve ten recently updated vulnerabilities | `updates=10` |
| Retrieve detailed vulnerability data | `id=67685&details=1` |
| Retrieve additional CTI data | `id=67685&details=1&cti=1` |

URL-encode dynamic parameter values before constructing the form body. Access to particular records, details and CTI features depends on the associated license.

See the [API documentation](https://vuldb.com/kb/api) for additional filters, query parameters and output formats.

## Response Handling

The JSON response is organized into three main sections:

| Section | Purpose |
| --- | --- |
| `response` | API status, item count, credit consumption and other metadata. |
| `request` | Information about the processed request. |
| `result` | Returned vulnerability or intelligence records, when available. |

For automated Excel integrations:

- Check both the HTTP status and the API's `response.status` value.
- Parse JSON using a suitable VBA-compatible parser before mapping values to worksheet columns.
- Handle missing fields, empty results and unsuccessful API responses.
- Treat externally retrieved text as data, not as executable worksheet formulas.

The API also supports XML and CSV responses via the `format` parameter. Adapt parsing and import logic to the selected format.

## API Credits and Limits

API access depends on the account's available credits and license. Basic vulnerability requests generally cost one credit; full-detail requests charge per returned vulnerability. CTI-specific requests have separate consumption rules.

Monitor `response.consumption`, `response.remaining` and `response.querylimit` to control usage. Avoid unnecessary repeated requests and use limited queries during development.

See [API Credits](https://vuldb.com/kb/apicredits) for further information.

## Troubleshooting

| Issue | Suggested action |
| --- | --- |
| Excel blocks macros | Review the workbook and follow your organization's approved macro execution policy. |
| VBA cannot create the HTTP object | Confirm Windows WinHTTP COM support for the standalone example. |
| No connection or timeout | Verify internet connectivity, proxy configuration and firewall permissions. |
| API status `204` | No matching records were returned; review the query. |
| API status `400` or `405` | Review query values and supported request types. |
| API status `401` | Verify the API key and account access. |
| API status `402` | Review subscription entitlements. |
| API status `403` | Check remaining credits and account limits before retrying. |
| HTTP status `429` | Reduce request frequency and respect the server's retry guidance. |

Use the API request history in your VulDB user profile to investigate request and response details where available.

## Security Considerations

- Store credentials outside shared workbooks and source code; use an approved secret store for production integrations.
- Use HTTPS with normal certificate validation enabled.
- Review downloaded macros before execution; do not disable macro protection globally.
- Remove API keys and other confidential information from logs, screenshots and issue reports.
- Apply appropriate timeout, error handling and retry policies for production workloads.
- Respect the applicable VulDB data license and access limits.

## Documentation and Support

| Resource | Link |
| --- | --- |
| VulDB | [vuldb.com](https://vuldb.com/) |
| API documentation | [vuldb.com/kb/api](https://vuldb.com/kb/api) |
| API credits | [vuldb.com/kb/apicredits](https://vuldb.com/kb/apicredits) |
| Knowledge base | [vuldb.com/kb](https://vuldb.com/kb) |
| Official GitHub repositories | [github.com/vuldb](https://github.com/vuldb) |
| Repository issues | [GitHub Issues](https://github.com/vuldb/vuldb-api-vba-examples/issues) |
| Support | [support@vuldb.com](mailto:support@vuldb.com) |

For repository-related bug reports or improvements, use GitHub Issues. For questions regarding API access, licensing or integration, contact VulDB support.

## License

The source code and examples in this repository are licensed under the [GNU General Public License v3.0 (GPL-3.0)](LICENSE).

Access to VulDB data and its use are subject to the applicable [VulDB Terms of Use](https://vuldb.com/kb/terms) and subscription conditions. The software license does not grant unrestricted rights to redistribute API data.
