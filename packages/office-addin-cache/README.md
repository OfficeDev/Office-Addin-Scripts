# office-addin-cache

This package clears the Office Add-in cache folders on the user's computer. It supports Windows and macOS. It doesn't clear the browser cache.

## Usage

```
office-addin-cache clear [options]
```

### Options

| Option | Description |
|---|---|
| `--force-close` | Close all running Microsoft Office applications without prompting. |
| `--verbose` | Report the path of each folder being emptied. |

## Cache folders cleared

### Windows

| Label | Path |
|---|---|
| WEF | `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` |
| WebView Cache | `%USERPROFILE%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy` |
| Outlook Hub App Cache | `%USERPROFILE%\AppData\Local\Microsoft\Outlook\HubAppFileCache` |

### macOS

| Label | Path |
|---|---|
| OsfWebHost | `~/Library/Containers/com.Microsoft.OsfWebHost/Data` |

## Office application behavior

If Office applications are running when the command is invoked, the tool will list them and offer to force-close them. If `--force-close` is specified, they will be closed automatically without a prompt. If the applications cannot be closed, the tool will exit without clearing the cache.
