# Clean Recent DLL

## `QuickAccessHandler`

```csharp
public class QuickAccess.QuickAccessHandler

```

Methods

| Type | Name | Summary |
| --- | --- | --- |
| `Boolean` | IsAdminPrivilege() | Check whether current user has system administrator privilege. |
| `String` | GetVersion() | Get current assembly's version. |
| `String` | GetSystemUICultureCode() | Get current system's ui culture code. |
| `Boolean` | IsSupportedSystem() | Check whether current system is supported. |
| `void` | AddQuickAccessMenuName(`String` name) | Add given menu name to support unsupported UI system. |
| `void` | AddFileExplorerMenuName(`String`name) | Add given name to fileExplorerName dict. |
| `List<String>` | GetDefaultSupportLanguages() | Get the default support langugaes. <br /> Support `zh-CN, zh-TW, en-US, fr-FR, ru-RU` by default |
| `Dictionary<String, String>` | GetQuickAccessDict() | Get the quick access items in dictionary. <br /> `<item path, item name>` |
| `List<string>` | GetQuickAccessList() | Get the quick access items' paths in list. |
| `Dictionary<String, String>` | GetFrequentFoldersDict() | Get the frequent folders in quick access in dictionary.  <br /> `<folder path, folder name>` |
| `List<string>` | GetFrequentFoldersList() | Get the frequent folders' paths in list. |
| `Dictionary<String, String>` | GetRecentFilesDict() | Get the recent files in quick access in dictionary.  <br /> `<file path, file name>` |
| `List<string>` | GetRecentFilesList() | Get the recent files' paths in list. |
| `Boolean` | IsInQuickAccess(`String` data) | Check whether given string is in quick access. |
| `void`                       | UpdateQuickAccess()                        | Update quick access items.                                   |
| `void` | RemovePathsFromQuickAccess(`List<String>` paths) | Remove given paths from quick access. |
| `void` | RemoveKeywordsFromQuickAccess(`List<String>` keywords) | Remove given keywords from quick access. |
| `void` | RemoveFromQuickAccess(`List<String>` data) | Remove given data from quick access. |
| `void` | EmptyRecentFiles() | Empty recent files. |
| `void` | EmptyFrequentFolders() | Empty frequent folders. |
| `void` | EmptyQuickAccess() | Empty all items in quick access. |


## Notice

### ℹ️ Please check whether current system is supported first!

Use function `IsSupportedSystem` to check whether system is supported, it depends on your system's ui culture.

By default, ui culture `zh-CN, zh-TW, en-US, fr-FR, ru-RU` is supported.

If your system's ui culture is not supported by default, use function `AddQuickAccessMenuName` to add your system's menu name about remove item from quick access before removing items.

