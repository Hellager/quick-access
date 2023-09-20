using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Globalization;


// https://social.msdn.microsoft.com/Forums/sqlserver/en-US/155aba8d-2fa4-49fe-b5ef-1b114a19e5f2/how-to-programmatically-invoke-unpinfromhome-from-c?forum=windowssdk
public enum HRESULT : int
{
    S_OK = 0,
    S_FALSE = 1,
    E_NOINTERFACE = unchecked((int)0x80004002),
    E_NOTIMPL = unchecked((int)0x80004001),
    E_FAIL = unchecked((int)0x80004005)
}

public enum SIGDN : uint
{
    SIGDN_NORMALDISPLAY = 0,
    SIGDN_PARENTRELATIVEPARSING = 0x80018001,
    SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000,
    SIGDN_PARENTRELATIVEEDITING = 0x80031001,
    SIGDN_DESKTOPABSOLUTEEDITING = 0x8004c000,
    SIGDN_FILESYSPATH = 0x80058000,
    SIGDN_URL = 0x80068000,
    SIGDN_PARENTRELATIVEFORADDRESSBAR = 0x8007c001,
    SIGDN_PARENTRELATIVE = 0x80080001,
    SIGDN_PARENTRELATIVEFORUI = 0x80094001
}

public enum SICHINTF : uint
{
    SICHINT_DISPLAY = 0,
    SICHINT_ALLFIELDS = 0x80000000,
    SICHINT_CANONICAL = 0x10000000,
    SICHINT_TEST_FILESYSPATH_IF_NOT_EQUAL = 0x20000000
}

public enum MenuItemInfo_fMask : uint
{
    MIIM_BITMAP = 0x00000080,           //Retrieves or sets the hbmpItem member.
    MIIM_CHECKMARKS = 0x00000008,       //Retrieves or sets the hbmpChecked and hbmpUnchecked members.
    MIIM_DATA = 0x00000020,             //Retrieves or sets the dwItemData member.
    MIIM_FTYPE = 0x00000100,            //Retrieves or sets the fType member.
    MIIM_ID = 0x00000002,               //Retrieves or sets the wID member.
    MIIM_STATE = 0x00000001,            //Retrieves or sets the fState member.
    MIIM_STRING = 0x00000040,           //Retrieves or sets the dwTypeData member.
    MIIM_SUBMENU = 0x00000004,          //Retrieves or sets the hSubMenu member.
    MIIM_TYPE = 0x00000010,             //Retrieves or sets the fType and dwTypeData members.
                                        //MIIM_TYPE is replaced by MIIM_BITMAP, MIIM_FTYPE, and MIIM_STRING.
}

public enum MenuString_Pos : uint
{
    MF_BYCOMMAND = 0x00000000,
    MF_BYPOSITION = 0x00000400,
}

public enum ShowWindowCommands : uint
{
    SW_HIDE = 0,
    SW_SHOWNORMAL = 1,
    SW_NORMAL = 1,
    SW_SHOWMINIMIZED = 2,
    SW_SHOWMAXIMIZED = 3,
    SW_MAXIMIZE = 3,
    SW_SHOWNOACTIVATE = 4,
    SW_SHOW = 5,
    SW_MINIMIZE = 6,
    SW_SHOWMINNOACTIVE = 7,
    SW_SHOWNA = 8,
    SW_RESTORE = 9,
    SW_SHOWDEFAULT = 10,
    SW_FORCEMINIMIZE = 11,
    SW_MAX = 11
}

public enum ShellAddToRecentDocsFlags
{
    SHARD_PIDL = 0x001,
    SHARD_PATHA = 0x002,
    SHARD_PATHW = 0x003
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
public class MENUITEMINFO
{
    public int cbSize;
    public uint fMask;
    public uint fType;
    public uint fState;
    public uint wID;
    public IntPtr hSubMenu;
    public IntPtr hbmpChecked;
    public IntPtr hbmpUnchecked;
    public IntPtr dwItemData;
    public IntPtr dwTypeData;
    public uint cch;
    public IntPtr hbmpItem;

    public MENUITEMINFO()
    {
        cbSize = Marshal.SizeOf(typeof(MENUITEMINFO));
    }
}

[ComImport]
[Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IShellItem
{
    HRESULT BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
    HRESULT GetParent(out IShellItem ppsi);
    HRESULT GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
    HRESULT GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    HRESULT Compare(IShellItem psi, SICHINTF hint, out int piOrder);
}

[ComImport]
[Guid("70629033-e363-4a28-a567-0db78006e6d7")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IEnumShellItems
{
    HRESULT Next(uint celt, out IShellItem rgelt, out uint pceltFetched);
    HRESULT Skip(uint celt);
    HRESULT Reset();
    HRESULT Clone(out IEnumShellItems ppenum);
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
public struct CMINVOKECOMMANDINFO
{
    public int cbSize;
    public int fMask;
    public IntPtr hwnd;
    //public string lpVerb;
    public int lpVerb;
    public string lpParameters;
    public string lpDirectory;
    public int nShow;
    public int dwHotKey;
    public IntPtr hIcon;
}

[ComImport]
[Guid("000214e4-0000-0000-c000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IContextMenu
{
    HRESULT QueryContextMenu(IntPtr hmenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);
    HRESULT InvokeCommand(ref CMINVOKECOMMANDINFO pici);
    HRESULT GetCommandString(uint idCmd, uint uType, out uint pReserved, StringBuilder pszName, uint cchMax);
}

namespace QuickAccess
{
    /// <summary>
    /// Class <c>QuickAccessHandler</c> Handle windows quick access list.
    /// </summary>
    public class QuickAccessHandler
    {
        /// <summary>
        /// Instance variable <c>quickAccessShell</c> <br /> 
        /// A shell instance to handle various actions.
        /// </summary>
        private dynamic quickAccessShell;

        /// <summary>
        /// Instance variable <c>isSupportedSystem</c> <br /> 
        /// Whether current system is supported, depending on system's ui language.
        /// </summary>
        private bool isSupportedSystem;

        /// <summary>
        /// Instance variable <c>quickAccessMenuNames</c> <br />
        /// Menus names about quick access in different language.
        /// </summary>
        private Dictionary<string, List<string>> quickAccessMenuNames;

        /// <summary>
        /// Instance variable <c>fileExplorerName</c> <br />
        /// File explorer name in different language.
        /// </summary>
        private Dictionary<string, List<string>> fileExplorerName;

        /// <summary>
        /// Instance variable <c>SEE_MASK_ASYNCOK</c> fMask value for "unpinfromhome".
        /// </summary>
        private const int SEE_MASK_ASYNCOK = 0x00100000;

        /// <summary>
        /// Instance variable <c>SEE_MASK_ASYNCOK</c> fMask value for "remove from quick access".
        /// </summary>
        private const int CMIC_MASK_ASYNCOK = SEE_MASK_ASYNCOK;

        /// <summary>
        /// Instance variable <c>quickAccessDict</c> Store quick access items. key: item path, value: item name.
        /// </summary>
        private Dictionary<string, string> quickAccessDict;

        /// <summary>
        /// Instance variable <c>frequentFolders</c> <br />
        /// Frequent folders in quick access. <br />
        /// Key: folder path, value: folder name.
        /// </summary>
        private Dictionary<string, string> frequentFolders;

        /// <summary>
        /// Instance variable <c>recentFiles</c> <br />
        /// Recent files in quick access. <br />
        /// Key: file path, value: file name.
        /// </summary>
        private Dictionary<string, string> recentFiles;

        /// <summary>
        /// Instance variable <c>unspecificContent</c> <br />
        /// Items with unspecific type in quick access. <br />
        /// key: item path, value: item name.
        /// </summary>
        private Dictionary<string, string> unspecificContent;

        /// <summary>
        /// Instance variable <c>QuickAccessType</c> Enumeration of quick access items type: FrequentFolder, RecentFile_Win10, RecentFile_Win11, UnSpecificed.
        /// </summary>
        private enum QuickAccessType
        {
            FrequentFolder = 1,
            RecentFile_Win10 = 2,
            RecentFile_Win11 = 3,
            UnSpecificed = 4,
        };

        [DllImport("Shell32.dll", BestFitMapping = false, ThrowOnUnmappableChar = true)]
        private static extern void SHAddToRecentDocs(ShellAddToRecentDocsFlags flags, [MarshalAs(UnmanagedType.LPWStr)] string file);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern HRESULT SHILCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath, out IntPtr ppIdl, ref uint rgflnOut);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern HRESULT SHCreateItemFromIDList(IntPtr pidl, [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

        [DllImport("User32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreatePopupMenu();

        [DllImport("User32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetMenuItemCount(IntPtr hMenu);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool GetMenuItemInfo(IntPtr hMenu, UInt32 uItem, bool fByPosition, [In, Out] MENUITEMINFO lpmii);

        [DllImport("user32.dll")]
        private static extern int GetMenuString(IntPtr hMenu, uint uIDItem, [Out] StringBuilder lpString, int nMaxCount, uint uFlag);

        [DllImport("User32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool DestroyMenu(IntPtr hMenu);

        public QuickAccessHandler()
        {
            this.quickAccessMenuNames = new Dictionary<string, List<string>>
            {
                {"zh-CN", new List<string>{ /*Win10*/ "从“快速访问”", "最近使用", /*Win11*/ "从“最近使用的"}},
                {"zh-TW", new List<string>{ /*Win10*/ "從快速存取移除", "從 [快速存取]", /*Win11*/ "從最近使用中移" }},
                {"en-US", new List<string>{ /*Win10*/ "Remove from Qui", "Unpin from Quic", /*Win11*/ "Remove from Rec"}},
                {"fr-FR", new List<string>{ /*Win10*/ "Supprimer de l'", "Désépingler d", /*Win11*/ "Supprimer de R"}},
                {"ru-RU", new List<string>{ /*Win10*/ "Удалить" /*Win11*/}},
                {"unspecific", new List<string> { } }
            };
            this.fileExplorerName = new Dictionary<string, List<string>>
            {
                {"zh-CN", new List<string>{ "文件资源管理器" } },
                {"zh-TW", new List<string>{ "檔案總管" } },
                {"en-US", new List<string>{ "File Explorerr" } },
                {"fr-FR", new List<string>{ "Explorateur de fichiers" } },
                {"ru-RU", new List<string>{ "Проводник" } },
                {"unspecific", new List<string> { } }
            };
            this.quickAccessShell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));

            this.quickAccessDict = new Dictionary<string, string>();
            this.frequentFolders = new Dictionary<string, string>();
            this.recentFiles = new Dictionary<string, string>();
            this.unspecificContent = new Dictionary<string, string>();

            this.CheckQuickAccess();
            this.isSupportedSystem = CheckIsSupportedSystem();
        }

        /******************************************* General Funtions *******************************************/

        /// <summary>
        /// Remove '"' in string sides.
        /// </summary>
        /// (<paramref name="data"/>).
        /// <returns>
        /// Trimmed string, like "\"hellow\"" to "hello".
        /// </returns>
        /// <param><c>data</c> Given string.</param>
        private string TrimQuotes(string data)
        {
            return data.Trim('"');
        }

        /// <summary>
        /// Check whether given path exists.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <returns>
        /// True if the given path exists, else false.
        /// </returns>
        /// <param><c>path</c> Given path.</param>
        private bool IsValidPath(string path)
        {
            return (File.Exists(path) ^ Directory.Exists(path));
        }

        /// <summary>
        /// Check whether current user has system administrator privilege.
        /// </summary>
        /// <returns>
        /// True if current user has system administrator privilege, else false.
        /// </returns>
        public bool IsAdminPrivilege()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        /// <summary>
        /// Get current assembly's version.
        /// </summary>
        /// <returns>
        /// Current assembly's version.
        /// </returns>
        public string GetVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString(); 
        }

        /// <summary>
        /// Refreshe the file explorer.
        /// </summary>
        private void RefreshFileExplorer()
        {
            // Refer from [Refresh Windows Explorer in Win7](https://stackoverflow.com/questions/2488727/refresh-windows-explorer-in-win7)
            Guid CLSID_ShellApplication = new Guid("13709620-C279-11CE-A49E-444553540000");
            Type shellApplicationType = Type.GetTypeFromCLSID(CLSID_ShellApplication, true);

            object shellApplication = Activator.CreateInstance(shellApplicationType);
            object windows = shellApplicationType.InvokeMember("Windows", BindingFlags.InvokeMethod, null, shellApplication, new object[] { });

            Type windowsType = windows.GetType();
            object count = windowsType.InvokeMember("Count", BindingFlags.GetProperty, null, windows, null);
            for (int i = 0; i < (int)count; i++)
            {
                object item = windowsType.InvokeMember("Item", BindingFlags.InvokeMethod, null, windows, new object[] { i });
                Type itemType = item.GetType();

                // only refresh windows explorers
                string itemName = (string)itemType.InvokeMember("Name", BindingFlags.GetProperty, null, item, null);
                foreach (var nameArr in this.fileExplorerName.Values)
                {
                    if (nameArr.Contains(itemName))
                    {
                        itemType.InvokeMember("Refresh", BindingFlags.InvokeMethod, null, item, null);
                    }
                    else
                    {
                        foreach (var name in nameArr)
                        {
                            if (name.Contains(itemName))
                            {
                                itemType.InvokeMember("Refresh", BindingFlags.InvokeMethod, null, item, null);
                            }
                        }
                    }
                }
            }
        }

        /******************************* Funtions About Internationalization Support *******************************/

        /// <summary>
        /// Check whether given name is in quickAccessMenuNames dict.
        /// </summary>
        /// (<paramref name="name"/>).
        /// <returns>
        /// True if the given name is in quickAccessMenuNames dict, else false.
        /// </returns>
        /// <param><c>name</c> Given name.</param>
        private bool IsInQuickAccessMenuNames(string name)
        {
            foreach (var menuNamesArr in this.quickAccessMenuNames.Values)
            {
                if (menuNamesArr.Contains(name))
                {
                    return true;
                }
                else
                {
                    foreach (var menuName in menuNamesArr)
                    {
                        if (menuName.Contains(name))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Check whether given menu name is supported.
        /// </summary>
        /// (<paramref name="name"/>).
        /// <returns>
        /// True if the given menu name is supported, else false.
        /// </returns>
        /// <param><c>name</c> Given name.</param>
        private bool IsSupportedMenuName(string name)
        {
            foreach (var menuNamesArr in this.quickAccessMenuNames.Values)
            {
                if (menuNamesArr.Contains(name))
                {
                    return true;
                }
                else
                {
                    foreach (var menuName in menuNamesArr)
                    {
                        if (menuName.Contains(name))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Check whether given language code is supported.
        /// </summary>
        /// (<paramref name="code"/>).
        /// <returns>
        /// True if the given language code is supported, else false.
        /// </returns>
        /// <param><c>code</c> Given language code, like 'en-US'.</param>
        private bool IsSupportedLanguage(string code)
        {
            return this.quickAccessMenuNames.ContainsKey(code);
        }

        /// <summary>
        /// Check whether current system's menu name about quick access is supported by default.
        /// </summary>
        /// <returns>
        /// True if is supported by default, else false.
        /// </returns>
        private bool CheckIsDefaultSupportedMenuName()
        {
            if (this.frequentFolders.Keys.Count == 0 && this.recentFiles.Keys.Count == 0)
                return false;

            var selectedFolder = this.frequentFolders.Keys.ToList()[0];
            var selectedFile = recentFiles.Keys.ToList()[0];

            // declare variables
            HRESULT hr = HRESULT.E_FAIL;
            IntPtr pidlFull = IntPtr.Zero;
            uint rgflnOut = 0;
            string sPath = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";

            // Creates a pointer to an item identifier list (PIDL) from a path.
            hr = SHILCreateFromPath(sPath, out pidlFull, ref rgflnOut);
            if (hr == HRESULT.S_OK)
            {
                IntPtr pszName = IntPtr.Zero;
                IShellItem pShellItem;

                // Creates and initializes a Shell item object from a pointer to an item identifier list (PIDL).
                hr = SHCreateItemFromIDList(pidlFull, typeof(IShellItem).GUID, out pShellItem);
                if (hr == HRESULT.S_OK)
                {
                    // Get Windows Quick Access Folder
                    hr = pShellItem.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, out pszName);
                    if (hr == HRESULT.S_OK)
                    {
                        string sDisplayName = Marshal.PtrToStringUni(pszName);
                        // Console.WriteLine(string.Format("Folder NameisDefaultSupport : {0}", sDisplayName));
                        Marshal.FreeCoTaskMem(pszName);
                    }

                    IEnumShellItems pEnumShellItems = null;
                    IntPtr pEnum;
                    Guid BHID_EnumItems = new Guid("94f60519-2850-4924-aa5a-d15e84868039");
                    Guid BHID_SFUIObject = new Guid("3981e225-f559-11d3-8e3a-00c04f6837d5");

                    hr = pShellItem.BindToHandler(IntPtr.Zero, BHID_EnumItems, typeof(IEnumShellItems).GUID, out pEnum);
                    if (hr == HRESULT.S_OK)
                    {
                        pEnumShellItems = Marshal.GetObjectForIUnknown(pEnum) as IEnumShellItems;
                        IShellItem psi = null;
                        uint nFetched = 0;

                        while (HRESULT.S_OK == pEnumShellItems.Next(1, out psi, out nFetched) && nFetched == 1)
                        {
                            // Get Quick Access Item Absolute Path
                            pszName = IntPtr.Zero;
                            hr = psi.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out pszName);

                            if (hr == HRESULT.S_OK)
                            {
                                string sDisplayName = Marshal.PtrToStringUni(pszName);
                                Marshal.FreeCoTaskMem(pszName);

                                if (sDisplayName == selectedFile || sDisplayName == selectedFolder)
                                {
                                    IContextMenu pContextMenu;
                                    IntPtr pcm;
                                    hr = psi.BindToHandler(IntPtr.Zero, BHID_SFUIObject, typeof(IContextMenu).GUID, out pcm);
                                    if (hr == HRESULT.S_OK)
                                    {
                                        pContextMenu = Marshal.GetObjectForIUnknown(pcm) as IContextMenu;
                                        IntPtr hMenu = CreatePopupMenu();
                                        hr = pContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7fff, 0);

                                        if (hr == HRESULT.S_OK)
                                        {
                                            int nNbItems = GetMenuItemCount(hMenu);
                                            for (int i = nNbItems - 1; i >= 0; i--)
                                            {
                                                MENUITEMINFO mii = new MENUITEMINFO();
                                                mii.fMask = (uint)(MenuItemInfo_fMask.MIIM_FTYPE |
                                                            MenuItemInfo_fMask.MIIM_ID |
                                                            MenuItemInfo_fMask.MIIM_SUBMENU |
                                                            MenuItemInfo_fMask.MIIM_DATA);

                                                if (GetMenuItemInfo(hMenu, (uint)i, true, mii))
                                                {
                                                    StringBuilder menuName = new StringBuilder();
                                                    GetMenuString(hMenu, (uint)i, menuName, menuName.Capacity, (uint)MenuString_Pos.MF_BYPOSITION);

                                                    if (menuName.ToString() == "") continue;

                                                    foreach (var menuNamesArr in this.quickAccessMenuNames.Values)
                                                    {
                                                        foreach (var name in menuNamesArr)
                                                        {
                                                            if (name.Contains(menuName.ToString()))
                                                            {
                                                                return true;
                                                            }

                                                        }
                                                    }

                                                }
                                            }
                                        }
                                        DestroyMenu(hMenu);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Check whether current system is supported.
        /// </summary>
        /// <returns>
        /// True if current system is supported, else false.
        /// </returns>
        private bool CheckIsSupportedSystem()
        {
            var currentCulture = CultureInfo.CurrentUICulture.Name;
            if (this.quickAccessMenuNames.ContainsKey(currentCulture))
                return true;

            return CheckIsDefaultSupportedMenuName();
        }

        /// <summary>
        /// Get current system's ui culture code.
        /// </summary>
        /// <returns>
        /// Current system's ui culture code in string.
        /// </returns>
        public string GetSystemUICultureCode()
        {
            return CultureInfo.CurrentUICulture.Name;
        }

        /// <summary>
        /// Check whether current system is supported.
        /// </summary>
        /// <returns>
        /// True if current system is supported, else false.
        /// </returns>
        public bool IsSupportedSystem()
        {
            return this.isSupportedSystem;
        }

        /// <summary>
        /// Add given language code with given menu names list to quickAccessMenuNames dict.
        /// </summary>
        /// (<paramref name="languageCode"/>,<paramref name="meunNames"/>).
        /// <param><c>languageCode</c> Given language code like 'en-US'.</param>
        /// <param><c>meunNames</c> Given menu names list.</param>
        private void AddQuickAccessMenuName(string languageCode, List<string> meunNames)
        {
            if (IsSupportedLanguage(languageCode)) return;

            this.quickAccessMenuNames.Add(languageCode, meunNames);
        }

        /// <summary>
        /// Add given menu name to quickAccessMenuName dict.
        /// </summary>
        /// (<paramref name="name"/>).
        /// <param><c>name</c> Given menu name.</param>
        public void AddQuickAccessMenuName(string name)
        {
            if (name == "") return;

            var unspecificArr = this.quickAccessMenuNames["unspecific"];

            unspecificArr.Add(name);

            this.quickAccessMenuNames["unspecific"] = unspecificArr;
        }

        /// <summary>
        /// Check whether given name is in fileExplorerName dict.
        /// </summary>
        /// (<paramref name="name"/>).
        /// <returns>
        /// True if the given name is in fileExplorerName dict, else false.
        /// </returns>
        /// <param><c>name</c> Given name.</param>
        private bool IsInFileExplorerName(string name)
        {
            foreach (var nameArr in this.fileExplorerName.Values)
            {
                if (nameArr.Contains(name))
                {
                    return true;
                }
                else
                {
                    foreach (var _name in nameArr)
                    {
                        if (_name.Contains(name))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Check whether given languageCode is supported language about file explorer.
        /// </summary>
        /// (<paramref name="languageCode"/>).
        /// <returns>
        /// True if the given languageCode is supported language about file explorer, else false.
        /// </returns>
        /// <param><c>languageCode</c> Given language code like 'en-US'.</param>
        private bool IsSupportedFileExplorerLanguage(string languageCode)
        {
            return this.fileExplorerName.ContainsKey(languageCode);
        }

        /// <summary>
        /// Add given languageCode with given menuName to fileExplorerName dict.
        /// </summary>
        /// (<paramref name="languageCode"/>,<paramref name="menuNames"/>).
        /// <param><c>languageCode</c> Given language code like 'en-US'.</param>
        /// <param><c>menuNames</c> Given menu name list.</param>
        private void AddFileExplorerMenuName(string languageCode, List<string> menuNames)
        {
            if (IsSupportedFileExplorerLanguage(languageCode)) return;

            this.fileExplorerName.Add(languageCode, menuNames);
        }

        /// <summary>
        /// Add given name to fileExplorerName dict.
        /// </summary>
        /// (<paramref name="name"/>).
        /// <param><c>name</c> Given name.</param>
        public void AddFileExplorerMenuName(string name)
        {
            if (name == "") return;

            var unspecificArr = this.fileExplorerName["unspecific"];

            unspecificArr.Add(name);

            this.fileExplorerName["unspecific"] = unspecificArr;
        }

        /// <summary>
        /// Get the default support langugaes. <br/>
        /// Support zh-CN, zh-TW, en-US, fr-FR, ru-RU by default.
        /// </summary>
        /// <returns>
        /// Support language codes in list.
        /// </returns>
        public List<string> GetDefaultSupportLanguages()
        {
            List<string> quickAccess = new List<string>(this.quickAccessMenuNames.Keys);
            List<string> fileExplorer = new List<string>(this.fileExplorerName.Keys);

            return quickAccess.Intersect(fileExplorer).ToList().Distinct().ToList();
        }


        /******************************* Funtions About Quick Access Actions *******************************/

        /// <summary>
        /// Check quick access items.<br />
        /// </summary>
        private void CheckQuickAccess()
        {
            // Refer from [How do I get the name of each item in windows 10 'quick access' items list and put it on a list?](https://stackoverflow.com/questions/41048080/how-do-i-get-the-name-of-each-item-in-windows-10-quick-access-items-list-and-p)
            var CLSID_HomeFolder = new Guid("679f85cb-0220-4080-b29b-5540cc05aab6");
            var quickAccess = this.quickAccessShell.Namespace("shell:::" + CLSID_HomeFolder.ToString("B"));

            this.frequentFolders.Clear();
            this.recentFiles.Clear();
            this.unspecificContent.Clear();
            this.quickAccessDict.Clear();

            foreach (var item in quickAccess.Items())
            {
                var grouping = (int)item.ExtendedProperty("System.Home.Grouping");
                switch (grouping)
                {
                    case (int)QuickAccessType.FrequentFolder:
                        if (this.frequentFolders.ContainsKey(item.path))
                        {
                            this.frequentFolders[item.path] = item.name;
                        }
                        else
                        {
                            this.frequentFolders.Add(item.path, item.name);
                        }
                        break;

                    case (int)QuickAccessType.RecentFile_Win10:
                        if (this.recentFiles.ContainsKey(item.path))
                        {
                            this.recentFiles[item.path] = item.name;
                        }
                        else
                        {
                            this.recentFiles.Add(item.path, item.name);
                        }
                        break;

                    case (int)QuickAccessType.RecentFile_Win11:
                        if (this.recentFiles.ContainsKey(item.path))
                        {
                            this.recentFiles[item.path] = item.name;
                        }
                        else
                        {
                            this.recentFiles.Add(item.path, item.name);
                        }
                        break;

                    default:
                        if (this.unspecificContent.ContainsKey(item.path))
                        {
                            this.unspecificContent[item.path] = item.name;
                        }
                        else
                        {
                            this.unspecificContent.Add(item.path, item.name);
                        }
                        break;
                }
            }

            List<Dictionary<string, string>> quickAccessList = new List<Dictionary<string, string>> { this.frequentFolders, this.recentFiles, this.unspecificContent };

            foreach (Dictionary<string, string> listItem in quickAccessList)
            {
                foreach (KeyValuePair<string, string> quickAccessItem in listItem)
                {
                    this.quickAccessDict[quickAccessItem.Key] = quickAccessItem.Value;
                }
            }
        }

        /// <summary>
        /// Get the quick access items in dictionary. <br />
        /// Key for item path, value for item name.
        /// </summary>
        /// <returns>
        /// Quick access items in dictionary.
        /// </returns>
        public Dictionary<string, string> GetQuickAccessDict()
        {
            CheckQuickAccess();

            return this.quickAccessDict;
        }

        /// <summary>
        /// Get the quick access items' paths in list. <br />
        /// </summary>
        /// <returns>
        /// Quick access items' paths in list.
        /// </returns>
        public List<string> GetQuickAccessList()
        {
            CheckQuickAccess();

            return this.quickAccessDict.Keys.ToList<string>();
        }

        /// <summary>
        /// Get the frequent folders in quick access in dictionary. <br />
        /// Key for folder path, value for folder name.
        /// </summary>
        /// <returns>
        /// Frequent folders in dictionary.
        /// </returns>
        public Dictionary<string, string> GetFrequentFoldersDict()
        {
            CheckQuickAccess();

            return this.frequentFolders;
        }

        /// <summary>
        /// Get the frequent folders' paths in quick access in list. <br />
        /// </summary>
        /// <returns>
        /// Frequent folders' paths in list.
        /// </returns>
        public List<string> GetFrequentFoldersList()
        {
            CheckQuickAccess();

            return this.frequentFolders.Keys.ToList<string>();
        }

        /// <summary>
        /// Get the recent files in quick access in dictionary. <br />
        /// Key for file path, value for file name.
        /// </summary>
        /// <returns>
        /// Recent files in dictionary.
        /// </returns>
        public Dictionary<string, string> GetRecentFilesDict()
        {
            CheckQuickAccess();

            return this.recentFiles;
        }

        /// <summary>
        /// Get the recent files' paths in quick access in list. <br />
        /// </summary>
        /// <returns>
        /// Recent files' paths in list.
        /// </returns>
        public List<string> GetRecentFilesList()
        {
            CheckQuickAccess();

            return this.recentFiles.Keys.ToList<string>();
        }

        /// <summary>
        /// Check whether given path is in quick access.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <returns>
        /// True if the given path is valid and in quick access, else false.
        /// </returns>
        /// <param><c>path</c> Given path,</param>
        private bool IsPathInQuickAccess(string path)
        {
            CheckQuickAccess();

            if (IsValidPath(path))
            {
                foreach (var item in this.quickAccessDict.Keys)
                {
                    if (item.Contains(path))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Check whether given keyword is in quick access.
        /// </summary>
        /// (<paramref name="keyword"/>).
        /// <returns>
        /// True if the given keyword is in quick access, else false.
        /// </returns>
        /// <param><c>keyword</c> Given keyword.</param>
        private bool IsKeywordInQuickAccess(string keyword)
        {
            CheckQuickAccess();

            String[] pathArr = this.quickAccessDict.Keys.ToArray<String>();
            String[] nameArr = this.quickAccessDict.Values.ToArray<String>();
            for (int i = 0; i < pathArr.Length; i++)
            {
                if (pathArr[i].Contains(keyword) || nameArr[i].Contains(keyword))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Check whether given string is in quick access.
        /// </summary>
        /// (<paramref name="data"/>).
        /// <returns>
        /// True if the given string is in quick access, else false.
        /// </returns>
        /// <param><c>data</c> is the given string, whether it's full path or just keyword..</param>
        public bool IsInQuickAccess(string data)
        {
            if (data == "") return false;

            return (IsPathInQuickAccess(data) || IsKeywordInQuickAccess(data));
        }

        /// <summary>
        /// Update quick access items.
        /// </summary>
        public void UpdateQuickAccess()
        {
            this.CheckQuickAccess();
        }

        /// <summary>
        /// Pin folder to quick access. <br />
        /// It may get stuck, timeout treatment is necessary.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <param><c>path</c> Given folder path.</param>
        private void PinFolderToQuickAccess(string path)
        {
            try
            {
                // Solution 1
                // Refer from [Programatically Pin\UnPin the folder from quick access menu in windows 10](https://stackoverflow.com/questions/36739317/programatically-pin-unpin-the-folder-from-quick-access-menu-in-windows-10)
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    runspace.Open();

                    var ps = PowerShell.Create();
                    var shellApplication =
                        ps.AddCommand("New-Object").AddParameter("ComObject", "shell.application").Invoke();

                    dynamic nameSpace = shellApplication.FirstOrDefault()?.Methods["NameSpace"].Invoke(path);
                    nameSpace?.Self.InvokeVerb("pintohome");
                }

                //// Solution 2
                //// Refer from [Is it possible programmatically add folders to the Windows 10 Quick Access panel in the explorer window?](https://stackoverflow.com/questions/30051634/is-it-possible-programmatically-add-folders-to-the-windows-10-quick-access-panel/50032421#50032421)
                //Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
                //Object shell = Activator.CreateInstance(shellAppType);
                //Shell32.Folder2 f = (Shell32.Folder2)shellAppType.InvokeMember("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { path });
                //f.Self.InvokeVerb("pintohome");
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Err: {0}", e);
            }
        }

        /// <summary>
        /// Unpin folder from quick access. <br />
        /// It may get stuck, timeout treatment is necessary.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <param><c>path</c> Given folder path.</param>
        private void UnpinFolderFromQuickAccess(string path)
        {
            try
            {
                // Solution 1
                // Refer from [Programatically Pin\UnPin the folder from quick access menu in windows 10](https://stackoverflow.com/questions/36739317/programatically-pin-unpin-the-folder-from-quick-access-menu-in-windows-10)
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    runspace.Open();
                    var ps = PowerShell.Create();
                    var removeScript =
                        $"((New-Object -ComObject shell.application).Namespace(\"shell:::{{679f85cb-0220-4080-b29b-5540cc05aab6}}\").Items() | Where-Object {{ $_.Path -EQ \"{path}\" }}).InvokeVerb(\"unpinfromhome\")";
                    ps.AddScript(removeScript);
                    ps.Invoke();
                }

                //// Solution 2
                //// Refer from [Is it possible programmatically add folders to the Windows 10 Quick Access panel in the explorer window?](https://stackoverflow.com/questions/30051634/is-it-possible-programmatically-add-folders-to-the-windows-10-quick-access-panel/50032421#50032421)
                //Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
                //Object shell = Activator.CreateInstance(shellAppType);
                //Shell32.Folder2 f2 = (Shell32.Folder2)shellAppType.InvokeMember("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}" });
                //foreach (FolderItem fi in f2.Items())
                //{
                //    if (fi.Path == path)
                //    {
                //        ((FolderItem)fi).InvokeVerb("unpinfromhome");
                //    }
                //}
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Err: {0}", e);
            }
        }

        /// <summary>
        /// Add given path to quick access.<br />
        /// Not working as expect for now, see 
        /// </summary>
        /// (<paramref name="path"/>).
        /// <returns>
        /// True if the given string is valid and in quick access after adding, else false.
        /// </returns>
        /// <param><c>path</c> Given path.</param>
        private bool AddToQuickAccess(string path)
        {
            if (!IsValidPath(path)) return false;
            if (IsPathInQuickAccess(path)) return true;

            if (File.Exists(path))
            {
                SHAddToRecentDocs(ShellAddToRecentDocsFlags.SHARD_PATHW, path);
            }
            else if (Directory.Exists(path))
            {
                PinFolderToQuickAccess(path);
            }
            
            return IsPathInQuickAccess(path);
        }

        /// <summary>
        /// Remove given paths from quick access.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <param><c>path</c> Given path list.</param>
        private void RemoveFromQuickAccessWithList(List<string> paths)
        {
            // declare variables
            HRESULT hr = HRESULT.E_FAIL;
            IntPtr pidlFull = IntPtr.Zero;
            uint rgflnOut = 0;
            string sPath = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";

            // Creates a pointer to an item identifier list (PIDL) from a path.
            hr = SHILCreateFromPath(sPath, out pidlFull, ref rgflnOut);
            if (hr == HRESULT.S_OK)
            {
                IntPtr pszName = IntPtr.Zero;
                IShellItem pShellItem;

                // Creates and initializes a Shell item object from a pointer to an item identifier list (PIDL).
                hr = SHCreateItemFromIDList(pidlFull, typeof(IShellItem).GUID, out pShellItem);
                if (hr == HRESULT.S_OK)
                {
                    // Get Windows Quick Access Folder
                    hr = pShellItem.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, out pszName);
                    if (hr == HRESULT.S_OK)
                    {
                        string sDisplayName = Marshal.PtrToStringUni(pszName);
                        // Console.WriteLine(string.Format("Folder Name : {0}", sDisplayName));
                        Marshal.FreeCoTaskMem(pszName);
                    }

                    IEnumShellItems pEnumShellItems = null;
                    IntPtr pEnum;
                    Guid BHID_EnumItems = new Guid("94f60519-2850-4924-aa5a-d15e84868039");
                    Guid BHID_SFUIObject = new Guid("3981e225-f559-11d3-8e3a-00c04f6837d5");

                    hr = pShellItem.BindToHandler(IntPtr.Zero, BHID_EnumItems, typeof(IEnumShellItems).GUID, out pEnum);
                    if (hr == HRESULT.S_OK)
                    {
                        pEnumShellItems = Marshal.GetObjectForIUnknown(pEnum) as IEnumShellItems;
                        IShellItem psi = null;
                        uint nFetched = 0;

                        while (HRESULT.S_OK == pEnumShellItems.Next(1, out psi, out nFetched) && nFetched == 1)
                        {
                            // Get Quick Access Item Absolute Path
                            pszName = IntPtr.Zero;
                            hr = psi.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out pszName);

                            if (hr == HRESULT.S_OK)
                            {
                                string sDisplayName = Marshal.PtrToStringUni(pszName);
                                // Console.WriteLine(string.Format("\tItem Name : {0}", sDisplayName));
                                Marshal.FreeCoTaskMem(pszName);

                                if (paths.Contains(sDisplayName))
                                {
                                    IContextMenu pContextMenu;
                                    IntPtr pcm;
                                    hr = psi.BindToHandler(IntPtr.Zero, BHID_SFUIObject, typeof(IContextMenu).GUID, out pcm);
                                    if (hr == HRESULT.S_OK)
                                    {
                                        pContextMenu = Marshal.GetObjectForIUnknown(pcm) as IContextMenu;
                                        IntPtr hMenu = CreatePopupMenu();
                                        hr = pContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7fff, 0);

                                        if (hr == HRESULT.S_OK)
                                        {
                                            // http://pastebin.fr/111943
                                            // Handle Quick Access File
                                            int nCommand = -1;
                                            int nNbItems = GetMenuItemCount(hMenu);
                                            for (int i = nNbItems - 1; i >= 0; i--)
                                            {
                                                bool isTargetCommand = false;
                                                MENUITEMINFO mii = new MENUITEMINFO();
                                                mii.fMask = (uint)(MenuItemInfo_fMask.MIIM_FTYPE |
                                                            MenuItemInfo_fMask.MIIM_ID |
                                                            MenuItemInfo_fMask.MIIM_SUBMENU |
                                                            MenuItemInfo_fMask.MIIM_DATA);

                                                // Check Whether Target File has an 'remove from quick access' menu option
                                                if (GetMenuItemInfo(hMenu, (uint)i, true, mii))
                                                {
                                                    StringBuilder menuName = new StringBuilder();
                                                    GetMenuString(hMenu, (uint)i, menuName, menuName.Capacity, (uint)MenuString_Pos.MF_BYPOSITION);

                                                    foreach (var menuNameArr in this.quickAccessMenuNames.Values)
                                                    {
                                                        foreach (var name in menuNameArr)
                                                        {
                                                            if (menuName.ToString().Contains(name))
                                                            {
                                                                nCommand = (int)mii.wID;
                                                                isTargetCommand = true;
                                                                break;
                                                            }
                                                        }
                                                        if (isTargetCommand) break;
                                                    }
                                                }
                                                if (isTargetCommand) break;
                                            }

                                            CMINVOKECOMMANDINFO ici = new CMINVOKECOMMANDINFO();
                                            ici.cbSize = Marshal.SizeOf(ici);
                                            ici.lpVerb = nCommand - 1;
                                            ici.nShow = (int)ShowWindowCommands.SW_NORMAL;
                                            ici.fMask = CMIC_MASK_ASYNCOK;

                                            hr = pContextMenu.InvokeCommand(ici);

                                            if (hr == HRESULT.S_OK)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                Console.Error.WriteLine("Failed to remove target: " + sDisplayName);
                                            }
                                        }
                                        DestroyMenu(hMenu);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Remove given paths from quick access.
        /// </summary>
        /// (<paramref name="paths"/>).
        /// <param><c>data</c> Given paths list.</param>
        public void RemovePathsFromQuickAccess(List<string> paths)
        {
            this.RemoveFromQuickAccess(paths);
        }

        /// <summary>
        /// Remove given keywords from quick access.
        /// </summary>
        /// (<paramref name="paths"/>).
        /// <param><c>data</c> Given paths list.</param>
        public void RemoveKeywordsFromQuickAccess(List<string> keywords)
        {
            this.RemoveFromQuickAccess(keywords);
        }

        /// <summary>
        /// Remove given data from quick access.
        /// </summary>
        /// (<paramref name="data"/>).
        /// <param><c>data</c> Given data list.</param>
        public void RemoveFromQuickAccess(List<string> data)
        {
            CheckQuickAccess();

            List<string> targetList = new List<string> { };

            String[] pathArr = this.quickAccessDict.Keys.ToArray<String>();
            String[] nameArr = this.quickAccessDict.Values.ToArray<String>();

            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < pathArr.Length; j++)
                {
                    if (pathArr[j].Contains(data[i]) || nameArr[j].Contains(data[i]))
                    {
                        targetList.Add(pathArr[j]);
                    }
                }
            }

            if (targetList.Count == 0) return;

            this.RemoveFromQuickAccessWithList(targetList);
        }

        /// <summary>
        /// Empty recent files.
        /// </summary>
        public void EmptyRecentFiles()
        {
            // SHAddToRecentDocs(ShellAddToRecentDocsFlags.SHARD_PIDL, null);
            this.RemoveFromQuickAccessWithList(this.GetRecentFilesList());
        }

        /// <summary>
        /// Empty frequent folders.
        /// </summary>
        public void EmptyFrequentFolders()
        {
            this.RemoveFromQuickAccessWithList(this.GetFrequentFoldersList());
        }

        /// <summary>
        /// Empty all items in quick access.
        /// </summary>
        public void EmptyQuickAccess()
        {
            EmptyRecentFiles();

            EmptyFrequentFolders();
        }
    }
}
