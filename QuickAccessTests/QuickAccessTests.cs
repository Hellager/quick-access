namespace QuickAccessTests
{
    /// <summary>
    /// Do not directly run or debug or tests!!! <br />
    /// Some of them are risky actions!!! <br />
    /// </summary>
    [TestClass]
    public class QuickAccessTests
    {
        [TestMethod]
        public void GetIsCurrentSystemDefaultSupported()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var isSupportedByDefault = handler.IsSupportedSystem();

            Console.WriteLine("Current system whether supported by default: " + isSupportedByDefault);
        }

        [TestMethod]
        public void AddQuickAccessMenuName_WithGivenName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();
            bool isDefaultSupported = handler.IsSupportedSystem();

            string menuName = "";

            if (!isDefaultSupported)
            {
                handler.AddQuickAccessMenuName(menuName);

                bool isCurrentSupported = handler.IsSupportedSystem();

                Assert.AreNotEqual(isDefaultSupported, isCurrentSupported, "Invalid command name for current system");
            }
        }

        [TestMethod]
        public void AddFileExplorerMenuName_WithGivenName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            string fileExplorerName = "";

            handler.AddFileExplorerMenuName(fileExplorerName);
        }

        [TestMethod]
        public void CheckSupportLanguage_ByDefault()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            List<string> defaultSupportLanguage = new List<string> { "zh-CN", "zh-TW", "en-US", "fr-FR", "ru-RU", "unspecific" };
            var handlerSupportLanguage = handler.GetDefaultSupportLanguages();


            var isSame = defaultSupportLanguage.All(handlerSupportLanguage.Contains) && handlerSupportLanguage.All(defaultSupportLanguage.Contains);

            Assert.IsTrue(isSame, "Missing default support language");
        }

        [TestMethod]
        public void GetQuickAccessDict_CheckReturnType()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var quickAccess = handler.GetQuickAccessDict();

            bool isEqualObj = quickAccess is Dictionary<string, string>;

            Assert.IsTrue(isEqualObj, "Incorrect return type");
        }

        [TestMethod]
        public void GetFrequentFolders_CheckReturnType()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var frequentFolders = handler.GetFrequentFoldersDict();

            bool isEqualObj = frequentFolders is Dictionary<string, string>;
            Assert.IsTrue(isEqualObj, "Incorrect return type");
        }

        [TestMethod]
        public void GetRecentFiles_CheckReturnType()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var recentFiles = handler.GetRecentFilesDict();

            bool isEqualObj = recentFiles is Dictionary<string, string>;
            Assert.IsTrue(isEqualObj, "Incorrect return type");
        }

        [TestMethod]
        public void CheckIsInQuickAccess()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var recentFiles = handler.GetRecentFilesList();

            bool pathRes = handler.IsInQuickAccess(recentFiles.ElementAt(0));
            Assert.IsTrue(pathRes, "Should be in quick access");

            bool keywordRes = handler.IsInQuickAccess("pneumonoultramicroscopicsilicovolcanoconiosis");
            Assert.IsFalse(keywordRes, "Should not be in quick access");
        }

        //[TestMethod]
        //public void RemoveFromQuickAccess_ItemAtFirst()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    var recentFiles = handler.GetRecentFilesList();
        //    var toRemoveTarget = recentFiles.ElementAt(0);

        //    handler.RemoveFromQuickAccess(new List<string> { toRemoveTarget });

        //    bool res = handler.IsInQuickAccess(toRemoveTarget);
        //    Assert.IsFalse(res, "Failed to remove item in quick access");
        //}

        //[TestMethod]
        //public void EmptyRecentFiles()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    handler.EmptyRecentFiles();

        //    var currentRecentFiles = handler.GetRecentFilesDict();
        //    var numberOfCurrentQuickAccess = currentRecentFiles.Count;

        //    Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to empty recent files");
        //}

        //[TestMethod]
        //public void EmptyFrequentFolders()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    handler.EmptyFrequentFolders();

        //    var currentFrequentFolders = handler.GetFrequentFoldersDict();
        //    var numberOfCurrentQuickAccess = currentFrequentFolders.Count;

        //    Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to empty frequent folders");
        //}

        //[TestMethod]
        //public void EmptyQuickAccess()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    handler.EmptyQuickAccess();

        //    var currentQuickAccess = handler.GetQuickAccessDict();
        //    var numberOfCurrentQuickAccess = currentQuickAccess.Count;

        //    Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to empty quick access");
        //}

        [TestMethod]
        public void IsAdminPrivilege_DefaultFalse()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var isAdmin = handler.IsAdminPrivilege();

            Assert.IsFalse(isAdmin, "Current user has no admin priviledge!");
        }
    }
}