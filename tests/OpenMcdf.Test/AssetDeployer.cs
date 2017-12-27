using System;
using System.IO;
using NUnit.Framework;

namespace OpenMcdf.Test
{
    [SetUpFixture]
    public class AssetDeployer
    {
        [OneTimeSetUp]
        public void DeployAssetOnce()
        {
            Directory.SetCurrentDirectory(TestContext.CurrentContext.WorkDirectory);

            Console.WriteLine($"Test context working directory: {TestContext.CurrentContext.WorkDirectory}");
            Console.WriteLine($"Test context test directory: {TestContext.CurrentContext.TestDirectory}");

            var codeBaseDir = new DirectoryInfo(TestContext.CurrentContext.TestDirectory);
            var testAssets = Path.Combine("tests", "assets");
            var assetsDir = Path.Combine(codeBaseDir.FullName, testAssets);
            while (!Directory.Exists(assetsDir) &&
                   !Path.GetPathRoot(codeBaseDir.FullName).Equals(codeBaseDir.FullName))
            {
                codeBaseDir = codeBaseDir.Parent;
                assetsDir = Path.Combine(codeBaseDir.FullName, testAssets);
            }

            foreach (var assetPath in Directory.GetFiles(assetsDir))
            {
                var assetInTests = Path.Combine(TestContext.CurrentContext.WorkDirectory, Path.GetFileName(assetPath));
                File.Copy(assetPath, assetInTests, true);
            }
        }
    }
}