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

            var codeBaseDir = new DirectoryInfo(TestContext.CurrentContext.TestDirectory);
            var assetsDir = Path.Combine(codeBaseDir.FullName, "assets");
            while (!Directory.Exists(assetsDir) &&
                   !Path.GetPathRoot(codeBaseDir.FullName).Equals(codeBaseDir.FullName))
            {
                codeBaseDir = codeBaseDir.Parent;
                assetsDir = Path.Combine(codeBaseDir.FullName, "assets");
            }

            foreach (var assetPath in Directory.GetFiles(assetsDir))
            {
                var assetInTests = Path.Combine(TestContext.CurrentContext.WorkDirectory, Path.GetFileName(assetPath));
                File.Copy(assetPath, assetInTests, true);
            }
        }
    }
}