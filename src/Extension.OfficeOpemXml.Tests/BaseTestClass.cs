using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Extension.OfficeOpemXml.Tests
{
    public abstract class BaseTestClass
    {

        public string GetGeneratedFilePath(string name)
        {
            var dir = Path.Combine(TestContext.CurrentContext.TestDirectory, "Generated Files");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, name);
        }

        public string GetResourcesFilePath(string name)
        {
            var dir = Path.Combine(TestContext.CurrentContext.TestDirectory, "Resources");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, name);
        }

    }
}
