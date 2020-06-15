using Dexiom.EPPlusExporter.Extensions;
using Shouldly;
using System;
using System.ComponentModel;
using System.Reflection;
using Xunit;

namespace Dexiom.EPPlusExporter.Extensions.Tests
{
    public class MemberInfoExtensionsTests
    {
        [DisplayName("MyDisplayName")]
        public DateTime MyTestProperty => DateTime.Now;

        [Fact]
        public void GetCustomAttributeTest()
        {
            var prop = typeof(MemberInfoExtensionsTests).GetProperty("MyTestProperty");
            var attr1 = prop.GetCustomAttribute<DisplayNameAttribute>();
            var attr2 = prop.GetCustomAttribute<DisplayNameAttribute>(true);

            attr1.DisplayName.ShouldBe("MyDisplayName");
            attr2.DisplayName.ShouldBe("MyDisplayName");
        }
    }
}