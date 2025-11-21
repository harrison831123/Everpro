using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EP.PSL.WorkResources.MeetingMng.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var expected = 1;
            var actual = 1;

            // 比較 expected 與 actual 必須是同一個位址參考的物件
            Assert.AreEqual(expected, actual);
        }
    }
}
