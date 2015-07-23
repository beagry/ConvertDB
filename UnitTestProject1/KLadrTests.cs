using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelRLibrary;
using Formater.SupportWorksheetsClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject1
{
    [TestClass]
    public class KLadrTests
    {
        [TestMethod]
        public void DbContains()
        {

            var values = new[]
            {
                new[] {"Брянск","Тимоновский 2-й"},
                new[] {"Брянск","1-й Крайний"},
                new[] {"Брянск","Рылеева"},
                new[] {"Брянск","Ромашина"},
                new[] {"Новосибирск","Весенняя"},
                new[] {"Волгоград","Осенняя"},
            };

            var kladr = new KladrRepository();
            var results = values.Select(ar =>  kladr.IsStreetFromNearCity(ar[1], ar[0]));

            Assert.IsTrue(results.All(r => r));
        }
    }
}
