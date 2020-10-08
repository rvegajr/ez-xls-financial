using System;
using Xunit;
using EzXlsFinancial.Objects;
using System.Collections.Generic;

namespace EzXlsFinancial.Tests
{
    public class FinancialTests
    {
        [Fact]
        public void NPVCalcOk()
        {
            var xlsFinancial = new NPV();
            
            /* This will allow you to set the formula that will modify the rate before it is used in the calculation,  if left out, it will just use the number passed in the first parameter of calculate */
            xlsFinancial.RateFormula = string.Format("(({0}+1)^(1/12))-1", NPV.RATE_VAR);

            List<double> inputValues = new List<double>();
            /*inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(9021.99);
            inputValues.Add(1020882.11);*/
            
            inputValues.Add(0);
            inputValues.Add(-98089.5);
            inputValues.Add(-14410.5);
            inputValues.Add(-14410.5);
            inputValues.Add(-14410.5);
            inputValues.Add(-14910.5);
            inputValues.Add(-16910.5);
            inputValues.Add(-77410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11910.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-72410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(-11410.5);
            inputValues.Add(821228);
            inputValues.Add(-9872);
            inputValues.Add(1795578);
            inputValues.Add(-2790);
            inputValues.Add(1099710);
            inputValues.Add(999912);

            var val = xlsFinancial.Calculate(0.20, inputValues);

            Assert.True(xlsFinancial.Calculate(0.22, inputValues).ToString("0.00").Equals("890152.51"), "NPV with .22 not correct 890152.51");
            Assert.True(xlsFinancial.Calculate(0.18, inputValues).ToString("0.00").Equals("932217.54"), "NPV with .18 not correct 932217.54");
            Assert.True(xlsFinancial.Calculate(0.15, inputValues).ToString("0.00").Equals("966142.30"), "NPV with .15 not correct 966142.30");

            /* this will write out the resulting excel file to let you play around with it */
            //xlsFinancial.SaveToFile(@"C:\Temp\npv_.xls");
        }

        [Fact]
        public void IRRCalcOk()
        {
            var xlsFinancial = new IRR();

            List<double> inputValues = new List<double>();
            inputValues.Add(0);
            inputValues.Add(-1571991);
            inputValues.Add(13100);
            inputValues.Add(10161);
            inputValues.Add(13100);
            inputValues.Add(13100);
            inputValues.Add(13100);
            inputValues.Add(13100);
            inputValues.Add(13100);
            inputValues.Add(13100);
            inputValues.Add(18100);
            inputValues.Add(13100);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(60000);
            inputValues.Add(0);
            inputValues.Add(5500);
            inputValues.Add(0);
            inputValues.Add(39300);
            inputValues.Add(0);
            inputValues.Add(13100);
            inputValues.Add(0);
            inputValues.Add(13100);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(13100);
            inputValues.Add(13100);
            inputValues.Add(65500);
            inputValues.Add(99965);
            inputValues.Add(113085);
            inputValues.Add(13065);
            inputValues.Add(13065);
            inputValues.Add(13065);
            inputValues.Add(13065);
            inputValues.Add(13065);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(199950);
            inputValues.Add(0);
            inputValues.Add(67015);
            inputValues.Add(13065);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(-8070);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(-10102);
            inputValues.Add(-111);
            inputValues.Add(-17719);
            inputValues.Add(0);
            inputValues.Add(-44453);
            inputValues.Add(-8767);
            inputValues.Add(0);
            inputValues.Add(-11988);
            inputValues.Add(0);
            inputValues.Add(-7363);
            inputValues.Add(-1388);
            inputValues.Add(-26151);
            inputValues.Add(-19162);
            inputValues.Add(-15256);
            inputValues.Add(-23682);
            inputValues.Add(-5901);
            inputValues.Add(0);
            inputValues.Add(-11831);
            inputValues.Add(0);
            inputValues.Add(-2253);
            inputValues.Add(-319);
            inputValues.Add(0);
            inputValues.Add(-2153);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(-158);
            inputValues.Add(-49768);
            inputValues.Add(-16497);
            inputValues.Add(-1035);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(-7000);
            inputValues.Add(-37000);
            inputValues.Add(1929000);
            inputValues.Add(-13000);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);
            inputValues.Add(0);

            var val = xlsFinancial.Calculate(0.01, inputValues);
            Assert.True(xlsFinancial.Calculate(0.01, inputValues).ToString("0.00%").Equals("8.95%"), "IRR with .01 should be 8.95%");
            /* this will write out the resulting excel file to let you play around with it */
            //xlsFinancial.SaveToFile(@"C:\Temp\irr_TEST.xls");
        }
    }
}
