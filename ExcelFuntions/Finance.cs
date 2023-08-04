using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSTODATN.FunctionsExcel
{
    internal class Finance
    {
        public static double OptimizeInterestRate1(double principal, int loanTerm, double targetProfit)
        {
            double minInterestRate = 0; 
            double maxInterestRate = 100; 
            double step = 0.01; 

            double optimalInterestRate = minInterestRate;
            double optimalProfit = CalculateProfit(principal, loanTerm, optimalInterestRate);

           
            for (double interestRate = minInterestRate; interestRate <= maxInterestRate; interestRate += step)
            {
                double profit = CalculateProfit(principal, loanTerm, interestRate);
                if (profit >= targetProfit)
                {
                    optimalInterestRate = interestRate;
                    break;
                }
                else if (profit > optimalProfit)
                {
                    optimalInterestRate = interestRate;
                    optimalProfit = profit;
                }
            }

            return optimalInterestRate;
        }

        public static double CalculateProfit(double principal, int loanTerm, double interestRate)
        {
            double monthlyInterestRate = interestRate / 12 / 100; // Lãi suất hàng tháng

            double monthlyPayment = principal * monthlyInterestRate / (1 - Math.Pow(1 + monthlyInterestRate, -loanTerm));

            double totalProfit = monthlyPayment * loanTerm - principal;

            return totalProfit;
        }
        /// <summary>
        ///         OptimizeInterestRate function, trả về lãi suất tối ưu
        /// <param name="principal">Số tiền vay</param>
        /// <param name="loanTerm">Số kì trả (tháng)</param>
        /// <param name="targetProfit">Lợi nhuận mục tiêu</param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Trả về lãi suất tối ưu", Name = "OptimizeInterestRate")]
        public static object OptimizeInterestRate(
            [ExcelDna.Integration.ExcelArgument(Description = "Số tiền vay")] double principal,
            [ExcelDna.Integration.ExcelArgument(Description = "Số kỳ trả (tháng)")] int loanTerm,
            [ExcelDna.Integration.ExcelArgument(Description = "Lợi nhuận mục tiêu")] double targetProfit
            )
        {
            return OptimizeInterestRate1(principal, loanTerm, targetProfit);
        }
    }
}
