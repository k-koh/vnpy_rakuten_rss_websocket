using System;
using System.Collections.Generic;
using System.Linq;

public class VICalculator
{
    private const double SECONDS_PER_YEAR = 365.0 * 24.0 * 60.0 * 60.0;
    private const double MINUTES_PER_YEAR = 365.0 * 24.0 * 60.0;
    private const double RISK_FREE_RATE = 0.001; // Default 0.1%, should be configurable

    public class UnderlyingData
    {
        public double Price { get; set; }
        public DateTime ExpiryDate { get; set; }
    }

    public class OptionData
    {
        public string Symbol { get; set; }
        public double StrikePrice { get; set; }
        public double MidPrice { get; set; }
        public bool IsCall { get; set; }
    }

    /// <summary>
    /// Calculate Nikkei VI from option data
    /// </summary>
    public static double CalculateVI(
        UnderlyingData nearTerm,
        UnderlyingData nextTerm,
        List<OptionData> nearTermOptions,
        List<OptionData> nextTermOptions,
        DateTime currentTime)
    {
        try
        {
            // Calculate time to expiration in minutes
            double T1 = (nearTerm.ExpiryDate - currentTime).TotalMinutes / MINUTES_PER_YEAR;
            double T2 = (nextTerm.ExpiryDate - currentTime).TotalMinutes / MINUTES_PER_YEAR;

            if (T1 <= 0 || T2 <= 0)
            {
                AddinMain.Log("[VICalculator] Error: Expiry time must be in the future");
                return 0;
            }

            // Calculate variance for each term
            double sigma1_squared = CalculateVariance(nearTermOptions, nearTerm.Price, T1);
            double sigma2_squared = CalculateVariance(nextTermOptions, nextTerm.Price, T2);

            // Target 30 days (30/365 years)
            double NT1 = T1 * MINUTES_PER_YEAR;
            double NT2 = T2 * MINUTES_PER_YEAR;
            double N30 = 30.0 * 24.0 * 60.0; // 30 days in minutes
            double N365 = 365.0 * 24.0 * 60.0; // 365 days in minutes

            // Interpolate to get 30-day variance
            double variance30 = (T1 * sigma1_squared * (NT2 - N30) / (NT2 - NT1) +
                                T2 * sigma2_squared * (N30 - NT1) / (NT2 - NT1)) * (N365 / N30);

            // VI = sqrt(variance) * 100
            double vi = Math.Sqrt(variance30) * 100.0;

            //AddinMain.Log($"[VICalculator] T1={T1:F6}, T2={T2:F6}, σ1²={sigma1_squared:F6}, σ2²={sigma2_squared:F6}, VI={vi:F2}");

            return vi;
        }
        catch (Exception ex)
        {
            AddinMain.Log($"[VICalculator] Error calculating VI: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// Calculate variance for a single expiration using CBOE methodology
    /// </summary>
    private static double CalculateVariance(List<OptionData> options, double F, double T)
    {
        if (options == null || options.Count == 0)
            return 0;

        // Group by strike
        var strikes = options.Select(o => o.StrikePrice).Distinct().OrderBy(s => s).ToList();

        // Find ATM strike (closest to forward price)
        double K0 = strikes.OrderBy(k => Math.Abs(k - F)).First();

        // Calculate contribution from each strike
        double sum = 0;
        double prevStrike = 0;

        foreach (var K in strikes)
        {
            // Get put and call at this strike
            var call = options.FirstOrDefault(o => o.StrikePrice == K && o.IsCall);
            var put = options.FirstOrDefault(o => o.StrikePrice == K && !o.IsCall);

            if (call == null && put == null)
                continue;

            // Calculate delta K (width of strike interval)
            double deltaK;
            int idx = strikes.IndexOf(K);
            if (idx == 0)
                deltaK = strikes[1] - strikes[0];
            else if (idx == strikes.Count - 1)
                deltaK = strikes[idx] - strikes[idx - 1];
            else
                deltaK = (strikes[idx + 1] - strikes[idx - 1]) / 2.0;

            // Select option price based on strike position relative to K0
            double optionPrice;
            if (K < K0)
            {
                // Use OTM puts
                optionPrice = put != null ? put.MidPrice : 0;
            }
            else if (K > K0)
            {
                // Use OTM calls
                optionPrice = call != null ? call.MidPrice : 0;
            }
            else
            {
                // At K0, average call and put
                double callPrice = call != null ? call.MidPrice : 0;
                double putPrice = put != null ? put.MidPrice : 0;
                optionPrice = (callPrice + putPrice) / 2.0;
            }

            if (optionPrice > 0)
            {
                double contribution = (deltaK / (K * K)) * Math.Exp(RISK_FREE_RATE * T) * optionPrice;
                sum += contribution;
            }

            prevStrike = K;
        }

        // Calculate variance
        double variance = (2.0 / T) * sum - (1.0 / T) * Math.Pow((F / K0) - 1.0, 2);

        return Math.Max(0, variance); // Ensure non-negative
    }

    /// <summary>
    /// Parse expiry date from DerivMonth string (e.g., "25-01" -> January 2025)
    /// Assumes SQ day is 2nd Friday of the month
    /// </summary>
    public static DateTime ParseExpiryDate(string derivMonth, int currentYear)
    {
        try
        {
            var parts = derivMonth.Split('-');
            if (parts.Length != 2)
                return DateTime.MinValue;

            int year = 2000 + int.Parse(parts[0]);
            int month = int.Parse(parts[1]);

            // Find 2nd Friday (SQ day for Nikkei options)
            DateTime firstDay = new DateTime(year, month, 1);
            int daysUntilFriday = ((int)DayOfWeek.Friday - (int)firstDay.DayOfWeek + 7) % 7;
            DateTime firstFriday = firstDay.AddDays(daysUntilFriday);
            DateTime secondFriday = firstFriday.AddDays(7);

            // Set to 15:15 JST (SQ calculation time)
            return secondFriday.AddHours(15).AddMinutes(15);
        }
        catch
        {
            return DateTime.MinValue;
        }
    }
}
