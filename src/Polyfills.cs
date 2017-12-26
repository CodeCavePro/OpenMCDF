using System;
using System.IO;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("OpenMcdf.Test")]
[assembly: InternalsVisibleTo("OpenMcdf.Extensions")]

#if NETSTANDARD1_6

public static class StreamExtension
    {
        public static void Close(this Stream stream)
        {
        }
    }

    public static class BinaryReaderExtension
    {
        public static void Close(this BinaryReader stream)
        {
        }
    }

    public class SerializableAttribute : Attribute
    {
    }

    /// <summary>
    /// Adapted this implementation for .NET Standard 1.6:
    /// https://github.com/dotnet/coreclr/blob/release/1.0.0-rc1/src/mscorlib/src/System/DateTime.cs
    /// </summary>
    public static class DateTimeExtensions
    {
        // Number of 100ns ticks per time unit
        private const long TICKS_PER_MILLISECOND = 10000;
        private const long TICKS_PER_DAY = TICKS_PER_MILLISECOND * 1000 * 60 * 60 * 24;

        // Number of milliseconds per time unit
        private const int MILLIS_PER_DAY = 1000 * 60 * 60 * 24;

        // The minimum OA date is 0100/01/01 (Note it's year 100).
        // The maximum OA date is 9999/12/31
        private const long OA_DATE_MIN_AS_TICKS = (36524 - 365) * TICKS_PER_DAY;

        // All OA dates must be greater than (not >=) OADateMinAsDouble
        private const double OA_DATE_MIN_AS_DOUBLE = -657435.0;
        // All OA dates must be less than (not <=) OADateMaxAsDouble
        private const double OA_DATE_MAX_AS_DOUBLE = 2958466.0;

        private static readonly long DoubleDateOffset = new DateTime(1899, 1, 1).Ticks;

        // Converts the DateTime instance into an OLE Automation compatible
        // double date.
        public static double ToOADate(this DateTime date)
        {
            long value = date.Ticks;
            if (value == 0)
                return 0.0; // Returns OleAut's zero'ed date value.
            if (value < TICKS_PER_DAY
            ) // This is a fix for VB. They want the default day to be 1/1/0001 rather then 12/30/1899.
                value += DoubleDateOffset; // We could have moved this fix down but we would like to keep the bounds check.
            if (value < OA_DATE_MIN_AS_TICKS)
                throw new OverflowException();

            // Currently, our max date == OA's max date (12/31/9999), so we don't 
            // need an overflow check in that direction.
            long millis = (value - DoubleDateOffset) / TICKS_PER_MILLISECOND;
            if (millis < 0)
            {
                long frac = millis % MILLIS_PER_DAY;
                if (frac != 0) millis -= (MILLIS_PER_DAY + frac) * 2;
            }
            return (double)millis / MILLIS_PER_DAY;
        }
        /// <summary>
        /// Creates a DateTime from an OLE Automation Date.
        /// </summary>
        /// <param name="doubleDate">The date in double format.</param>
        /// <returns></returns>
        public static DateTime FromOADate(double doubleDate)
        {
            return new DateTime(DoubleDateToTicks(doubleDate), DateTimeKind.Unspecified);
        }

        // Converts an OLE Date to a tick count.
        // This function is duplicated in COMDateTime.cpp
        internal static long DoubleDateToTicks(double value)
        {
            // The check done this way will take care of NaN
            if (!(value < OA_DATE_MAX_AS_DOUBLE) || !(value > OA_DATE_MIN_AS_DOUBLE))
                throw new ArgumentException();

            // Conversion to long will not cause an overflow here, as at this point the "value" is in between OADateMinAsDouble and OADateMaxAsDouble
            long millis = (long)(value * MILLIS_PER_DAY + (value >= 0 ? 0.5 : -0.5));
            // The interesting thing here is when you have a value like 12.5 it all positive 12 days and 12 hours from 01/01/1899
            // However if you a value of -12.25 it is minus 12 days but still positive 6 hours, almost as though you meant -11.75 all negative
            // This line below fixes up the millis in the negative case
            if (millis < 0)
            {
                millis -= (millis % MILLIS_PER_DAY) * 2;
            }

            millis += DoubleDateOffset / TICKS_PER_MILLISECOND;

            if (millis < 0 || millis >= DateTime.MaxValue.Ticks)
                throw new ArgumentException();
            return millis * TICKS_PER_MILLISECOND;
        }
    }

#endif
