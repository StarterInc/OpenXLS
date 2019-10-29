/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with OpenXLS.  If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.OpenXLS;

//import java.text.SimpleDateFormat;

import io.starter.toolkit.Logger;

import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.StringTokenizer;

/**
 * Provides methods for conversion to and from Excel serial date values.
 * <p>
 * Excel stores dates as the number of days since midnight on January 1, 1900.
 * Times are represented as fractional days. For example, 6:00 AM on February 2,
 * 1900 is represented as 33.25. Excel incorrectly treats 1900 as a leap year,
 * so serial dates after February 28, 1900 are one higher than they otherwise
 * should be and the value 60 is unmapped. It also interprets the value 0 as
 * January 0, 1900.
 * <p>
 * Excel does not support negative serial date values, so it cannot handle dates
 * prior to 1900. It also does not currently accept date values with a year of
 * 10000 or greater. OpenXLS does not currently support negative date values,
 * but this feature is planned. If you wish to restrict the output to the subset
 * of values supported by Excel, you may enable input validation by calling
 * {@link #setValidate}.
 * <p>
 * Due to the inherent inaccuracy of floating-point types, values from this
 * class can only be guaranteed to equal values generated by Excel to eight
 * decimal places. This provides accuracy to the unit milliseconds, the maximum
 * precision of Java's date classes. Accuracy outside the range supported by
 * Excel is not guaranteed and will degrade as the values get farther from zero.
 */
public class DateConverter {
    /**
     * The number of milliseconds in a day.
     */
    private static final int MILLIS_DAY = 86400000;

    /**
     * The extra day caused by the 1900 leap year bug.
     */
    private static final int EXTRA_DAY = 60;

    /**
     * Calendar used for date calculation.
     */
    private static Calendar calendar = Calendar.getInstance();

    /**
     * Whether to validate input dates for Excel compatibility.
     */
    private static boolean validate = false;

    /**
     * The set of supported serial date encoding schemes.
     */
    public enum DateFormat {
        /**
         * 1900 epoch with negative value support as used in OOXML.
         * <p>
         * Lower limit: -9999/01/01 00:00:00, value -4 346 018 <br>
         * Epoch date:   1899/12/30 00:00:00, value 0 <br>
         * Upper limit:  9999/12/31 23:59:59, value 2 958 465.999 988 4 <br>
         */
        OOXML_1900(25569, -4346018, 2958465.9999884),

        /**
         * 1900 epoch without negative value support as used in BIFF8.
         * <p>
         * Epoch date:  1899/12/31 00:00:00, value 0 <br>
         * Lower limit: 1900/01/01 00:00:00, value 1 <br>
         * Upper limit: 9999/12/31 23:59:59, value 2 958 465.999 988 4 <br>
         * <p>
         * In this system 1900 is (incorrectly) considered a leap year. Serial
         * dates after 1900/02/28 are one higher than they otherwise should be
         * and the value 60 is unmapped.
         */
        LEGACY_1900(25568, 1, 2958465.9999884),

        /**
         * 1904 epoch without negative value support as used by Excel for Mac.
         * <p>
         * Epoch date:  1904/01/01 00:00:00, value 0 <br>
         * Lower limit: 1904/01/01 00:00:00, value 0 <br>
         * Upper limit: 9999/12/31 23:59:59, value 2 957 003.999 988 4 <br>
         */
        LEGACY_1904(24107, 0, 2957003.9999884);

        private final int epoch_delta;
        private final double limit_lower, limit_upper;

        DateFormat(int delta, double min, double max) {
            epoch_delta = delta;
            limit_lower = min;
            limit_upper = max;
        }

        protected int getEpochDelta() {
            return epoch_delta;
        }

        public double getLowerLimit() {
            return limit_lower;
        }

        public double getUpperLimit() {
            return limit_upper;
        }
    }

    /**
     * Gets a clone of the calendar used for date calculation.
     */
    public static Calendar getCalendar() {
        return (Calendar) calendar.clone();
    }

    /**
     * Sets the calendar used to perform date calculation. This allows you to
     * set the locale and time zone to something other than the system default.
     *
     * @param cal the calendar that should be used for date calculations
     */
    public static void setCalendar(Calendar cal) {
        calendar = cal;
    }

    /**
     * Returns whether input validation is on.
     */
    public static boolean getValidate() {
        return validate;
    }

    /**
     * Sets whether to perform input validation.
     */
    public static void setValidate(boolean validate) {
        DateConverter.validate = validate;
    }

    /**
     * returns whether this method will work with your input string
     */
    public static boolean isParseableDateString(String str) {
        // fix problem with timestamps tagged on end
        if (str.indexOf(" ") > 0) {
            str = str.substring(0, str.indexOf(" "));
        }
        try {
            java.sql.Date d1 = java.sql.Date.valueOf(str);
            getXLSDateVal(d1);
            return true;
        } catch (IllegalArgumentException e) {
            return false;
        }
    }

    /**
     * attempt to interpret a date string into a date
     * returns null if cannot be converted to date
     */
    public static Date getDate(String str) {
        try {
            java.sql.Timestamp d1 = java.sql.Timestamp.valueOf(str);
            Calendar cal = (Calendar) calendar.clone();
            cal.setTime(d1);
            return cal.getTime();
        } catch (Exception e) {
        }
        return null;

    }

    /**
     * Converts the value of the given Calendar to an Excel serial date.
     * The date will be returned in the
     * {@linkplain DateFormat#LEGACY_1900 legacy 1900} format.
     *
     * @param cal the Calendar to convert
     * @return the Excel serial date representing the given calendar's value in
     * the calendar's time zone
     * @deprecated Use {@link #getXLSDateVal(Calendar, DateFormat)} instead.
     */
    public static double getXLSDateVal(Calendar cal) {
        return getXLSDateVal(cal, DateFormat.LEGACY_1900);
    }

    /**
     * Converts the value of the given Calendar to an Excel serial date.
     *
     * @param cal    the Calendar to convert
     * @param format the serial date format to use
     * @return the Excel serial date representing the given calendar's value in
     * the calendar's time zone
     */
    public static double getXLSDateVal(Calendar cal, DateFormat format) {
        // Get the UTC milliseconds since the epoch
        long millis = cal.getTimeInMillis();

        // Add the GMT offset and daylight savings offset for the time zone
        millis += cal.get(Calendar.ZONE_OFFSET) + cal.get(Calendar.DST_OFFSET);

        // Convert from milliseconds to days
        double days = (double) millis / MILLIS_DAY;

        // Switch from UNIX epoch to the Excel epoch
        days += format.getEpochDelta();

        // If the date is after February 28, 1900 add one day.
        // This compensates for Excel incorrectly treating 1900 as a leap year
        if (format == DateFormat.LEGACY_1900 && days >= EXTRA_DAY)
            days += 1;

        // Perform validation
        if (validate && (days < format.getLowerLimit()
                || days > format.getUpperLimit()))
            throw new IllegalArgumentException(
                    "the given date is not supported by Excel");

        // Return the Excel serial date value
        return days;
    }

    /**
     * Converts the given Date to an Excel serial date in the default time zone.
     *
     * @param date   the date to be converted
     * @param format the serial date format to use
     * @return the Excel serial date value corresponding to the given date
     */
    public static double getXLSDateVal(Date date, DateFormat format) {
        Calendar cal = (Calendar) calendar.clone();
        cal.setTime(date);
        return getXLSDateVal(cal, format);
    }

    /**
     * Converts the given Date to an Excel serial date in the default time zone.
     * The date will be returned in the
     * {@linkplain DateFormat#LEGACY_1900 legacy 1900} format.
     *
     * @param date the date to be converted
     * @return the Excel serial date value corresponding to the given date
     * @deprecated Use {@link #getXLSDateVal(Date, DateFormat)} instead.
     */
    public static double getXLSDateVal(Date date) {
        return getXLSDateVal(date, DateFormat.LEGACY_1900);
    }

    /**
     * Parses the the given Excel serial date and returns a Calendar.
     * The date will be interpreted in the
     * {@linkplain DateFormat#LEGACY_1900 legacy 1900} format.
     *
     * @param date the Excel serial date to be interpreted
     * @return a Calendar representing the given Excel date interpreted in the
     * default time zone
     * @deprecated Use {@link #getCalendarFromNumber(double, DateFormat)} instead.
     */
    public static Calendar getCalendarFromNumber(double date) {
        return getCalendarFromNumber(date, DateFormat.LEGACY_1900);
    }

    /**
     * Parses the the given Excel serial date and returns a Calendar.
     *
     * @param date   the Excel serial date to be interpreted
     * @param format the date format with which to interpret the serial date
     * @return a Calendar representing the given Excel date interpreted in the
     * default time zone
     */
    public static Calendar getCalendarFromNumber(double date, DateFormat format) {
        Calendar cal = getCalendar();
        double days = date;

        // For the legacy 1900 epoch, if the date is after 1900/02/28 subtract
        // one day. This compensates for 1900 being considered a leap year.
        // Not matching the non-existent February 29 causes it to become
        // March 1. This behavior matches that of Calendar for the same input.
        if (format == DateFormat.LEGACY_1900 && days > EXTRA_DAY)
            days -= 1;

        // Switch from the Excel epoch to the UNIX epoch
        days -= format.getEpochDelta();

        // Convert from days to milliseconds
        long millis = Math.round(days * MILLIS_DAY);

        // Set the calendar's approximate time so zone offsets are correct
        // The offsets can still be wrong for certain border cases.
        cal.setTimeInMillis(millis);

        // Adjust for time zone and daylight saving time offsets
        long offset = 0;
        for (int count = 0; offset != (offset = cal.get(Calendar.ZONE_OFFSET)
                + cal.get(Calendar.DST_OFFSET))
                && count < 3; count++)
            cal.setTimeInMillis(millis - offset);

        return cal;
    }

    /**
     * Parses the the given Excel serial date and returns a Date.
     * The date will be interpreted in the
     * {@linkplain DateFormat#LEGACY_1900 legacy 1900} format.
     *
     * @param date   the Excel serial date to be interpreted
     * @param format the date format with which to interpret the serial date
     * @return a Calendar representing the given Excel date interpreted in the
     * default time zone
     * @deprecated Use {@link #getCalendarFromNumber(double, DateFormat)} instead.
     */
    public static Date getDateFromNumber(double date) {
        return getCalendarFromNumber(date).getTime();
    }

    /**
     * Gets the Date for the given Excel serial date.
     *
     * @param number a Number representing the serial date to be interpreted
     * @return a Date representing the given Excel date interpreted in the
     * default time zone
     * @throws ClassCastException if the passed object is not a Number
     * @deprecated Use {@link #getDateFromNumber(double)} instead.
     */
    public static Date getDateFromNumber(Object number) {
        if (number instanceof Number)
            return getDateFromNumber(((Number) number).doubleValue());

        else
            throw new ClassCastException("passed object was not a number");
    }

    /**
     * Gets the Date for the given Excel serial date.
     *
     * @param number a Number representing the serial date to be interpreted
     * @return a Date representing the given Excel date interpreted in the
     * default time zone
     * @throws ClassCastException if the passed object is not a Number
     * @deprecated Identical to {@link #getDateFromNumber(Object)}.
     */
    public static Date getNonLocalizedDateFromNumber(Object number) {
        return getDateFromNumber(number);
    }

    /**
     * Gets the Calendar for the given Excel serial date.
     *
     * @param date the Excel serial date or datestring to be interpreted
     * @return a Calendar representing the given Excel date interpreted in the
     * default time zone
     * @throws ClassCastException if the passed object is not a Number
     * @deprecated Use {@link #getCalendarFromNumber(double)} instead.
     */
    public static Calendar getCalendarFromNumber(Object number) {
        if (number == null)
            throw new ClassCastException("object cannot be converted to a date");
        if (!(number instanceof Number))
            number = new Double(number.toString());

        if (number instanceof Number)
            return getCalendarFromNumber(((Number) number).doubleValue());

        else
            throw new ClassCastException("passed object was not a number");
    }

    /**
     * Gets the Date for the given Excel serial date.
     *
     * @param number a Number representing the serial date to be interpreted
     * @return a Date representing the given Excel date interpreted in the
     * default time zone
     * @throws ClassCastException if the passed object is not a Number
     * @deprecated Identical to {@link #getDateFromNumber(Object)}.
     */
    public static Calendar getNonLocalizedCalendarFromNumber(Object number) {
        return getCalendarFromNumber(number);
    }

    /**
     * Gets the Date for the given cell.
     *
     * @param cell a CellHandle whose value should be interpreted
     * @return a Date representing the given Excel date interpreted in the
     * default time zone
     */
    public static Date getDateFromCell(CellHandle cell) {
        return getCalendarFromCell(cell).getTime();
    }

    /**
     * returns a Java Calendar from a CellHandle containing an Excel-formatted
     * Date
     * <p>
     * The Excel date format does not map 100% accurately to Java dates, due to
     * the limitation of the precision of the Excel floating-point value record.
     * Due to this, OpenXLS dates may be too precise, this method will round
     * the java.util.Date returned to the precision entered. Rounding is handled
     * by the ROUND_HALF_UP method.
     * <p>
     * Pass in a static Calendar precision, options are Calendar.HOUR
     * Calendar.MINUTE Calendar.SECOND Calendar.MILLISECOND
     *
     * @return Calendar - A GregorianCalendar value of the Cell
     */
    public static Calendar getCalendarFromCellWithPrecision(Cell cell,
                                                            int roundingCalendarField) {
        Calendar tmp = getCalendarFromCell(cell);
        GregorianCalendar newCalendar = new GregorianCalendar();
        newCalendar.set(Calendar.YEAR, tmp.get(Calendar.YEAR));
        newCalendar.set(Calendar.MONTH, tmp.get(Calendar.MONTH));
        newCalendar.set(Calendar.DAY_OF_MONTH, tmp.get(Calendar.DAY_OF_MONTH));
        newCalendar.set(Calendar.HOUR_OF_DAY, tmp.get(Calendar.HOUR_OF_DAY));
        if (roundingCalendarField == Calendar.HOUR) {
            if (tmp.get(Calendar.MINUTE) > 29) {
                newCalendar.set(Calendar.HOUR_OF_DAY,
                        tmp.get(Calendar.HOUR_OF_DAY) + 1);
            }
            newCalendar.set(Calendar.MINUTE, 0);
            newCalendar.set(Calendar.SECOND, 0);
            newCalendar.set(Calendar.MILLISECOND, 0);
            return newCalendar;
        }
        newCalendar.set(Calendar.MINUTE, tmp.get(Calendar.MINUTE));
        if (roundingCalendarField == Calendar.MINUTE) {
            if (tmp.get(Calendar.SECOND) > 29) {
                newCalendar.set(Calendar.MINUTE, tmp.get(Calendar.MINUTE) + 1);
            }
            newCalendar.set(Calendar.SECOND, 0);
            newCalendar.set(Calendar.MILLISECOND, 0);
            return newCalendar;
        }
        newCalendar.set(Calendar.SECOND, tmp.get(Calendar.SECOND));
        if (roundingCalendarField == Calendar.SECOND) {
            if (tmp.get(Calendar.MILLISECOND) > 499) {
                newCalendar.set(Calendar.SECOND, tmp.get(Calendar.SECOND) + 1);
            }
            newCalendar.set(Calendar.MILLISECOND, 0);
            return newCalendar;
        }
        // round the milliseconds
        newCalendar.set(Calendar.MILLISECOND, tmp.get(Calendar.MILLISECOND));
        return newCalendar;

    }

    /**
     * Gets the Calendar for the given cell.
     *
     * @param cell a CellHandle whose value should be interpreted
     * @return a Calendar representing the given Excel date interpreted in the
     * default time zone
     */
    @SuppressWarnings("deprecation")
    public static Calendar getCalendarFromCell(Cell cell) {
        double value;
        DateFormat format = DateFormat.LEGACY_1900;

        if (cell instanceof CellHandle) {
            CellHandle realCell = (CellHandle) cell;
            value = realCell.getDoubleVal();

            WorkBook book = realCell.getWorkBook();
            if (null != book) {
                format = book.getDateFormat();
            }
        } else {
            value = Double.parseDouble(cell.getVal().toString());
        }

        return getCalendarFromNumber(value, format);
    }

    /**
     * Inspects a string to determine if it is a date. Looks for a pattern such
     * as mm/dd/yy, m/d/yy, mm/dd/yyyy etc.
     *
     * @param possibleDate the string to check for date formats
     * @return whether the given string matches a known date format
     */
    public static boolean isDatePattern(String possibleDate) {
        if (possibleDate.indexOf("/") == -1)
            return false;
        StringTokenizer st = new StringTokenizer(possibleDate, "/");
        return st.countTokens() == 3;
    }

    /**
     * Converts a string representation of a date into a valid calendar object
     * date must be in format mm/dd/yy (or yyyy)
     *
     * @param dateStr
     * @return null if not a valid date
     */
    public static Calendar convertStringToCalendar(String dateStr) {
        if (!isDatePattern(dateStr))
            return null;
        int m, d, y;
        StringTokenizer st = new StringTokenizer(dateStr, "/");
        try {
            m = Integer.valueOf(st.nextToken()).intValue() - 1;
            d = Integer.valueOf(st.nextToken()).intValue();
            String s = st.nextToken();
            int h, mn, sc;
            h = mn = sc = 0;
            if (s.indexOf(" ") > -1) {
                int i = s.indexOf(" ");
                String time = s.substring(i + 1);
                s = s.substring(0, i); // rest of date
                String[] timetokens = time.split(":");
                if (timetokens.length > 0)
                    h = Integer.valueOf(timetokens[0]).intValue();
                if (timetokens.length > 1)
                    mn = Integer.valueOf(timetokens[1]).intValue();
                if (timetokens.length > 2)
                    sc = Integer.valueOf(timetokens[2]).intValue();
            }
            y = Integer.valueOf(s).intValue();
            if (y < 100)
                y += 2000;
            GregorianCalendar cdr = new GregorianCalendar(y, m, d, h, mn, sc);
            return cdr;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Returns the value of the cell as a date formatted as a String date
     * representation. The format is determined by inspecting the excel format.
     * <p>
     * Currently supported formats in this method are: mm/dd/yy dd-mmm-yy dd-mmm
     * mmm-yy mm/dd/yy hh:mm
     * <p>
     * If the cell's date forrmat pattern falls outside of this range, the
     * default output will be in the following format mm/dd/yyyy
     *
     * @return String, the value of a cell formatted as a date
     * @deprecated The date format handling in this method is wildly incorrect.
     * It is retained only to provide compatibility with legacy
     * OpenXLS XML files.
     */
    public static String getFormattedDateVal(CellHandle cell) {
        // assemble the formatting for the javascript.
        if (!cell.isDate()) return "Not a Date";
        if (Double.isNaN(cell.getDoubleVal()))
            return "";    // 20060623 KSC
        Date cal = DateConverter.getDateFromCell(cell);
        FormatHandle f = cell.getFormatHandle();
        int pat = f.getFormatPatternId();
        // SimpleDateFormat sdf = new SimpleDateFormat(); KSC: reuse

        switch (pat) {

            case 0xe:
                WorkBookHandle.simpledateformat.applyPattern("MM/dd/yy");
                return WorkBookHandle.simpledateformat.format(cal);

            case 0xf:
                WorkBookHandle.simpledateformat.applyPattern("dd-MMM-yy");
                return WorkBookHandle.simpledateformat.format(cal);

            case 0x10:
                WorkBookHandle.simpledateformat.applyPattern("dd-MMM");
                return WorkBookHandle.simpledateformat.format(cal);

            case 0x11:
                WorkBookHandle.simpledateformat.applyPattern("MMM-yy");
                return WorkBookHandle.simpledateformat.format(cal);

            case 0x16:
                WorkBookHandle.simpledateformat.applyPattern("MM/dd/yy HH:mm");
                return WorkBookHandle.simpledateformat.format(cal);

            default:
                WorkBookHandle.simpledateformat.applyPattern("MM/dd/yyyy");
                return WorkBookHandle.simpledateformat.format(cal);

        }
    }

    /**
     * @deprecated The date format handling in this method is wildly incorrect.
     * It is retained only to provide compatibility with legacy
     * OpenXLS XML files.
     */
    public static Date parseDate(String s, int pat) {
        if (s.equals(""))
            return null;
        //SimpleDateFormat sdf = new SimpleDateFormat();
        try {
            switch (pat) {
                case 0xe:
                    WorkBookHandle.simpledateformat.applyPattern("dd/MM/yy");
                    return WorkBookHandle.simpledateformat.parse(s);

                case 0xf:
                    WorkBookHandle.simpledateformat.applyPattern("dd-MMM-yy");
                    return WorkBookHandle.simpledateformat.parse(s);

                case 0x10:
                    WorkBookHandle.simpledateformat.applyPattern("dd-MMM");
                    return WorkBookHandle.simpledateformat.parse(s);

                case 0x11:
                    WorkBookHandle.simpledateformat.applyPattern("MMM-yy");
                    return WorkBookHandle.simpledateformat.parse(s);

                case 0x16:
                    WorkBookHandle.simpledateformat.applyPattern("MM/dd/yy HH:mm");
                    return WorkBookHandle.simpledateformat.parse(s);

                default:
                    WorkBookHandle.simpledateformat.applyPattern("dd/MM/yyyy");
                    return WorkBookHandle.simpledateformat.parse(s);

            }
        } catch (Exception e) {
            Logger.logWarn("Failed to parse date " + s + " format pattern: "
                    + pat);
            return Calendar.getInstance().getTime();
        }
    }

    /**
     * DATEVALUE
     * Returns the serial number of the date represented by date_text. Use DATEVALUE to convert a date represented by text to a serial number.
     * <p>
     * Syntax
     * DATEVALUE(date_text)
     * <p>
     * Date_text   is text that represents a date in a Microsoft Excel date format. For example, "1/30/2008" or "30-Jan-2008" are text strings
     * within quotation marks that represent dates. Using the default date system in Excel for Windows,
     * date_text must represent a date from January 1, 1900, to December 31, 9999. Using the default date system in Excel for the Macintosh,
     * date_text must represent a date from January 1, 1904, to December 31, 9999. DATEVALUE returns the #VALUE! error value if date_text is out of this range.
     * <p>
     * If the year portion of date_text is omitted, DATEVALUE uses the current year from your computer's built-in clock. Time information in date_text is ignored.
     * <p>
     * Remarks
     * <p>
     * Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Excel for the Macintosh uses a different date system as its default.
     * Most functions automatically convert date values to serial numbers.
     *
     * @param String dateString
     * @return
     */

    public static Double calcDateValue(String dateString) {
        String[] formats = {"MM/dd/yyyy HH:mm:ss",
                "MM/dd/yy HH:mm:ss",
                "MM/dd/yy",
                "MM/dd/yyyy",
                "MM/d/yyyy HH:mm:ss",
                "MM/d/yy HH:mm:ss",
                "yy-M-d hh:mm:ss a",
                "yy-M-d HH:mm:ss",
                "yy-M-d hh:mm a",
                "yy-M-d HH:mm",
                "dd-M-yyyy hh:mm:ss a",    // 10
                "dd-M-yyyy HH:mm:ss",
                "dd-M-yyyy hh:mm a",
                "dd-M-yyyy HH:mm",
                "dd-MMM-yy hh:mm:ss a",
                "dd-MMM-yy HH:mm:ss",
                "dd-MMM-yy hh:mm a",
                "dd-MMM-yy HH:mm",
                "yyMMdd",
                "MM/d/yy",
                "MM/d/yyyy",            // 20
                "yyyy/MM/dd",
                "d-MMM-yy",
                "d-M-yy",
                "d-M-yyyy",
                "dd-MMM-yy",
                "dd-MMM-yyyy",
                "dd-M-yyyy",
                "dd-MM-yy",
                "dd-MM-yyyy",
                "d-MMM-yyyy",            // 30
                "d-M-yyyy",
                "d-MMM-yyyy",        // 32 really is d-MMM but inorder to match, must append year
                "d/MMM/yyyy",        // 33 really is d/MMM but in order to match, must append year
                "M-yy",
                "MMM-yy",
                "M/d/yyyy",
                "M d, yyyy",
                "yyyy-MM-dd",
                "yy-M-d",
                "yyyy-MM-dd",
                "MMyy",
                "yyMM",
                "yyyyMMddHHmm",
                "yyyyMMddHHmmss",
                "yyMMddHHmmss",
                "MMDDHHmm",
                "MMMM dd, yyyy",
                "E, MMM d, yyyy",
                "E MMM dd, yyyy",
                "EE, MMM dd, yyyy",
                "E, MMMM d, yyyy",
                "E, MMMM dd, yyyy",
                "EE, MMMM dd, yyyy",
                "hh:mm:ss a",
                "HH:mm:ss",
                "hh:mm a",
                "HH:mm",
        };
        for (int i = 0; i < formats.length; i++) {
            try {
                String ds = dateString;
                if (i == 32 || i == 33) { // then must add year to date
                    GregorianCalendar calendar = new GregorianCalendar();
                    int curyear = calendar.get(Calendar.YEAR);
                    if (i == 32) ds = dateString + "-" + curyear;
                    else ds = dateString + "/" + curyear;
                }
                //SimpleDateFormat format= new SimpleDateFormat(formats[i], Locale.ENGLISH);	// 20090701 KSC: apparently need Locale -- why now though??
                //format.setLenient(false);
                WorkBookHandle.simpledateformat.applyLocalizedPattern(formats[i]);
                WorkBookHandle.simpledateformat.setLenient(false);
                //Date d= format.parse(ds);
                Date d = WorkBookHandle.simpledateformat.parse(ds);
                return DateConverter.getXLSDateVal(d);
            } catch (Exception e) {
            }
        }
        return null;
    }

}