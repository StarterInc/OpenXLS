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
package io.starter.formats.XLS.formulas

//import java.text.SimpleDateFormat;

import io.starter.OpenXLS.DateConverter

import java.util.ArrayList
import java.util.Calendar
import java.util.GregorianCalendar
import java.util.TimeZone


/**
 * DateTimeCalculator is a collection of static methods that operate
 * as the Microsoft Excel function calls do.
 *
 *
 * All methods are called with an array of ptg's, which are then
 * acted upon.  A Ptg of the type that makes sense (ie boolean, number)
 * is returned after each method.
 */

object DateTimeCalculator {
    /**
     * utliity that takes either a PtgStr or a Number and converts it into a calendar
     * for use in below functions
     *
     * @return
     */
    private fun getDateFromPtg(op: Ptg): GregorianCalendar {
        var o: Any
        if (op is PtgStr)
            o = calcDateValue(arrayOf(op)).value
        else if (op is PtgRef) {
            o = op.value
            if (o is String)
                o = calcDateValue(arrayOf(PtgStr(o.toString()))).value
        } else if (op is PtgName) {
            o = op.value    //getComponents()[0];
            o = op.value
            if (o is String)
                o = calcDateValue(arrayOf(PtgStr(o.toString()))).value
        } else
            o = op.value

        return DateConverter.getCalendarFromNumber(o) as GregorianCalendar
    }

    /**
     * DATE
     *
     * @param operands
     * @return
     */
    internal fun calcDate(operands: Array<Ptg>): Ptg {
        val alloperands = PtgCalculator.getLongValueArray(operands)
        if (alloperands.size != 3) return PtgCalculator.error
        val year = alloperands[0].toInt()
        var month = alloperands[1].toInt()
        month = month - 1
        val day = alloperands[2].toInt()
        val c = GregorianCalendar(year, month, day)
        val date = DateConverter.getXLSDateVal(c)
        val i = date.toInt()
        return PtgInt(i)
    }

    /**
     * DATEVALUE
     * Returns the serial number of the date represented by date_text. Use DATEVALUE to convert a date represented by text to a serial number.
     *
     *
     * Syntax
     * DATEVALUE(date_text)
     *
     *
     * Date_text   is text that represents a date in a Microsoft Excel date format. For example, "1/30/2008" or "30-Jan-2008" are text strings
     * within quotation marks that represent dates. Using the default date system in Excel for Windows,
     * date_text must represent a date from January 1, 1900, to December 31, 9999. Using the default date system in Excel for the Macintosh,
     * date_text must represent a date from January 1, 1904, to December 31, 9999. DATEVALUE returns the #VALUE! error value if date_text is out of this range.
     *
     *
     * If the year portion of date_text is omitted, DATEVALUE uses the current year from your computer's built-in clock. Time information in date_text is ignored.
     *
     *
     * Remarks
     *
     *
     * Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Excel for the Macintosh uses a different date system as its default.
     * Most functions automatically convert date values to serial numbers.
     *
     * @param operands Ptg[]
     * @return Ptg
     */

    internal fun calcDateValue(operands: Array<Ptg>?): Ptg {
        // TODO: there may be formats that need to be input
        if (operands == null || operands[0].string == null)
            return PtgErr(PtgErr.ERROR_VALUE)

        val dateString = operands[0].string
        val d = DateConverter.calcDateValue(dateString) ?: return PtgErr(PtgErr.ERROR_VALUE)
        return PtgNumber(d)

    }

    /**
     * DAY
     * Return the day of the month
     */
    internal fun calcDay(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.error
        try {
            val o = operands[0].value
            val c = DateConverter.getCalendarFromNumber(o) as GregorianCalendar
            val retdate = c.get(Calendar.DAY_OF_MONTH)
            return PtgInt(retdate)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * DAYS360
     * Calculate the difference between 2 dates based on a 360
     * day year, 12 mos, 30days each.
     *
     *
     * first date is lower than second, otherwise a negative
     * number is returned
     *
     *
     * Seems pretty dumb to me, but what do I know?
     */
    internal fun calcDays360(operands: Array<Ptg>): Ptg {
        if (operands.size < 2) return PtgErr(PtgErr.ERROR_VALUE)
        try {
            val o1 = operands[0].value
            val o2 = operands[1].value
            val dt1 = DateConverter.getCalendarFromNumber(o1) as GregorianCalendar
            val dt2 = DateConverter.getCalendarFromNumber(o2) as GregorianCalendar
            val yr1 = dt1.get(Calendar.YEAR)
            val yr2 = dt2.get(Calendar.YEAR)
            var diff = yr2 - yr1
            diff *= 360 // turn years to days.
            var mo1 = dt1.get(Calendar.MONTH)
            val mo2 = dt2.get(Calendar.MONTH)
            var mos = 0
            if (mo2 > mo1) {
                mos = mo2 - mo1
            } else {
                diff -= 360
                while (mo2 != mo1) {
                    mos++
                    mo1++
                    if (mo1 == 12) {
                        mo1 = 0
                    }
                }
            }
            diff += mos * 30
            var dy1 = dt1.get(Calendar.DAY_OF_MONTH)
            val dy2 = dt2.get(Calendar.DAY_OF_MONTH)
            if (dy2 > dy1) {
                diff += dy2 - dy1
            } else {
                diff -= 30
                while (dy2 != dy1) {
                    diff++
                    dy1++
                    if (dy1 == 30) {
                        dy1 = 0
                    }
                }
            }
            return PtgInt(diff)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }


    /**
     * EDATE
     * Returns the serial number that represents the date that is the indicated number of months
     * before or after a specified date (the start_date).
     * Use EDATE to calculate maturity dates or due dates that fall on the same day of the month as the date of issue.
     *
     *
     * EDATE(start_date,months)
     *
     *
     * Start_date   is a date that represents the start date. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
     *
     *
     * Months   is the number of months before or after start_date. A positive value for months yields a future date; a negative value yields a past date.
     *
     *
     * If start_date is not a valid date, EDATE returns the #VALUE! error value.
     * If months is not an integer, it is truncated.
     */
    internal fun calcEdate(operands: Array<Ptg>): Ptg {
        try {
            val startDate = getDateFromPtg(operands[0])
            val inc = PtgCalculator.getLongValue(operands[1]).toInt()
            var mm = startDate.get(Calendar.MONTH) + inc
            var y = startDate.get(Calendar.YEAR)
            val d = startDate.get(Calendar.DAY_OF_MONTH)
            if (mm < 0) {
                mm += 12    // 0-based
                y--
            } else if (mm > 11) {
                mm -= 12
                y++
            }
            val resultDate: GregorianCalendar
            resultDate = GregorianCalendar(y, mm, d)
            val retdate = DateConverter.getXLSDateVal(resultDate)
            val i = retdate.toInt()
            return PtgInt(i)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }


    /**
     * EOMONTH
     * Returns the serial number for the last day of the month that is the indicated number of months
     * before or after start_date. Use EOMONTH to calculate maturity dates or due dates that fall on
     * the last day of the month.
     *
     *
     * EOMONTH(start_date,months)
     *
     *
     * Start_date   is a date that represents the starting date. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
     *
     *
     * Months   is the number of months before or after start_date. A positive value for months yields a future date; a negative value yields a past date.
     *
     *
     * If months is not an integer, it is truncated.
     *
     *
     * If start_date is not a valid date, EOMONTH returns the #NUM! error value.
     * If start_date plus months yields an invalid date, EOMONTH returns the #NUM! error value.
     */
    internal fun calcEOMonth(operands: Array<Ptg>): Ptg {
        try {
            val startDate = getDateFromPtg(operands[0])
            val inc = operands[1].intVal
            var mm = startDate.get(Calendar.MONTH) + inc
            var y = startDate.get(Calendar.YEAR)
            var d = startDate.get(Calendar.DAY_OF_MONTH)
            if (mm < 0) {
                mm += 12    // 0-based
                y--
            } else if (mm > 11) {
                mm -= 12
                y++
            }
            if (mm == 3 || mm == 5 || mm == 8 || mm == 10)
            // 0-based
                d = 30
            else if (mm == 1) {// february
                if (y % 4 == 0)
                    d = 29
                else
                    d = 28
            } else
                d = 31
            val resultDate: GregorianCalendar
            resultDate = GregorianCalendar(y, mm, d)
            val retdate = DateConverter.getXLSDateVal(resultDate)
            val i = retdate.toInt()
            return PtgInt(i)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_NUM)
    }


    /**
     * HOUR
     * Converts a serial number to an hour
     */
    internal fun calcHour(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.error
        val dt = getDateFromPtg(operands[0])
        val retdate = dt.get(Calendar.HOUR)
        return PtgInt(retdate)

    }

    /**
     * MINUTE
     * Converts a serial number to a minute
     */
    internal fun calcMinute(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.error
        try {
            val o = operands[0].value
            val dt = DateConverter.getCalendarFromNumber(o) as GregorianCalendar
            val retdate = dt.get(Calendar.MINUTE)
            return PtgInt(retdate)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * MONTH
     * Converts a serial number to a month
     */
    internal fun calcMonth(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.error
        try {
            val dt = getDateFromPtg(operands[0])
            var retdate = dt.get(Calendar.MONTH)
            retdate++ //month is ordinal
            return PtgInt(retdate)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_NA)
    }


    /**
     * NETWORKDAYS
     * Returns the number of whole working days between start_date and end_date. Working days exclude weekends and any dates identified in holidays.
     * Use NETWORKDAYS to calculate employee benefits that accrue based on the number of days worked during a specific term.
     *
     *
     * NETWORKDAYS(start_date,end_date,holidays)
     *
     *
     * Start_date   is a date that represents the start date.
     * End_date   is a date that represents the end date.
     * Holidays   is an optional range of one or more dates to exclude from the working calendar, such as state and federal holidays and floating holidays. The list can be either a range of cells that contains the dates or an array constant of the serial numbers that represent the dates.
     *
     *
     * Remarks
     * If any argument is not a valid date, NETWORKDAYS returns the #VALUE! error value.
     */
    internal fun calcNetWorkdays(operands: Array<Ptg>): Ptg {
        try {
            val holidays = ArrayList()
            val startDate = getDateFromPtg(operands[0])
            val endDate = getDateFromPtg(operands[1])
            if (operands.size > 2 && operands[2] != null) {
                if (operands[2] is PtgRef) {
                    val dts = operands[2].components
                    for (i in dts.indices) {
                        holidays.add(getDateFromPtg(dts[i]))
                    }
                } else
                // assume it's a string or a number rep of a date
                    holidays.add(getDateFromPtg(operands[2]))
            }
            var count = 0
            val countUp = endDate.after(startDate)
            while (startDate != endDate) {
                val d = startDate.get(Calendar.DAY_OF_WEEK)
                if (d != Calendar.SATURDAY && d != Calendar.SUNDAY) {
                    var OKtoIncrement = true
                    // check if on a holidays
                    if (holidays.size > 0) {
                        for (i in holidays.indices) {
                            if (startDate == holidays.get(i)) {
                                OKtoIncrement = false
                                break
                            }
                        }
                    }
                    if (OKtoIncrement) {
                        count++
                    }
                }
                if (countUp)
                    startDate.add(Calendar.DAY_OF_MONTH, 1)
                else
                    startDate.add(Calendar.DAY_OF_MONTH, -1)
            }
            return PtgInt(count)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * NOW
     * Returns the serial number of the current date and time
     */
    internal fun calcNow(operands: Array<Ptg>): Ptg {
        val gc = GregorianCalendar(TimeZone.getDefault())//java.sql.Date dt = new java.sql.Date();
        // io.starter.toolkit.Logger.log(dt.toGMTString());
        val retdate = DateConverter.getXLSDateVal(gc)
        return PtgNumber(retdate)
    }

    /**
     * SECOND
     * Converts a serial number to a second
     */
    internal fun calcSecond(operands: Array<Ptg>): Ptg {
        val dt = getDateFromPtg(operands[0])
        val retdate = dt.get(Calendar.SECOND)
        return PtgInt(retdate)

    }

    /**
     * TIME
     * Returns the serial number of a particular time
     * takes 3 arguments, hour, minute, second;
     */
    internal fun calcTime(operands: Array<Ptg>): Ptg {
        val o: Ptg
        if (operands[0] is PtgStr)
            o = calcDateValue(arrayOf(operands[0]))
        else
            o = operands[0]
        val hour = o.intVal
        val minute = operands[1].intVal
        val second = operands[1].intVal
        val g = GregorianCalendar(2000, 1, 1, hour, minute, second)
        val g2 = GregorianCalendar(2000, 1, 1, 0, 0, 0)
        var dub = DateConverter.getXLSDateVal(g)
        val dub2 = DateConverter.getXLSDateVal(g2)
        dub -= dub2
        return PtgNumber(dub)
    }

    /**
     * TIMEVALUE
     * Converts a time in the form of text to a serial number
     * Returns the decimal number of the time represented by a text string.
     * The decimal number is a value ranging from 0 (zero) to 0.99999999, representing the times from 0:00:00 (12:00:00 AM) to 23:59:59 (11:59:59 P.M.).
     * TIMEVALUE(time_text)
     * Time_text   is a text string that represents a time in any one of the Microsoft Excel time formats; for example, "6:45 PM" and "18:45" text strings within quotation marks that represent time.
     */
    internal fun calcTimevalue(operands: Array<Ptg>): Ptg {
        var result = 0.0
        try {
            val d = getDateFromPtg(operands[0])
            val h = d.get(Calendar.HOUR_OF_DAY)
            val m = d.get(Calendar.MINUTE)
            val s = d.get(Calendar.SECOND)
            val t = h.toDouble() + m / 60.0 + s / (60 * 60.0)
            result = t / 24.0
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_VALUE)
        }

        return PtgNumber(result)
    }


    /**
     * TODAY
     * Returns the serial number of today's date
     */
    internal fun calcToday(operands: Array<Ptg>): Ptg {
        val dt = java.sql.Date(System.currentTimeMillis())
        val retdate = DateConverter.getXLSDateVal(dt)
        val i = retdate.toInt()
        return PtgInt(i)
    }

    /**
     * WEEKDAY  Converts a serial number to a day of the week
     */
    internal fun calcWeekday(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgErr(PtgErr.ERROR_VALUE)
        val dt = getDateFromPtg(operands[0])
        val retdate = dt.get(Calendar.DAY_OF_WEEK)
        return PtgInt(retdate)
    }

    /**
     * WEEKNUM
     * Returns a number that indicates where the week falls numerically within a year.
     *
     *
     * WEEKNUM(serial_num,return_type)
     *
     *
     * Serial_num   is a date within the week. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
     *
     *
     * Return_type   is a number that determines on which day the week begins. The default is 1.
     *
     * @param operands
     * @return
     */
    internal fun calcWeeknum(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        var returnType = 1
        val dt = getDateFromPtg(operands[0])
        if (operands[1] != null) {
            returnType = operands[1].intVal
        }
        returnType -= 1    // 1 is default, 2 =start on monday
        val retdate = dt.get(Calendar.WEEK_OF_YEAR) - returnType
        return PtgInt(retdate)
    }

    /**
     * WORKDAY
     * Returns a number that represents a date that is the indicated number of working days before or after a date (the starting date).
     * Working days exclude weekends and any dates identified as holidays. Use WORKDAY to exclude weekends or holidays when you calculate
     * invoice due dates, expected delivery times, or the number of days of work performed.
     *
     *
     * WORKDAY(start_date,days,holidays)
     *
     *
     * Important   Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
     * Start_date   is a date that represents the start date.
     * Days   is the number of nonweekend and nonholiday days before or after start_date. A positive value for days yields a future date; a negative value yields a past date.
     * Holidays   is an optional list of one or more dates to exclude from the working calendar, such as state and federal holidays and floating holidays. The list can be either a range of cells that contain the dates or an array constant of the serial numbers that represent the dates.
     *
     *
     * Remarks
     *
     *
     * If any argument is not a valid date, WORKDAY returns the #VALUE! error value.
     * If start_date plus days yields an invalid date, WORKDAY returns the #NUM! error value.
     * If days is not an integer, it is truncated.
     *
     * @param operands
     * @return
     */
    internal fun calcWorkday(operands: Array<Ptg>): Ptg {
        val days: Int
        val holidays = ArrayList()
        try {
            val dt = getDateFromPtg(operands[0])
            days = operands[1].intVal
            if (operands.size > 2 && operands[2] != null) {    // holidays
                if (operands[2] is PtgRef) {
                    val dts = operands[2].components
                    for (i in dts.indices) {
                        holidays.add(getDateFromPtg(dts[i]))
                    }
                } else
                // assume it's a string or a number rep of a date
                    holidays.add(getDateFromPtg(operands[2]))
            }
            var absDays = Math.abs(days)
            while (absDays > 0) {
                val d = dt.get(Calendar.DAY_OF_WEEK)
                if (d != Calendar.SATURDAY && d != Calendar.SUNDAY) {
                    var OKtoIncrement = true
                    // check if on a holidays
                    if (holidays.size > 0) {
                        for (i in holidays.indices) {
                            if (dt == holidays.get(i)) {
                                OKtoIncrement = false
                                break
                            }
                        }
                    }
                    if (OKtoIncrement) {
                        absDays--
                    }
                }
                if (days > 0)
                    dt.add(Calendar.DAY_OF_MONTH, 1)
                else
                    dt.add(Calendar.DAY_OF_MONTH, -1)
            }
            return PtgNumber(DateConverter.getXLSDateVal(dt))

        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * YEAR
     * Converts a serial number to a year
     */
    internal fun calcYear(operands: Array<Ptg>): Ptg {
        if (operands.size != 1) return PtgCalculator.error
        try {
            val dt = getDateFromPtg(operands[0])
            val retdate = dt.get(Calendar.YEAR)
            return PtgInt(retdate)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_NA)
    }

    /**
     * YEARFRAC function
     * Calculates the fraction of the year represented by the number of whole days between two dates (the start_date and the end_date). Use the YEARFRAC worksheet function to identify the proportion of a whole year's benefits or obligations to assign to a specific term.
     * YEARFRAC(start_date,end_date,basis)
     * Important  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
     * Start_date   is a date that represents the start date.
     * End_date   is a date that represents the end date.
     * Basis   is the type of day count basis to use.
     * 0 or omitted US (NASD) 30/360
     * 1 Actual/actual
     * 2 Actual/360
     * 3 Actual/365
     * 4 European 30/360
     */
    internal fun calcYearFrac(operands: Array<Ptg>): Ptg {
        val startDate: Long
        val endDate: Long
        try {
            var d = getDateFromPtg(operands[0])
            startDate = DateConverter.getXLSDateVal(d).toLong()
            d = getDateFromPtg(operands[1])
            endDate = DateConverter.getXLSDateVal(d).toLong()
        } catch (e: Exception) {
            //If start_date or end_date are not valid dates, YEARFRAC returns the #VALUE! error value.
            return PtgErr(PtgErr.ERROR_VALUE)
        }

        var basis = 0
        if (operands.size > 2)
            basis = operands[2].intVal
        //If basis < 0 or if basis > 4, YEARFRAC returns the #NUM! error value.
        if (basis < 0 || basis > 4)
            return PtgErr(PtgErr.ERROR_NUM)
        var yf = FinancialCalculator.yearFrac(basis, startDate, endDate)
        if (yf < 0) yf *= -1.0    // =# days between dates, no negatives
        return PtgNumber(yf)
    }
}	