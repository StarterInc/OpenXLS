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

import io.starter.OpenXLS.DateConverter
import io.starter.formats.XLS.FunctionNotSupportedException
import io.starter.toolkit.Logger

import java.util.Calendar
import java.util.GregorianCalendar
import java.util.Vector


/*
 * FinancialCalculator is a collection of static methods that operate as the
 * Microsoft Excel function calls do.
 *
 * All methods are called with an array of ptg's, which are then acted upon. A
 * Ptg of the type that makes sense (ie boolean, number) is returned after each
 * method.
 */

object FinancialCalculator {
    var DEBUG = false

    /*
     * Basis
     *
     * 0, default= US (NASD) 30/360 - As with the European 30/360, with the
     * additional provision that if the end date occurs on the 31st of a month
     * it is moved to the 1st of the next month if the start date is earlier
     * than the 30th. 1= Uses the exact number of elapsed days between the two
     * dates, as well as the exact length of the year. 2= Uses the exact number
     * of elapsed days between two dates but assumes the year only have 360 days
     * 3- Uses the exact number of elapsed days between two dates but assumes
     * the year always has 365 days 4= European 30/360 - Each month is assumed
     * to have 30 days, such that the year has only 360 days. Start and end
     * dates that occur on the 31st of a month become equal to the 30th of the
     * same month.
     *
     * 30/360: If the accrual period ends on a 31st, do not change the date
     * unless the period started on a 30th or 31st, in which case change the end
     * date to 30th. In addition, if the accrual period ends on the last day of
     * February, the month of February should not be extended to a 30 day month.
     * 30/Actual: Method whereby interest is calculated based on a 30-day month
     * and the assumed number of days in a year, i.e. the actual number of days
     * in the accrual period multiplied by the number of interest payments in
     * the year. Eg, a semi-annual bond (one paying two coupons per year) can
     * display a period between coupons of 181 to 184 days. In this case, the
     * number of days in a year will be 362 to 368 days.
     *
     * Euro: Method whereby interest is calculated based on a 30-day month (no
     * exceptions, ie, February should always be extended to a 30 day month) and
     * a 360-day year.
     *
     * Actual: Method whereby interest is calculated based on the actual number
     * of accrued days (falling on a normal year) divided by 365, added to the
     * actual number of accrued days (falling on a leap year) divided by 366.
     *
     * Actual/Actual: Is used for Treasury bonds and notes. This convention it
     * refers to an interest accrual method that utilizes the actual number of
     * days in a month and the actual number of days in a year.
     *
     * Actual/360: A day count fraction equal to actual days divided by 360
     * except in the
     * United Kingdom and several countries where the denominator is 365 or
     * actual days. It is used for bank deposits and in calculating rates pegged
     * to some indices, such as LIBOR. 30/360 Rules: It is used for corporate
     * bonds, U.S. Agency bonds and all mortgage backed securities. It assumes
     * that all months have 30 days, and all years have 360 days. The number of
     * days from M1/D1/Y1 to M2/D2/Y2 is computed according to the following
     * procedure: If D1 is 31, change D1 to 30. If D2 is 31 and D1 is 30 or 31,
     * then change D2 to 30. If M1 is 2, and D1 is 28 (in a non-leap year) or
     * 29, then change D1 to 30. Then the number of days, N is: N = 360(Y2-Y1) +
     * 30(M2-M1) + (D2-D1).
     */
    internal fun yearFrac(basis: Int, date0: Long, date1: Long): Double {
        val result: Double
        // deep breath ...
        val fromDate = DateConverter
                .getCalendarFromNumber(date0) as GregorianCalendar
        val toDate = DateConverter
                .getCalendarFromNumber(date1) as GregorianCalendar
        val y0 = fromDate.get(Calendar.YEAR)
        val y1 = toDate.get(Calendar.YEAR)
        var d0 = fromDate.get(Calendar.DAY_OF_MONTH)
        var d1 = toDate.get(Calendar.DAY_OF_MONTH)
        val m0 = fromDate.get(Calendar.MONTH)
        val m1 = toDate.get(Calendar.MONTH)
        var yearFrac = 1.0
        if (basis == 0) { // 30/360 US/NASD
            if (d0 == 31)
                d0 = 30
            if (d1 == 31 && d0 >= 30)
                d1 = 30
            if (m0 == 1 && d0 >= 28)
                d0 = 30
            yearFrac = (360 * (y1 - y0) + 30 * (m1 - m0) + (d1 - d0)).toDouble() / 360.0
        } else if (basis == 1) { // Actual/Actual
            var ndays = 0
            var i: Int // average # days between dates
            i = y0
            while (i <= y1) {
                ndays += if (isLeapYear(i)) 366 else 365
                i++
            }
            if (i != y0)
                yearFrac = (date1 - date0).toDouble() / (ndays.toDouble() / (i - y0)) // yes I know it's redundant ...
            else
                yearFrac = (date1 - date0).toDouble()
        } else if (basis == 2) { // Actual/360
            yearFrac = (date1 - date0).toDouble() / 360.0
        } else if (basis == 3) { // Actual/365
            yearFrac = (date1 - date0).toDouble() / 365.0
        } else if (basis == 4) { // 30/360 EURO
            //			if (m0==1 && d0>=28) d0= 30; //???????????????????????????
            //				if (m1==1 && d1>=28) d1= 30; //???????????????????????????
            yearFrac = (360 * (y1 - y0) + 30 * (m1 - m0) + (d1 - d0)).toDouble() / 360.0
        }
        return yearFrac
    }

    internal fun getDaysInYearFromBasis(basis: Int, date0: Long, date1: Long): Double {
        var r = 0.0
        when (basis) {
            0    // 30/360
                , 2    // actual/360
                , 4    // 30/360 (EURO)
            -> r = 360.0
            1    // actual/actual
            -> {
                val fromDate = DateConverter.getCalendarFromNumber(date0) as GregorianCalendar
                val y0 = fromDate.get(Calendar.YEAR)
                if (isLeapYear(y0))
                    r = 365.4
                else
                    r = 365.25
            }
            3 -> r = 365.0    // actual/365
        }//				r= (date1-date0)/yearFrac(basis, date0, date1);
        return r
    }

    internal fun getDaysFromBasis(basis: Int, date0: Long, date1: Long): Long {
        if (basis == 0 || basis == 4) { // # months * 30 + extra days
            val fromDate = DateConverter.getCalendarFromNumber(date0) as GregorianCalendar
            val toDate = DateConverter.getCalendarFromNumber(date1) as GregorianCalendar
            val y0 = fromDate.get(Calendar.YEAR)
            val y1 = toDate.get(Calendar.YEAR)
            var d0 = fromDate.get(Calendar.DAY_OF_MONTH)
            var d1 = toDate.get(Calendar.DAY_OF_MONTH)
            val m0 = fromDate.get(Calendar.MONTH)
            val m1 = toDate.get(Calendar.MONTH)
            if (basis == 0) { // 30/360 US/NASD
                // TODO: PROBLEM: When date should be 2-28-02, get
                // 2-2-02!!!!!!!!!!!!!!!
                if (d0 == 31)
                    d0 = 30
                if (d1 == 31 && d0 >= 30)
                    d1 = 30
                //				if (m0==1 && d0 >=28) d0= 30; //
                // ??????????????????????????????????????????????????
                //				if (m1 == 1 && d1 >= 28)
                //					d1 = 30; // ?????????????????
            } else { // 30/360 EURO -- CORRECT FOR ALL DATES EXCEPT MONTH OF
                // FEBRUARY (M0)!!
                //				if (m0==1 && d0>=28) d0= 30; // ?????????????????????
                //				if (m1==1 && d1>=28) d1= 30; // ?????????????????????
            }
            val result = 360 * (y1 - y0) + 30 * (m1 - m0) + (d1 - d0)
            return result.toLong()
        }
        return date1 - date0 // actual (1, 2, 3)
    }

    internal fun validateDay(y: Int, m: Int, d: Int): Int {
        var d = d
        if (d > 28) { // m is 0-based
            if (m == 1)
            // TODO: Get maximum for year
                d = 28
            else if ((m == 3 || m == 5 || m == 8 || m == 10) && d == 31)
                d = 30
        }
        return d
    }

    /**
     * ACCRINT - Returns the accrued interest for a security that pays periodic
     * interest Analysis Pak function Issue is the security's issue date.
     * First_interest is the security's first interest date. Settlement is the
     * security's settlement date. The security settlement date is the date
     * after the issue date when the security is traded to the buyer. Rate is
     * the security's annual coupon rate. Par is the security's par value. If
     * you omit par, ACCRINT uses $1,000. Frequency is the number of coupon
     * payments per year. For annual payments, frequency = 1; for semiannual,
     * frequency = 2; for quarterly, frequency = 4. Basis is the type of day
     * count basis to use.
     *
     *
     * Basis Day count basis 0 or omitted US (NASD) 30/360 1 Actual/actual 2
     * Actual/360 3 Actual/365 4 European 30/360
     *
     *
     * ACCRINT is calculated as follows:
     *
     *
     * par X rate/frequency X Sum Ai/NLi over NC
     *
     *
     * where:
     *
     *
     * Ai = number of accrued days for the ith quasi-coupon period within odd
     * period. NC = number of quasi-coupon periods that fit in odd period. If
     * this number contains a fraction, raise it to the next whole number. NLi =
     * normal length in days of the ith quasi-coupon period within odd period.
     */
    // TODO: Works only for basis 0 & 4
    @Throws(FunctionNotSupportedException::class)
    internal fun calcAccrint(operands: Array<Ptg>): Ptg {
        if (true) {
            val wn = "WARNING: this version of OpenXLS does not support the formula ACCRINT."
            Logger.logWarn(wn)
            throw FunctionNotSupportedException(wn)
        }
        if (operands.size < 6) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG) debugOperands(operands, "ACCRINT")
        try {
            val iDate = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
            val fiDate = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
            val sDate = DateConverter.getCalendarFromNumber(operands[2].value) as GregorianCalendar
            val rate = operands[3].doubleVal
            var par = 1000.0
            if (operands[4] !is PtgMissArg)
                par = operands[4].doubleVal
            val frequency = operands[5].intVal
            var basis = 0
            if (operands.size > 6)
                basis = operands[6].intVal

            val issueDate = DateConverter.getXLSDateVal(iDate).toLong()
            val firstInterestDate = DateConverter.getXLSDateVal(fiDate).toLong()
            val settlementDate = DateConverter.getXLSDateVal(sDate).toLong()

            // TODO: if dates are not valid, return #VALUE! error
            if (!(frequency == 1 || frequency == 2 || frequency == 4)) return PtgErr(PtgErr.ERROR_NUM)
            if (rate <= 0 || par <= 0) return PtgErr(PtgErr.ERROR_NUM)
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
            if (issueDate >= settlementDate) return PtgErr(PtgErr.ERROR_NUM)

            // quasicoupon period: extend the series of equal payment periods to
            // before or after the actual payment periods
            // odd period= period between payments that differs from the usual
            // equally spaced periods at which payments are made
            // Ai= number of days which have accrued for the ith quasi-coupon period
            // within the odd period
            // equation:  ACCRINT= ((par*rate)/frequency)* Sum over NC (Ai/NLi)
            // where NC= # of quasi-coupon periods, rounded up
            //        Ai= # days in odd period i
            //        NLi=normal # days in odd period i
            // For cases where Previous Coupon Date < Issue Date, NC=1, A= getDaysFromBasis and NLi= coupDays

            var result = par * rate / frequency//* Sum #Accrued days for
            val sum = 0.0
            val E: Long
            var PCD: Long

            // ACCRINT in EXCEL handles this case, so we should, too!
            if (firstInterestDate < settlementDate) {
                if (DEBUG) Logger.logInfo(">>> S > FI!")
                /*			Ptg[] ops = new Ptg[4];
			ops[0] = new PtgNumber(settlementDate);
			ops[1] = new PtgNumber(firstInterestDate);
			ops[2] = new PtgInt(frequency);
			ops[3] = new PtgInt(basis);

			long x= PtgCalculator.getLongValue(calcCoupNCD(ops));
*/
                // ?????????????????????????????????
                E = 181
                PCD = issueDate
            } else {
                val ops = arrayOfNulls<Ptg>(4)
                ops[0] = PtgNumber(settlementDate.toDouble())
                ops[1] = PtgNumber(firstInterestDate.toDouble())
                ops[2] = PtgInt(frequency)
                ops[3] = PtgInt(basis)


                E = PtgCalculator.getLongValue(calcCoupDays(ops))
                PCD = PtgCalculator.getLongValue(calcCoupPCD(ops))
            }
            // testing
            if (PCD == 0L) {
                PCD = issueDate + 1
                //Logger.logInfo(">>PCD==0");
            }
            var A: Long = 0
            if (basis == 0 || basis == 4)
            // correct for basis 0, 4
            // INCORRECT for basis 1,2 and 3
                A = getDaysFromBasis(basis, issueDate, settlementDate)
            else if (issueDate >= PCD)
            // correct when issue >= PCD
                A = getDaysFromBasis(basis, Math.max(issueDate, PCD), settlementDate)
            else {
                // ???????????????????????????????????????????????
                A = getDaysFromBasis(basis, Math.max(issueDate, PCD), settlementDate)
                if (DEBUG) Logger.logInfo(">>I < PCD")
            }


            result *= A / E.toDouble()

            val pnum = PtgNumber(result)
            // TODO: Complete Accrint Alogorithm
            if (DEBUG) Logger.logInfo("Result from Accrint= $result")
            return pnum
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * ACCRINTM Returns the accrued interest for a security that pays interest
     * at maturity Analysis Pak function
     *
     *
     * Issue is the security's issue date. Maturity is the security's maturity
     * date. Rate is the security's annual coupon rate. Par is the security's
     * par value. If you omit par, ACCRINTM uses $1,000. Basis is the type of
     * day count basis to use. 0 or omitted US (NASD) 30/360 1 Actual/actual 2
     * Actual/360 3 Actual/365 4 European 30/360 ACCRINTM = par X rate x A/D
     * where: A = Number of accrued days counted according to a monthly basis.
     * For interest at maturity items, the number of days from the issue date to
     * the maturity date is used. D = Annual Year Basis.
     */
    internal fun calcAccrintm(operands: Array<Ptg>): Ptg {
        if (operands.size < 3) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG)
            debugOperands(operands, "ACCRINTM")
        try {
            val issueDate: Long
            val maturityDate: Long // dates are truncated to integers
            val rate: Double
            var par = 1000.0
            var basis = 0

            // Issue Date
            var dt = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
            issueDate = DateConverter.getXLSDateVal(dt).toLong()
            // Maturity Date
            dt = DateConverter
                    .getCalendarFromNumber(operands[1].value) as GregorianCalendar
            maturityDate = DateConverter.getXLSDateVal(dt)
                    .toLong()
            // Annual Coupon Rate
            rate = operands[2].doubleVal
            // Par value. If omitted, = 1000.
            if (operands.size > 3 && operands[3] !is PtgMissArg)
                par = operands[3].doubleVal
            // Basis. If omitted, = 0
            if (operands.size > 4)
                basis = operands[4].intVal

            // TODO: if dates are not valid, return #VALUE! error
            if (rate <= 0 || par <= 0)
                return PtgErr(PtgErr.ERROR_NUM)
            if (basis < 0 || basis > 4)
                return PtgErr(PtgErr.ERROR_NUM)
            if (issueDate > maturityDate)
                return PtgErr(PtgErr.ERROR_NUM)

            val result = par * rate * yearFrac(basis, issueDate, maturityDate)
            if (DEBUG) Logger.logInfo("Result from Accrintm= $result")
            return PtgNumber(result)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * AMORDEGRC Returns the depreciation for each accounting period
     */
    internal fun calcAmordegrc(operands: Array<Ptg>): Ptg {
        if (operands.size < 6) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG)
            debugOperands(operands, "AMORDEGRC")
        try {
            var cost = operands[0].doubleVal
            val dP = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
            val fP = DateConverter.getCalendarFromNumber(operands[2].value) as GregorianCalendar
            val salvage = operands[3].doubleVal
            val period = operands[4].intVal
            var rate = operands[5].doubleVal
            val basis = operands[6].intVal

            val datePurchased = DateConverter.getXLSDateVal(dP).toLong()
            val firstPeriod = DateConverter.getXLSDateVal(fP).toLong()

            // TODO: if dates are not valid, return #VALUE! error
            if (rate <= 0) return PtgErr(PtgErr.ERROR_NUM)
            if (basis < 0 || basis > 4 || basis == 2) return PtgErr(PtgErr.ERROR_NUM)
            if (datePurchased > firstPeriod) return PtgErr(PtgErr.ERROR_NUM)

            val coefficient: Double
            val life = 1.0 / rate
            if (life < 3)
                coefficient = 1.0
            else if (life < 5.0)
            // between 3 and 4 years
                coefficient = 1.5
            else if (life <= 6.0)
            // between 5 and 6 years
                coefficient = 2.0
            else
            // more than 6 years
                coefficient = 2.5

            rate *= coefficient
            //cost-= Math.round(yearFrac(basis, datePurchased,
            // firstPeriod)*rate*cost);
            cost -= yearFrac(basis, datePurchased, firstPeriod) * rate * cost
            var Remainder = cost - salvage
            var A = 0.0
            if (Remainder > 0) {
                var i = 0
                while (i < period) {
                    //A= Math.round(rate*cost);
                    A = rate * cost
                    Remainder -= A
                    if (Remainder < 0) {
                        /*
                         * if (period==i+1) { // A= Math.round(0.5*cost);
                         * Remainder=A; }
                         */
                        if (i + 1 < period)
                            A = 0.0
                        i = period // exit loop
                    }

                    cost -= A
                    i++
                }
            } else
                A = Math.round(rate * salvage).toDouble()
            val result = Math.max(Math.round(A), 0).toDouble()
            if (DEBUG)
                Logger.logInfo("Result from AMORDEGRC= $result")
            return PtgNumber(result)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }

    /**
     * AMORLINC Returns the depreciation for each accounting period
     */
    internal fun calcAmorlinc(operands: Array<Ptg>): Ptg {
        if (operands.size < 6) {
            return PtgErr(PtgErr.ERROR_NULL)
        }
        if (DEBUG)
            debugOperands(operands, "AMORLINC")
        try {
            val cost = operands[0].doubleVal
            val dP = DateConverter
                    .getCalendarFromNumber(operands[1].value) as GregorianCalendar
            val fP = DateConverter
                    .getCalendarFromNumber(operands[2].value) as GregorianCalendar
            val salvage = operands[3].doubleVal
            val period = operands[4].intVal
            val rate = operands[5].doubleVal
            val basis = operands[6].intVal

            val datePurchased = DateConverter.getXLSDateVal(dP)
                    .toLong()
            val firstPeriod = DateConverter.getXLSDateVal(fP)
                    .toLong()

            // TODO: if dates are not valid, return #VALUE! error
            if (rate <= 0)
                return PtgErr(PtgErr.ERROR_NUM)
            if (basis < 0 || basis > 4 || basis == 2)
                return PtgErr(PtgErr.ERROR_NUM)
            if (datePurchased > firstPeriod)
                return PtgErr(PtgErr.ERROR_NUM)

            var A = 0.0
            val B = cost - salvage
            val C = yearFrac(basis, datePurchased, firstPeriod) * rate * cost
            val D = cost * rate
            val n = Math.round((cost - salvage - C) / D)
            if (period == 0 || C == 0.0)
                A = C
            else if (period < n)
                A = D
            else if (period.toLong() == n)
                A = Math.min(D + (B - D * n - C), D)
            else if (period.toLong() == n + 1)
                A = B - D * n - C
            else
                A = 0.0
/*
		cost -= yearFrac(basis, datePurchased, firstPeriod) * rate * cost;
		if (Remainder > 0) {
			for (int i = 0; i < period; i++) {
				//A= Math.round(rate*cost);
				A = rate * cost;
				Remainder -= A;
				if (Remainder < 0) {
					/*
					 * if (period==i+1) { // A= Math.round(0.5*cost);
					 }
					 * /
					if (i + 1 < period)
						A = 0;
					i = period; // exit loop
				}
				cost -= A;
			}
		} else
			A = Math.round(rate * salvage);
*/
            val result = Math.max(Math.round(A), 0).toDouble()
 /*
		double A = cost * rate;
		double B = cost - salvage;
		//double yf= yearFrac(basis, datePurchased, firstPeriod);
		double C = yearFrac(basis, datePurchased, firstPeriod) * rate * cost;
		long n = Math.round((cost - salvage - C) / A);
		double result;
		if (period == 0 || C == 0)
			result = C;
		else if (period < n)
			result = A;
		else if (period == n)
			result = A - C; //B - A * n - C;
		else if (period == n + 1)
			result = A - C; //B - A * n - C;
		else
			result = 0;
*/
            if (DEBUG)
Logger.logInfo("Result from AMORLINC= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * COUPDAYBS --
 *
 *
 * Returns the number of days from the beginning of the coupon period to the
 * settlement date Settlement is the security's settlement date. The
 * security settlement date is the date after the issue date when the
 * security is traded to the buyer. Maturity is the security's maturity
 * date. The maturity date is the date when the security expires. Frequency
 * is the number of coupon payments per year. For annual payments, frequency =
 * 1; for semiannual, frequency = 2; for quarterly, frequency = 4. Basis is
 * the type of day count basis to use. (optional)
 */
    internal fun calcCoupDayBS(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG) debugOperands(operands, "COUPDAYSBS")
try
{
val sDate = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val settlementDate = (DateConverter.getXLSDateVal(sDate)).toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate)).toLong()
val frequency = operands[2].intVal
var basis = 0
if (operands.size > 3)
basis = operands[3].intVal
 //			 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (!(frequency == 1 || frequency == 2 || frequency == 4)) return PtgErr(PtgErr.ERROR_NUM)
if (settlementDate > maturityDate) return PtgErr(PtgErr.ERROR_NUM)

val pcd = PtgCalculator.getLongValue(calcCoupPCD(operands))

val result = getDaysFromBasis(basis, pcd, settlementDate).toDouble()

 //


            if (DEBUG)
Logger.logInfo("Result from calcCoupDaysBS= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * COUPDAYS Returns the number of days in the coupon period that contains
 * the settlement date Settlement is the security's settlement date. The
 * security settlement date is the date after the issue date when the
 * security is traded to the buyer Maturity is the security's maturity date.
 * The maturity date is the date when the security expires. Frequency is the
 * number of coupon payments per year. For annual payments, frequency = 1;
 * for semiannual, frequency = 2; for quarterly, frequency = 4. Basis is the
 * type of day count basis to use
 */
    internal fun calcCoupDays(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "COUPDAYS")
try
{
var dt = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val settlementDate = (DateConverter.getXLSDateVal(dt)).toLong()
dt = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val maturityDate = (DateConverter.getXLSDateVal(dt)).toLong()
val frequency = operands[2].intVal
var basis = 0
if (operands.size > 3)
basis = operands[3].intVal
 // TODO: if dates are not valid, return #VALUE! error
            if (!(frequency == 1 || frequency == 2 || frequency == 4)) return PtgErr(PtgErr.ERROR_NUM)
if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (settlementDate > maturityDate) return PtgErr(PtgErr.ERROR_NUM)

 // VERY strange, but seems to be correct
            val result:Double
if (basis == 1)
{    // actual/actual
val pcd = PtgCalculator.getLongValue(calcCoupPCD(operands))
val ncd = PtgCalculator.getLongValue(calcCoupNCD(operands))
result = getDaysFromBasis(basis, pcd, ncd).toDouble()
}
else if (basis == 0 || basis == 2 || basis == 4)
result = 360.0 / frequency
else
result = 365.0 / frequency
if (DEBUG)
Logger.logInfo("Result from calcCoupDays=" + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * COUPDAYSNC - Returns the number of days from the settlement date to the
 * next coupon date
 *
 *
 * Settlement is the security's settlement date. The security settlement
 * date is the date after the issue date when the security is traded to the
 * buyer Maturity is the security's maturity date. The maturity date is the
 * date when the security expires. Frequency is the number of coupon
 * payments per year. For annual payments, frequency = 1; for semiannual,
 * frequency = 2; for quarterly, frequency = 4. Basis is the type of day
 * count basis to use
 */
    internal fun calcCoupDaysNC(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "COUPDAYSNC")
val settlementDate = Long(operands[0].value.toString()).toLong()
val maturityDate = Long(operands[1].value.toString()).toLong()
val frequency = operands[2].intVal
var basis = 0
if (operands.size > 3)
basis = operands[3].intVal
 //		 TODO: if dates are not valid, return #VALUE! error
        if (basis < 0 || basis > 4)
return PtgErr(PtgErr.ERROR_NUM)
if (!(frequency == 1 || frequency == 2 || frequency == 4))
return PtgErr(PtgErr.ERROR_NUM)
if (settlementDate > maturityDate)
return PtgErr(PtgErr.ERROR_NUM)

val ncd = PtgCalculator.getLongValue(calcCoupNCD(operands))

val result = getDaysFromBasis(basis, settlementDate, ncd).toDouble()
if (DEBUG)
Logger.logInfo("Result from calcCoupDaysNC= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * COUPNCD Returns the next coupon date after the settlement date Returns a
 * number that represents the next coupon date after the settlement date.
 * COUPNCD(settlement,maturity,frequency,basis) Important Dates should be
 * entered by using the DATE function, or as results of other formulas or
 * functions. For example, use DATE(2008,5,23) for the 23rd day of May,
 * 2008. Problems can occur if dates are entered as text. Settlement is the
 * security's settlement date. The security settlement date is the date
 * after the issue date when the security is traded to the buyer. Maturity
 * is the security's maturity date. The maturity date is the date when the
 * security expires. Frequency is the number of coupon payments per year.
 * For annual payments, frequency = 1; for semiannual, frequency = 2; for
 * quarterly, frequency = 4. Basis is the type of day count basis to use.
 */
    internal fun calcCoupNCD(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "COUPNCD")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val frequency = operands[2].intVal
var basis = 0
if (operands.size > 3)
basis = operands[3].intVal
 //		 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4)
return PtgErr(PtgErr.ERROR_NUM)
if (!(frequency == 1 || frequency == 2 || frequency == 4))
return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate))
return PtgErr(PtgErr.ERROR_NUM)

val resultDate:GregorianCalendar
var mm = mDate.get(Calendar.MONTH) + 1 // months are 0-based but calc
 // needs 1-based - for now!
            val sm = sDate.get(Calendar.MONTH) + 1
var y = mDate.get(Calendar.YEAR)
val d = mDate.get(Calendar.DAY_OF_MONTH)

if (frequency == 1)
{ // annual
while (sDate.before((GregorianCalendar(y, mm - 1, d))))
{
y--
}
y++
}
if (frequency == 2)
{ // semi-annual
while (sDate.before((GregorianCalendar(y, mm - 1, d))))
{
mm -= 6
if (mm < 1)
{
mm += 12
y--
}
}
mm += 6
}
else if (frequency == 4)
{ // quarterly
while (sDate.before((GregorianCalendar(y, mm - 1, d))))
{
mm -= 3
if (mm < 1)
{
mm += 12
y--
}
}
mm += 3
}
resultDate = GregorianCalendar(y, mm - 1, d)
val date = DateConverter.getXLSDateVal(resultDate)
if (DEBUG)
Logger.logInfo(("Result from calcCoupNCD= "
+ date
+ " "
+ java.text.DateFormat.getDateInstance().format(
resultDate.getTime())))
val i = date.toInt()
val pi = PtgInt(i)
return pi
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * COUPNUM Returns the number of coupons payable between the settlement date
 * and maturity date Settlement is the security's settlement date. The
 * security settlement date is the date after the issue date when the
 * security is traded to the buyer Maturity is the security's maturity date.
 * The maturity date is the date when the security expires. Frequency is the
 * number of coupon payments per year. For annual payments, frequency = 1;
 * for semiannual, frequency = 2; for quarterly, frequency = 4. Basis is the
 * type of day count basis to use
 */
    internal fun calcCoupNum(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "COUPNUM")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()
val frequency = operands[2].intVal
var basis = 0
if (operands.size > 3)
basis = operands[3].intVal
 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4)
return PtgErr(PtgErr.ERROR_NUM)
if (!(frequency == 1 || frequency == 2 || frequency == 4))
return PtgErr(PtgErr.ERROR_NUM)
if (settlementDate > maturityDate)
return PtgErr(PtgErr.ERROR_NUM)
 /*
		double result = getDaysFromBasis(basis, settlementDate, maturityDate)
				/ (getDaysInYearFromBasis(basis, settlementDate, maturityDate) / frequency);
		result = Math.ceil(result);
		double delta= maturityDate-settlementDate;
		double result= delta/calcCoupDays(operands));
*/
            val result = Math.ceil(yearFrac(basis, settlementDate, maturityDate) * frequency)
if (DEBUG)
Logger.logInfo("Result from calcCoupNUM= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * COUPPCD Returns the previous coupon date before the settlement date
 */
    internal fun calcCoupPCD(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG) debugOperands(operands, "COUPPCD")
try
{
val sDate = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val frequency = operands[2].intVal
var basis = 0
if (operands.size > 3)
basis = operands[3].intVal
 //		 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (!(frequency == 1 || frequency == 2 || frequency == 4))
return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate)) return PtgErr(PtgErr.ERROR_NUM)

val resultDate:GregorianCalendar
var mm = mDate.get(Calendar.MONTH) + 1 // months are 0-based but calc
 // needs 1-based - for now!
            val sm = sDate.get(Calendar.MONTH) + 1
var y = mDate.get(Calendar.YEAR)
val d = mDate.get(Calendar.DAY_OF_MONTH)

if (frequency == 1)
{ // annual
while (sDate.before((GregorianCalendar(y, mm - 1, d))))
{
y--
}
}
if (frequency == 2)
{ // semi-annual
while (sDate.before((GregorianCalendar(y, mm - 1, d))))
{
mm -= 6
if (mm < 1)
{
mm += 12
y--
}
}
}
else if (frequency == 4)
{ // quarterly
while (sDate.before((GregorianCalendar(y, mm - 1, d))))
{
mm -= 3
if (mm < 1)
{
mm += 12
y--
}
}
}
resultDate = GregorianCalendar(y, mm - 1, d)
val date = DateConverter.getXLSDateVal(resultDate)
if (DEBUG)
Logger.logInfo(("Result from calcCoupPCD= "
+ date
+ " "
+ java.text.DateFormat.getDateInstance().format(
resultDate.getTime())))
val i = date.toInt()
val pi = PtgInt(i)
return pi
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * CUMIPMT Returns the cumulative interest paid between two periods
 * CUMIPMT(rate,nper,pv,start_period,end_period,type) All parameters are
 * require Rate is the interest rate. Nper is the total number of payment
 * periods. Pv is the present value. Start_period is the first period in the
 * calculation. Payment periods are numbered beginning with 1. End_period is
 * the last period in the calculation. Type is the timing of the payment.
 *
 *
 * Type Timing 0 (zero) Payment at the end of the period 1 Payment at the
 * beginning of the period
 */
    internal fun calcCumIPmt(operands:Array<Ptg>):Ptg {
if (operands.size < 6)
{
return PtgErr(PtgErr.ERROR_NULL)
}
val rate = operands[0].doubleVal
val nper = operands[1].doubleVal
val pv = operands[2].doubleVal
val startperiod = operands[3].intVal
val endperiod = operands[4].intVal
val type = operands[5].intVal

if (DEBUG)
debugOperands(operands, "CUMIPMT")
if (rate <= 0 || pv <= 0 || nper <= 0)
return PtgErr(PtgErr.ERROR_NUM)
if (startperiod < 1 || endperiod < 1)
return PtgErr(PtgErr.ERROR_NUM)
if (startperiod > endperiod)
return PtgErr(PtgErr.ERROR_NUM)
if (type < 0 || type > 1)
return PtgErr(PtgErr.ERROR_NUM)

 // CumIPMT= pmt*period - FV for start-1 - pmt - FV for end period and
        // pmt
        val A:Double
val B:Double
 //	PMT used in fv calc
        val Rn = Math.pow(1 + rate, nper)
A = -pv * Rn * rate
B = (Rn - 1) * (1 + rate * type)
val pmt = A / B

 // WORKS on everything BUT type=1 AND startperiod=1 !!!!!!
        val n = (startperiod - 1 - type).toDouble()
val period = endperiod - startperiod + 1

 // FVa (StartPeriod)
        A = Math.pow(1 + rate, n)
B = pmt * (1 + rate * type)
B *= (Math.pow(1 + rate, n) - 1) / rate
val fva = -(pv * A + B)
 // FVb (endPeriod)
        A = Math.pow(1 + rate, (endperiod - type).toDouble())
B = pmt * (1 + rate * type)
B *= (Math.pow(1 + rate, (endperiod - type).toDouble()) - 1) / rate
val fvb = -(pv * A + B)

var result = fva - fvb - (pmt * period) //- (fva - fvb);
if (startperiod == 1 && type == 1)
result = (pmt * period) + pv // I'm sure there's a good reason for
 // this!?!?!
        result *= -1.0
if (DEBUG)
Logger.logInfo("Result from calcCumIPmt= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * CUMPRINC Returns the cumulative principal paid on a loan between
 * start_period and end_period. Syntax
 *
 *
 * CUMPRINC(rate,nper,pv,start_period,end_period,type)
 *
 *
 * Rate is the interest rate. Nper is the total number of payment periods.
 * Pv is the present value. Start_period is the first period in the
 * calculation. Payment periods are numbered beginning with 1. End_period is
 * the last period in the calculation. Type is the timing of the payment.
 * Type Timing 0 (zero) Payment at the end of the period 1 Payment at the
 * beginning of the period
 */
    internal fun calcCumPrinc(operands:Array<Ptg>):Ptg {
if (operands.size < 6)
{
return PtgErr(PtgErr.ERROR_NULL)
}
val rate = operands[0].doubleVal
val nper = operands[1].doubleVal
val pv = operands[2].doubleVal
val startperiod = operands[3].intVal
val endperiod = operands[4].intVal
val type = operands[5].intVal

if (DEBUG)
debugOperands(operands, "CUMPRINC")
 // Cumprinc= FV for start-1 and pmt - FV for end period and pmt
        val A:Double
val B:Double
 //	PMT used in fv calc
        val Rn = Math.pow(1 + rate, nper)
A = -pv * Rn * rate
B = (Rn - 1) * (1 + rate * type)
val pmt = A / B

 // FVa (StartPeriod)
        A = Math.pow(1 + rate, (startperiod - type - 1).toDouble())
B = pmt * (1 + rate * type)
B *= (Math.pow(1 + rate, (startperiod - type - 1).toDouble()) - 1) / rate
val fva = -(pv * A + B)
 // FVb (endPeriod)
        A = Math.pow(1 + rate, (endperiod - type).toDouble())
B = pmt * (1 + rate * type)
B *= (Math.pow(1 + rate, (endperiod - type).toDouble()) - 1) / rate
val fvb = -(pv * A + B)

var result = fva - fvb
if (startperiod == 1 && type == 1)
result = pv // I'm sure there's a good reason for this!?!?!
if (DEBUG)
Logger.logInfo("Result from calcCUMPRINC= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * DB Returns the depreciation of an asset for a specified period using the
 * fixed-declining balance method
 *
 *
 * DB(cost,salvage,life,period,month)
 *
 *
 * Cost is the initial cost of the asset.
 *
 *
 * Salvage is the value at the end of the depreciation (sometimes called the
 * salvage value of the asset).
 *
 *
 * Life is the number of periods over which the asset is being depreciated
 * (sometimes called the useful life of the asset).
 *
 *
 * Period is the period for which you want to calculate the depreciation.
 * Period must use the same units as life.
 *
 *
 * Month is the number of months in the first year. If month is omitted, it
 * is assumed to be 12.
 */

    internal fun calcDB(operands:Array<Ptg>):Ptg {
if (operands.size < 4 || operands[0].components != null)
{ // not
 // supported
            // by
            // function
            val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
val cost:Double
val salvage:Double
val life:Int
val period:Int
val month:Int
cost = Double((operands[0].value).toString()).toDouble()
salvage = Double((operands[1].value).toString())
.toDouble()
life = Integer.valueOf((operands[2].value).toString()).toInt()
period = Integer.valueOf((operands[3].value).toString()).toInt()
if (operands.size > 4)
{
if (operands[4] is PtgMissArg)
{
month = 12
}
else
{
month = Integer.valueOf((operands[4].value).toString())
.toInt()
}
}
else
{
month = 12
}
val salCost = salvage / cost
 // this section longhand due to some wierd calcs when calling lifdiv =
        // 1/life;
        var lifdiv = 1.0
lifdiv /= life.toDouble()
var rate = Math.pow(salCost, lifdiv)
rate = 1 - rate
rate = rate * 1000
rate = Math.round(rate).toDouble()
rate /= 1000.0

var totalDepreciation = cost * rate * month.toDouble() / 12
var result = totalDepreciation
 // 1st and last (i.e. period==life) are special cases
        var i = 2
while (i < period || (i == period && period <= life))
{
result = (cost - totalDepreciation) * rate
totalDepreciation += (cost - totalDepreciation) * rate
i++
}
if (period > life)
 // last depreciation is special calc
            result = (cost - totalDepreciation) * rate * (12 - month).toDouble() / 12
val pnum = PtgNumber(result)
return pnum
}

/**
 * DDB Returns the depreciation of an asset for a spcified period using the
 * double-declining balance method or some other method you specify
 *
 *
 * DDB(cost,salvage,life,period,factor)
 *
 *
 * Cost is the initial cost of the asset.
 *
 *
 * Salvage is the value at the end of the depreciation (sometimes called the
 * salvage value of the asset).
 *
 *
 * Life is the number of periods over which the asset is being depreciated
 * (sometimes called the useful life of the asset).
 *
 *
 * Period is the period for which you want to calculate the depreciation.
 * Period must use the same units as life.
 *
 *
 * Factor is the rate at which the balance declines. If factor is omitted,
 * it is assumed to be 2 (the double-declining balance method).
 *
 *
 * All five arguments must be positive numbers.
 */
    internal fun calcDDB(operands:Array<Ptg>):Ptg {
if (operands.size < 4 || operands[0].components != null)
{ // not
 // supported
            // by
            // function
            val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
val cost:Double
val salvage:Double
val life:Int
val period:Int
val factor:Int
cost = Double((operands[0].value).toString()).toDouble()
salvage = Double((operands[1].value).toString())
.toDouble()
life = Integer.valueOf((operands[2].value).toString()).toInt()
period = Integer.valueOf((operands[3].value).toString()).toInt()
factor = 2
if (operands.size > 4)
{
if (!(operands[4] is PtgMissArg))
factor = Integer.valueOf((operands[4].value).toString())
.toInt()
}
val salCost = salvage / cost
 // this section longhand due to some wierd calcs when calling lifdiv =
        // 1/life;
        val facLife = factor / (life.toDouble())
var totalDepreciation = 0.0
 /*	original calc
		for (int i = 1; i < period; i++) {
			//((cost-salvage) - total depreciation from prior periods) * (factor/life)
			totalDepreciation += (cost - salvage - totalDepreciation)*facLife;
		}
//		double result= (cost - salvage - totalDepreciation) * (facLife);
 *
 */
        for (i in 1 until period)
{
totalDepreciation += (cost - totalDepreciation) * facLife
}
var result = 0.0
if (cost - salvage - totalDepreciation > 0)
result = (cost - totalDepreciation) * facLife
val pnum = PtgNumber(result)
return pnum
}

/**
 * DISC Returns the discount rate for a security
 *
 *
 * Settlement is the security's settlement date. The security settlement
 * date is the date after the issue date when the security is traded to the
 * buyer. Maturity is the security's maturity date. The maturity date is the
 * date when the security expires. Pr is the security's price per $100 face
 * value. Redemption is the security's redemption value per $100 face value.
 * Basis (optional) is the type of day count basis to use.
 */
    internal fun calcDISC(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcDISC")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()
val pr = operands[2].doubleVal
val redemption = operands[3].doubleVal
var basis = 0
if (operands.size > 4)
basis = operands[4].intVal
 //		 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4)
return PtgErr(PtgErr.ERROR_NUM)
if (settlementDate > maturityDate)
return PtgErr(PtgErr.ERROR_NUM)
if (pr <= 0 || redemption <= 0)
return PtgErr(PtgErr.ERROR_NUM)


 //		double DSM = maturityDate - settlementDate;
 //		double B = getDaysInYearFromBasis(basis, settlementDate, maturityDate);
 //		double result = (redemption - pr) / redemption * (B / DSM);
            val result = (redemption - pr) / (redemption * yearFrac(basis, settlementDate, maturityDate))

val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcDISC= " + result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * DOLLARDE Converts a dollar price, expressed as a fraction, into a dollar
 * price, expressed as a decimal number Fractional_dollar is a number
 * expressed as a fraction. Fraction is the integer to use in the
 * denominator of the fraction.
 */
    internal fun calcDollarDE(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcDOLLARDE")
val fractional_dollar = operands[0].doubleVal
val fraction = operands[1].intVal
if (fraction < 0)
return PtgErr(PtgErr.ERROR_NUM)
if (fraction == 0)
return PtgErr(PtgErr.ERROR_DIV_ZERO)

val n = (fraction).toString().length
val x = Math.floor(fractional_dollar)
val y = (fractional_dollar - x) * Math.pow(10.0, n.toDouble())
val result = x + y / fraction
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcDOLLARDE= " + result)
return pnum
}

/**
 * DOLLARFR Converts a dollar price, expressed as a decimal number, into a
 * dollar price, expressed as a fraction Decimal_dollar is a decimal number.
 * Fraction is the integer to use in the denominator of a fraction
 */
    internal fun calcDollarFR(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcDOLLARFR")
val decimal_dollar = operands[0].doubleVal
val fraction = operands[1].intVal
if (fraction < 0)
return PtgErr(PtgErr.ERROR_NUM)
if (fraction == 0)
return PtgErr(PtgErr.ERROR_DIV_ZERO)

val n = (fraction).toString().length

val x = Math.floor(decimal_dollar)
val y = (decimal_dollar - x)
val result = x + (y * fraction) / Math.pow(10.0, n.toDouble())
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcDOLLARFR= " + result)
return pnum
}

/**
 * DURATION Returns the annual duration of a security with periodic interest
 * payments
 * DURATION(settlement,maturity,coupon,yld,frequency,basis)
 * Settlement is the security's settlement date. Maturity is the security's
 * maturity date. Coupon is the security's annual coupon rate. Yld is the
 * security's annual yield. Frequency is the number of coupon payments per
 * year. For annual payments, frequency = 1; for semiannual, frequency = 2;
 * for quarterly, frequency = 4. Basis is the type of day count basis to use
 */
    internal fun calcDURATION(operands:Array<Ptg>):Ptg {
if (operands.size < 5)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcDURATION")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val coupon = operands[2].doubleVal
val yld = operands[3].doubleVal
val frequency = operands[4].intVal
var basis = 0
if (operands.size > 5)
basis = operands[5].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4)
return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate))
return PtgErr(PtgErr.ERROR_NUM)
if (coupon < 0 || yld < 0)
return PtgErr(PtgErr.ERROR_NUM)
if (!(frequency == 1 || frequency == 2 || frequency == 4))
return PtgErr(PtgErr.ERROR_NUM)

val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()

val ops = arrayOfNulls<Ptg>(4)
ops[0] = operands[0]
ops[1] = operands[1]
ops[2] = operands[4]
ops[3] = PtgInt(basis)

 /*
             * Duration= (A + SumC)/(D+ SumB * 1/frequency
             *
             * where
             * Y= 1+ yield/frequency
             * R= 100*rate
             * F= DSC/E
             * A= (F*100)/Y^(n-1+F)
             * SumC= Sum(1,n): R/(frequency*Y^(i-1+F)) * (i-1+F)
             * D= 100/Y^(n-1+F)
             * SumB= Sum(1,n): R/(frequency*Y^(i-1+F))
             *
             */
            val n = calcCoupNum(ops).doubleVal
val DSC = calcCoupDaysNC(ops).doubleVal
val E = calcCoupDays(ops).doubleVal
val F = DSC / E
val R = coupon * 100
val Y = 1 + yld / frequency
val Yx = Math.pow(Y, n - 1 + F)
var SumA = 0.0
run({ var i = 1
while (i <= n)
{
SumA += R * (i - 1 + F) / (Math.pow(Y, i - 1 + F) * frequency)
i++
} })
var SumB = 0.0
var i = 1
while (i <= n)
{
SumB += R / (Math.pow(Y, i - 1 + F) * frequency)
i++
}
var C = 0.0
var D = 0.0
if (n > 1)
{
C = ((n - 1 + F) * 100) / Yx
D = 100 / Yx
}
val result = (SumA + C) / ((SumB + D) * frequency)
val pnum = PtgNumber(result)
if (DEBUG) Logger.logInfo("Result from calcDURATION= " + result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * EFFECT Returns the effective annual interest rate Nominal_rate is the
 * nominal interest rate. Npery is the number of compounding periods per
 * year
 */
    internal fun calcEffect(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcEFFECT")
val nominal_rate = operands[0].doubleVal
val npery = operands[1].intVal
 // TODO: If either argument is non-numeric, #VALUE! error
        if (npery <= 0 || npery < 1)
 // funny guard!!!!
            return PtgErr(PtgErr.ERROR_NUM)
 /*
         * // KSC: TESTING for (int j= 1; j <= 10; j++) { double i; if ((j % 2) ==
         * 0) i= 1; else i= .5; npery= j; while (i <= 10) { nominal_rate= i /
         * 100;
         *
         *
         * double x= nominal_rate / npery; double result= Math.pow(1 + x, npery) -
         * 1; Logger.logInfo("nominal_rate= " + nominal_rate + " npery= " +
         * npery + " result= " + result); java.math.BigDecimal bd= new
         * java.math.BigDecimal(i-Math.floor(i)).setScale(5,
         * java.math.BigDecimal.ROUND_HALF_UP); if (((bd.doubleValue()*100) % 4) ==
         * 0) // if ((((i-Math.floor(i))*100) % 4)==0) i+= 0.3; else i+= 0.25; } }
         */
        val x = nominal_rate / npery
val result = Math.pow(1 + x, npery.toDouble()) - 1
val pnum = PtgNumber(result)
if (DEBUG) Logger.logInfo("Result from calcEFFECT= " + result)
return pnum
}

/**
 * FV Returns the future value of an investment FV(rate,nper,pmt,pv,type)
 *
 *
 *
 *
 * Rate is the interest rate per period. Nper is the total number of payment
 * periods in an annuity. Pmt is the payment made each period; it cannot
 * change over the life of the annuity . Typically, pmt contains principal
 * and interest but no other fees or taxes. If pmt is omitted, you must
 * include the pv argument. Pv is the present value, or the lump-sum amount
 * that a series of future payments is worth right now. If pv is omitted, it
 * is assumed to be 0 (zero), and you must include the pmt argument. Type is
 * the number 0 or 1 and indicates when payments are due. If type is
 * omitted, it is assumed to be 0.
 */
    internal fun calcFV(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcFV")
val rate = operands[0].doubleVal
val nper = operands[1].intVal
val pmt = operands[2].doubleVal
var pv = 0.0
var type = 0
if (operands.size > 3 && !(operands[3] is PtgMissArg))
pv = operands[3].doubleVal
if (operands.size > 4)
type = operands[4].intVal

val A = Math.pow(1 + rate, nper.toDouble())
var B = pmt * (1 + rate * type)
B *= (Math.pow(1 + rate, nper.toDouble()) - 1) / rate
val result = -(pv * A + B)
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcFV= " + result)
return pnum

}

/**
 * FVSCHEDULE Returns the future value of an initial principal after
 * applying a series of compound interest rates
 */
    internal fun calcFVSCHEDULE(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG) debugOperands(operands, "calcFVSCHEDULE")
val principal = operands[0].doubleVal
val schedule = PtgCalculator.getAllComponents(operands[1])
if (DEBUG) debugOperands(schedule, "calcFVSCHEDULE") // AFTER converting
 // references ...
        var result = 1.0
for (i in schedule.indices)
{
result *= principal + schedule[i].doubleVal
}
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcFVSCHEDULE= " + result)
return pnum
}

/**
 * INTRATE Returns the interest rate for a fully invested security
 */
    internal fun calcINTRATE(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcINTRATE")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val investment = operands[2].doubleVal
val redemption = operands[3].doubleVal
var basis = 0
if (operands.size > 4)
basis = operands[4].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate)) return PtgErr(PtgErr.ERROR_NUM)
if (investment <= 0 || redemption <= 0) return PtgErr(PtgErr.ERROR_NUM)

val settlementDate = (DateConverter.getXLSDateVal(sDate)).toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate)).toLong()
 //		double delta = maturityDate - settlementDate;
 //		double result = ((redemption - investment) / investment) * ((getDaysInYearFromBasis(basis, settlementDate, maturityDate) / delta));
            val result = ((redemption - investment) / investment) / yearFrac(basis, settlementDate, maturityDate)
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcINTRATE= " + result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * IPMT Returns the interest payment for an investment for a given period.
 *
 *
 * Syntax IPMT(rate,per,nper,pv,fv,type)
 *
 *
 * Rate is the interest rate per period. Per is the period for which you
 * want to find the interest and must be in the range 1 to nper. Nper is the
 * total number of payment periods in an annuity. Pv is the present value,
 * or the lump-sum amount that a series of future payments is worth right
 * now. Fv is the future value, or a cash balance you want to attain after
 * the last payment is made. If fv is omitted, it is assumed to be 0 (the
 * future value of a loan, for example, is 0). Type is the number 0 or 1 and
 * indicates when payments are due. If type is omitted, it is assumed to be
 * 0.
 */
    internal fun calcIPMT(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcIPMT")
val rate = operands[0].doubleVal
val per = operands[1].doubleVal
val nper = operands[2].doubleVal
val pv = operands[3].doubleVal
var fv = 0.0
var type = 0
if (operands.size > 4 && !(operands[4] is PtgMissArg))
fv = operands[4].doubleVal
if (operands.size > 5)
type = operands[5].intVal

if (per < 0 || per > nper)
return PtgErr(PtgErr.ERROR_NUM)

 // IPMT= pmt- (fv(b) - fv(a) i.e. = payement less the principal balance
        // btwn two periods

        val n:Double
if (type == 0)
n = per
else
n = per - 1
 // PMT
        val Rn = Math.pow(1 + rate, nper)
var A = -fv * rate - pv * Rn * rate
var B = (Rn - 1) * (1 + rate * type)
val pmt = A / B
 // FVa
        A = Math.pow(1 + rate, n)
B = pmt * (1 + rate * type)
B *= (Math.pow(1 + rate, n) - 1) / rate
val fva = -(pv * A + B)
 // FVb
        A = Math.pow(1 + rate, n - 1)
B = pmt * (1 + rate * type)
B *= (Math.pow(1 + rate, n - 1) - 1) / rate
val fvb = -(pv * A + B)

val result = pmt - (fvb - fva)
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcIPMT= " + result)
return pnum
}

/**
 * IRR Returns the internal rate of return for a series of cash flows
 * Returns the internal rate of return for a series of cash flows
 * represented by the numbers in values. These cash flows do not have to be
 * even, as they would be for an annuity. However, the cash flows must occur
 * at regular intervals, such as monthly or annually. The internal rate of
 * return is the interest rate received for an investment consisting of
 * payments (negative values) and income (positive values) that occur at
 * regular periods.
 *
 *
 * SEE: NPV
 *
 *
 * Syntax IRR(values,guess)
 *
 *
 * Values is an array or a reference to cells that contain numbers for which
 * you want to calculate the internal rate of return.
 *
 *
 * Values must contain at least one positive value and one negative value to
 * calculate the internal rate of return.
 *
 *
 * IRR uses the order of values to interpret the order of cash flows. Be
 * sure to enter your payment and income values in the sequence you want.
 *
 *
 * If an array or reference argument contains text, logical values, or empty
 * cells, those values are ignored. Guess is a number that you guess is
 * close to the result of IRR.
 *
 *
 * Microsoft Excel uses an iterative technique for calculating IRR. Starting
 * with guess, IRR cycles through the calculation until the result is
 * accurate within 0.00001 percent. If IRR can't find a result that works
 * after 20 tries, the #NUM! error value is returned.
 *
 *
 * In most cases you do not need to provide guess for the IRR calculation.
 * If guess is omitted, it is assumed to be 0.1 (10 percent).
 *
 *
 * If IRR gives the #NUM! error value, or if the result is not close to what
 * you expected, try again with a different value for guess
 */
    internal fun calcIRR(operands:Array<Ptg>):Ptg {
if (operands.size < 1 || operands[0].components == null)
{ // not
 // supported
            // by
            // function
            val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcIRR")
var guess = .1
if (operands.size > 1)
guess = operands[1].doubleVal
val params = PtgCalculator.getAllComponents(operands)
if (DEBUG)
debugOperands(params, "calcIRR") // AFTER converting references ...
val n = params.size

 // examine values array, sum all outflows (= negative values) + inflows
        // (= positive values)
        var outflow = 0.0
var inflow = 0.0
 // get outflow (- values)
        for (i in 0 until n)
{
val `val` = params[i].doubleVal
if (`val` < 0)
outflow += Math.abs(`val`)
else
inflow += `val`
}
if (outflow == 0.0 || inflow == 0.0)
return PtgErr(PtgErr.ERROR_VALUE) // TODO: Excel doesn't specify which
 // error to return

        // iterate over possible irr values; value is correct when
        // outflow-pv <= tolerance, defined as .00001%
        val TOLERANCE = 0.0000001
var bIsCorrect = false
val xl:Double
val xh:Double
val fl:Double
val fh:Double
val f:Double
var trial = guess
xl = 0.0
xh = guess
var delta = xh - xl
fl = outflow - inflow
fh = outflow
for (i in 0 until n)
{
val `val` = params[i].doubleVal
if (`val` > 0)
{
fh -= `val` / (Math.pow(1 + xh, i.toDouble()))
}
}
var j = 0
while (j < 50 && !bIsCorrect)
{ // maximum 20 tries - need
 // more!!!!
            trial = xl + delta * fl / (fl - fh)
f = outflow
for (i in 0 until n)
{
val `val` = params[i].doubleVal
if (`val` > 0)
{
f -= `val` / (Math.pow(1 + trial, i.toDouble()))
}
}

if (f < 0)
{
delta = xl - trial
xl = trial
fl = f
}
else
{
delta = xh - trial
xh = trial
fh = f
}
bIsCorrect = (Math.abs(delta) <= TOLERANCE)
delta = xh - xl
j++
}
if (!bIsCorrect)
return PtgErr(PtgErr.ERROR_NUM)

if (DEBUG)
Logger.logInfo("Result from calcIRR= " + trial)
val pnum = PtgNumber(trial)
return pnum
}

/**
 * ISPMT Calculates the interest paid during a specific period of an
 * investment. This function is provided for compatibility with Lotus 1-2-3.
 * ISPMT(rate,per,nper,pv) Rate is the interest rate for the investment. Per
 * is the period for which you want to find the interest, and must be
 * between 1 and nper. Nper is the total number of payment periods for the
 * investment. Pv is the present value of the investment. For a loan, pv is
 * the loan amount.
 */
    internal fun calcISPMT(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcISPMT")
val rate = operands[0].doubleVal
val per = operands[1].doubleVal
val nper = operands[2].doubleVal
val pv = operands[3].doubleVal
if (per < 0 || per > nper)
return PtgErr(PtgErr.ERROR_NUM)

val result = (-pv * rate * (nper - per)) / nper
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcISPMT= " + result)
return pnum
}

/**
 * MDURATION Returns the Macauley modified duration for a security with an
 * assumed par value of $100
 * DURATION(settlement,maturity,coupon,yld,frequency,basis) Settlement is
 * the security's settlement date. Maturity is the security's maturity date.
 * Coupon is the security's annual coupon rate. Yld is the security's annual
 * yield. Frequency is the number of coupon payments per year. For annual
 * payments, frequency = 1; for semiannual, frequency = 2; for quarterly,
 * frequency = 4. Basis is the type of day count basis to use
 */
    internal fun calcMDURATION(operands:Array<Ptg>):Ptg {
if (operands.size < 5)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcMDURATION")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val coupon = operands[2].doubleVal
val yld = operands[3].doubleVal
val frequency = operands[4].intVal
var basis = 0
if (operands.size > 5)
basis = operands[5].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4)
return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate))
return PtgErr(PtgErr.ERROR_NUM)
if (coupon < 0 || yld < 0)
return PtgErr(PtgErr.ERROR_NUM)
if (!(frequency == 1 || frequency == 2 || frequency == 4))
return PtgErr(PtgErr.ERROR_NUM)

var result = calcDURATION(operands).doubleVal
 // above is regular duration calculation; to get modified duration:
            result = result / (1 + yld / frequency)
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcMDURATION= " + result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * MIRR Returns the internal rate of return where positive and negative cash
 * flows are financed at different rates
 *
 *
 * MIRR(values,finance_rate,reinvest_rate)
 *
 *
 * Values is an array or a reference to cells that contain numbers. These
 * numbers represent a series of payments (negative values) and income
 * (positive values) occurring at regular periods.
 *
 *
 * Values must contain at least one positive value and one negative value to
 * calculate the modified internal rate of return. Otherwise, MIRR returns
 * the #DIV/0! error value.
 *
 *
 * If an array or reference argument contains text, logical values, or empty
 * cells, those values are ignored; however, cells with the value zero are
 * included.
 *
 *
 * Finance_rate is the interest rate you pay on the money used in the cash
 * flows.
 *
 *
 * Reinvest_rate is the interest rate you receive on the cash flows as you
 * reinvest them.
 */
    internal fun calcMIRR(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcMIRR")
val finance_rate = operands[1].doubleVal
val reinvest_rate = operands[2].doubleVal
val params = PtgCalculator.getAllComponents(operands)
if (DEBUG)
debugOperands(params, "calcMIRR") // AFTER converting references
 // ...

        // Get + Values and - Values in separate Ptg arrays
        val posVals = Vector()
val negVals = Vector()
val positiveValues:Array<Ptg>
val negativeValues:Array<Ptg>
val n = params.size - 2 // skip last 2 params (= rates)
for (i in 0 until n)
{
val `val` = params[i].doubleVal
if (`val` < 0)
negVals.addElement(params[i])
else
posVals.addElement(params[i])
}
positiveValues = arrayOfNulls<Ptg>(posVals.size + 1)
negativeValues = arrayOfNulls<Ptg>(negVals.size + 1)
System.arraycopy(posVals.toTypedArray(), 0, positiveValues, 1, posVals
.size)
System.arraycopy(negVals.toTypedArray(), 0, negativeValues, 1, negVals
.size)

 // add rate to Ptg array for call to calcNPV
        positiveValues[0] = operands[2] // reinvest rate
negativeValues[0] = operands[1] // finance rate

 // Calculate MIRR from NPV values
        var X = calcNPV(positiveValues).doubleVal
X = -1.0 * X * Math.pow(1 + reinvest_rate, posVals.size.toDouble())
var Y = calcNPV(negativeValues).doubleVal
Y = Y * (1 + finance_rate)
val result = Math.pow(X / Y, 1.0 / (n - 1)) - 1

if (DEBUG)
Logger.logInfo("Result from calcMIRR= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * NOMINAL Returns the annual nominal interest rate
 */
    internal fun calcNominal(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcNominal")
val effect = operands[0].doubleVal
val npery = operands[1].intVal
 // TODO: if either is non-numeric, return #VALUE!
        if (effect <= 0 || npery < 1)
return PtgErr(PtgErr.ERROR_NUM)

 // solve for nominal_rate in:
        // effect= (1 + nominal_rate/npery)^npery -1
        // nominal_rate= ((10^(log10(y)/n)) - 1)*npery
        val y = effect + 1
val log10y = Math.log(y) / Math.log(10.0) // base 10 log
val z = Math.pow(10.0, log10y / npery)
val result = (z - 1) * npery
if (DEBUG)
Logger.logInfo("Result from calcNominal= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * NPER Returns the number of periods for an investment NPER(rate, pmt, pv,
 * fv, type)
 *
 *
 * Rate is the interest rate per period. Pmt is the payment made each
 * period; it cannot change over the life of the annuity. Typically, pmt
 * contains principal and interest but no other fees or taxes. Pv is the
 * present value, or the lump-sum amount that a series of future payments is
 * worth right now. Fv is the future value, or a cash balance you want to
 * attain after the last payment is made. If fv is omitted, it is assumed to
 * be 0 (the future value of a loan, for example, is 0). Type is the number
 * 0 or 1 and indicates when payments are due.
 */
    internal fun calcNPER(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcNPER")
val rate = operands[0].doubleVal
val pmt = operands[1].doubleVal
val pv = operands[2].doubleVal
var fv = 0.0
if (operands.size > 3 && !(operands[3] is PtgMissArg))
fv = operands[3].doubleVal
var type = 0
if (operands.size > 4)
type = operands[4].intVal

val A = pmt * (1 + type * rate) - rate * fv
val B = pmt * (1 + type * rate) + rate * pv
val C = 1 + rate
val result = Math.log(A / B) / Math.log(C)
if (DEBUG)
Logger.logInfo("Result from calcNPER= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * NPV(rate,value1,value2, ...) Calculates the net present value of an
 * investment by using a discount rate and a series of future payments
 * (negative values) and income (positive values).
 *
 *
 * Rate is the rate of discount over the length of one period. Value1,
 * value2, ... are 1 to 29 arguments representing the payments and income.
 * Value1, value2, ... must be equally spaced in time and occur at the end
 * of each period
 *
 *
 * Returns the net present value of an investment based on a series of
 * periodic cash flows and a discount rate = Sum(values / (1 + rate) )
 */
    internal fun calcNPV(operands:Array<Ptg>):Ptg {
if (operands.size < 2 || operands[0].components != null)
{ // not
 // supported
            // by
            // function
            val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG) debugOperands(operands, "calcNPV")
val params = PtgCalculator.getAllComponents(operands)
if (DEBUG) debugOperands(params, "calcNPV") // AFTER converting references ...
val rate = params[0].doubleVal
val n = Math.min(params.size, 30) // at most 29 values
var result = 0.0
for (i in 1 until n)
{
val valuei = params[i].doubleVal
 // TODO: if valuei is an error, empty, etc., ignore
            result += valuei / Math.pow(1 + rate, i.toDouble())
}
if (DEBUG) Logger.logInfo("Result from calcNPV= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * ODDFPRICE Returns the price per $100 face value of a security with an odd
 * first period
 * ODDFPRICE(settlement,maturity,issue,first_coupon,rate,yld,redemption,frequency,basis)
 */
    @Throws(FunctionNotSupportedException::class)
internal fun calcODDFPRICE(operands:Array<Ptg>):Ptg {
if (true)
{
val wn = "WARNING: this version of OpenXLS does not support the formula ODDFPRICE."
Logger.logWarn(wn)
throw FunctionNotSupportedException(wn)
}
if (operands.size < 8)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcODDFPRICE")
try
{
val sDate = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val iDate = DateConverter.getCalendarFromNumber(operands[2].value) as GregorianCalendar
val fcDate = DateConverter.getCalendarFromNumber(operands[3].value) as GregorianCalendar
val rate = operands[4].doubleVal
val yld = operands[5].doubleVal
val redemption = operands[6].doubleVal
val frequency = operands[7].intVal
var basis = 0
if (operands.size > 8)
basis = operands[8].intVal

val settlementDate = (DateConverter.getXLSDateVal(sDate)).toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate)).toLong()
val issueDate = (DateConverter.getXLSDateVal(iDate)).toLong()
val firstCouponDate = (DateConverter.getXLSDateVal(fcDate)).toLong()

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate)) return PtgErr(PtgErr.ERROR_NUM)
if (yld < 0 || rate < 0) return PtgErr(PtgErr.ERROR_NUM)
if (redemption <= 0) return PtgErr(PtgErr.ERROR_NUM)
 /*maturity > first_coupon > settlement > issue */
            if ((issueDate >= settlementDate || settlementDate >= firstCouponDate ||
firstCouponDate >= maturityDate))
return PtgErr(PtgErr.ERROR_NUM)


val ops = arrayOfNulls<Ptg>(4)
ops[0] = PtgNumber(settlementDate.toDouble())
ops[1] = PtgNumber(maturityDate.toDouble())
ops[2] = PtgInt(frequency)
ops[3] = PtgInt(basis)
val A = calcCoupDayBS(ops).doubleVal // days from beg. coupon period to settlement
val DSC = calcCoupDaysNC(ops).doubleVal  // days from  settlement to next coupon
val E = calcCoupDays(ops).doubleVal  // total # days in coupon period
val N = PtgCalculator.getLongValue(calcCoupNum(ops)).toDouble() // n is the number of coupons btwn settlement and maturity
val NCD = calcCoupNCD(ops).doubleVal  // next coupon after 1st coupon date?????
val DFC = (firstCouponDate - settlementDate).toDouble()    // # days from odd first coupon to next coupon date
val z = getDaysFromBasis(basis, settlementDate, firstCouponDate).toDouble()

val R = 100 * rate / frequency
val Y = 1 + yld / frequency

var result = 0.0
if (DFC < E)
{    // odd short first coupon
val firstTerm = redemption / Math.pow(Y, N - 1 + DSC / E)
val secondTerm = (R * DFC / E) / Math.pow(Y, DSC / E)
var thirdTerm = 0.0
var i = 2
while (i <= N)
{
thirdTerm += R / Math.pow(Y, i - 1 + DSC / E)
i++
}
val fourthTerm = R * A / E
result = firstTerm + secondTerm + thirdTerm - fourthTerm
}
else
{    // odd long first coupon

}

if (DEBUG) Logger.logInfo("Result from calcODDFPRICE= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

 /*
     * ODDFYIELD Returns the yield of a security with an odd first period
     * NOT COMPLETED!!
     */
    @Throws(FunctionNotSupportedException::class)
internal fun calcODDFYIELD(operands:Array<Ptg>):Ptg {
if (true)
{
val wn = "WARNING: this version of OpenXLS does not support the formula ODDFYIELD."
Logger.logWarn(wn)
throw FunctionNotSupportedException(wn)
}
if (operands.size < 8)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcODDFYIELD")
try
{
val sDate = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val iDate = DateConverter.getCalendarFromNumber(operands[2].value) as GregorianCalendar
val fcDate = DateConverter.getCalendarFromNumber(operands[3].value) as GregorianCalendar
val rate = operands[4].doubleVal
val yld = operands[5].doubleVal
val redemption = operands[6].doubleVal
val frequency = operands[7].intVal
var basis = 0
if (operands.size > 8)
basis = operands[8].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate)) return PtgErr(PtgErr.ERROR_NUM)
if (yld < 0 || rate < 0) return PtgErr(PtgErr.ERROR_NUM)
if (redemption <= 0) return PtgErr(PtgErr.ERROR_NUM)

val settlementDate = (DateConverter.getXLSDateVal(sDate)).toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate)).toLong()
val issueDate = (DateConverter.getXLSDateVal(iDate)).toLong()
val firstCouponDate = (DateConverter.getXLSDateVal(fcDate)).toLong()

 /*
             * A= coupDayBS E= coupDays N= coupNum DSC= coupDaysNC
             */
            val ops = arrayOfNulls<Ptg>(4)
ops[0] = PtgNumber(settlementDate.toDouble())
ops[1] = PtgNumber(maturityDate.toDouble())
ops[2] = PtgInt(frequency)
ops[3] = PtgInt(basis)
val A = calcCoupDayBS(ops).doubleVal  // days from beg. coupon period to settlement
val DFC = getDaysFromBasis(basis, firstCouponDate, A.toLong()).toDouble()    // days from 1st odd coupon to 1st coupon
val DSC = calcCoupDaysNC(ops).doubleVal  // days from  settlement to next coupon
val E = calcCoupDays(ops).doubleVal  // total # days in coupon period
val N = PtgCalculator.getLongValue(calcCoupNum(ops)).toDouble() // n is the number of coupons btwn settlement and maturity
val R = 100 * rate / frequency
val Y = 1 + yld / frequency

val result = 0.0

if (DEBUG) Logger.logInfo("Result from calcODDFYIELD= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

 /*
     * ODDLPRICE Returns the price per $100 face value of a security with an odd
     * last period
     * NOT COMPLETED!!!!
     */
    @Throws(FunctionNotSupportedException::class)
internal fun calcODDLPRICE(operands:Array<Ptg>):Ptg {
if (true)
{
val wn = "WARNING: this version of OpenXLS does not support the formula ODDLPRICE."
Logger.logWarn(wn)
throw FunctionNotSupportedException(wn)
}
if (operands.size < 8)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG) debugOperands(operands, "calcODDLPRICE")
try
{
val sDate = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val iDate = DateConverter.getCalendarFromNumber(operands[2].value) as GregorianCalendar
val fcDate = DateConverter.getCalendarFromNumber(operands[3].value) as GregorianCalendar
val rate = operands[4].doubleVal
val yld = operands[5].doubleVal
val redemption = operands[6].doubleVal
val frequency = operands[7].intVal
var basis = 0
if (operands.size > 8)
basis = operands[8].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate)) return PtgErr(PtgErr.ERROR_NUM)
if (yld < 0 || rate < 0) return PtgErr(PtgErr.ERROR_NUM)
if (redemption <= 0) return PtgErr(PtgErr.ERROR_NUM)

val settlementDate = (DateConverter.getXLSDateVal(sDate)).toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate)).toLong()
val issueDate = (DateConverter.getXLSDateVal(iDate)).toLong()
val firstCouponDate = (DateConverter.getXLSDateVal(fcDate)).toLong()

 /*
             * A= coupDayBS E= coupDays N= coupNum DSC= coupDaysNC
             */
            val ops = arrayOfNulls<Ptg>(4)
ops[0] = PtgNumber(settlementDate.toDouble())
ops[1] = PtgNumber(maturityDate.toDouble())
ops[2] = PtgInt(frequency)
ops[3] = PtgInt(basis)
val A = calcCoupDayBS(ops).doubleVal // days from beg. coupon period to settlement
val DFC = getDaysFromBasis(basis, firstCouponDate, A.toLong()).toDouble()    // days from 1st odd coupon to 1st coupon
val DSC = calcCoupDaysNC(ops).doubleVal // days from  settlement to next coupon
val E = calcCoupDays(ops).doubleVal // total # days in coupon period
val N = PtgCalculator.getLongValue(calcCoupNum(ops)).toDouble() // n is the number of coupons btwn settlement and maturity
val R = 100 * rate / frequency
val Y = 1 + yld / frequency

val result = 0.0

if (DEBUG) Logger.logInfo("Result from calcODDLPRICE= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

 /*
     * ODDLYIELD Returns the yield of a security with an odd last period
     * NOT COMPLETED!!
     */
    @Throws(FunctionNotSupportedException::class)
internal fun calcODDLYIELD(operands:Array<Ptg>):Ptg {
if (true)
{
val wn = "WARNING: this version of OpenXLS does not support the formula ODDLYIELD."
Logger.logWarn(wn)
throw FunctionNotSupportedException(wn)
}
if (operands.size < 8)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG) debugOperands(operands, "calcODDLYIELD")
try
{
val sDate = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val iDate = DateConverter.getCalendarFromNumber(operands[2].value) as GregorianCalendar
val fcDate = DateConverter.getCalendarFromNumber(operands[3].value) as GregorianCalendar
val rate = operands[4].doubleVal
val yld = operands[5].doubleVal
val redemption = operands[6].doubleVal
val frequency = operands[7].intVal
var basis = 0
if (operands.size > 8)
basis = operands[8].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate)) return PtgErr(PtgErr.ERROR_NUM)
if (yld < 0 || rate < 0) return PtgErr(PtgErr.ERROR_NUM)
if (redemption <= 0) return PtgErr(PtgErr.ERROR_NUM)

val settlementDate = (DateConverter.getXLSDateVal(sDate)).toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate)).toLong()
val issueDate = (DateConverter.getXLSDateVal(iDate)).toLong()
val firstCouponDate = (DateConverter.getXLSDateVal(fcDate)).toLong()

 /*
             * A= coupDayBS E= coupDays N= coupNum DSC= coupDaysNC
             */
            val ops = arrayOfNulls<Ptg>(4)
ops[0] = PtgNumber(settlementDate.toDouble())
ops[1] = PtgNumber(maturityDate.toDouble())
ops[2] = PtgInt(frequency)
ops[3] = PtgInt(basis)
val A = calcCoupDayBS(ops).doubleVal // days from beg. coupon period to settlement
val DFC = getDaysFromBasis(basis, firstCouponDate, A.toLong()).toDouble()    // days from 1st odd coupon to 1st coupon
val DSC = calcCoupDaysNC(ops).doubleVal // days from  settlement to next coupon
val E = calcCoupDays(ops).doubleVal // total # days in coupon period
val N = PtgCalculator.getLongValue(calcCoupNum(ops)).toDouble() // n is the number of coupons btwn settlement and maturity
val R = 100 * rate / frequency
val Y = 1 + yld / frequency

val result = 0.0

if (DEBUG) Logger.logInfo("Result from calcODDLYIELD= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * PMT Returns the periodic payment for an annuity
 *
 *
 * pmt= - (fv + (1+rate)**nper * pv) * rate / ( ( (1+rate)**nper - 1 ) *
 * (1+rate*type) )
 *
 *
 * if fv or type are omitted they should be treated as 0 values.
 */
    internal fun calcPmt(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{ // not
 // supported
            // by
            // function
            val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
val rate:Double
val nper:Double
val pv:Double
val fv:Double
val type:Double
rate = Double((operands[0].value).toString()).toDouble()
nper = Double((operands[1].value).toString()).toDouble()
pv = Double((operands[2].value).toString()).toDouble()
if (operands.size > 3)
{
if (operands[3] is PtgMissArg)
{
fv = 0.0
}
else
{
fv = Double((operands[3].value).toString())
.toDouble()
}
}
else
{
fv = 0.0
}
if (operands.size > 4)
{
if (operands[4] is PtgMissArg)
{
type = 0.0
}
else
{
type = Double((operands[4].value).toString())
.toDouble()
}
}
else
{
type = 0.0
}

 //KSC: For some strange odd weird reason, original calculation was off
        // even though this should be exactly the same thing. Go figure!
        val Rn = Math.pow(1 + rate, nper)
val A = -fv * rate - pv * Rn * rate
val B = (Rn - 1) * (1 + rate * type)
val result = A / B
val pnum = PtgNumber(result)
return pnum
}

/**
 * PPMT Returns the payment on the principal for an investment for a given
 * period PPMT(rate,per,nper,pv,fv,type) Rate is the interest rate per
 * period. Per specifies the period and must be in the range 1 to nper. Nper
 * is the total number of payment periods in an annuity. Pv is the present
 * value  the total amount that a series of future payments is worth now.
 * Fv is the future value, or a cash balance you want to attain after the
 * last payment is made. If fv is omitted, it is assumed to be 0 (zero),
 * that is, the future value of a loan is 0. Type is the number 0 or 1 and
 * indicates when payments are due
 */
    internal fun calcPPMT(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcPPMT")
val rate = operands[0].doubleVal
val per = operands[1].intVal
val nper = operands[2].intVal
val pv = operands[3].doubleVal
var fv = 0.0
var type = 0
if (operands.size > 4 && !(operands[4] is PtgMissArg))
fv = operands[4].doubleVal
if (operands.size > 5)
type = operands[5].intVal

val result:Double
 // 1st, get payment for entire period
        val Rn = Math.pow(1 + rate, nper.toDouble())
var A = -fv * rate - pv * Rn * rate
var B = (Rn - 1) * (1 + rate * type)
val pmt = A / B

val n:Double
if (type == 0)
n = per.toDouble()
else
n = (per - 1).toDouble()
 // FVa
        A = Math.pow(1 + rate, n)
B = pmt * (1 + rate * type)
B *= (Math.pow(1 + rate, n) - 1) / rate
val fva = -(pv * A + B)
 // FVb
        A = Math.pow(1 + rate, n - 1)
B = pmt * (1 + rate * type)
B *= (Math.pow(1 + rate, n - 1) - 1) / rate
val fvb = -(pv * A + B)

result = fvb - fva
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcPPMT= " + result)
return pnum

}

/**
 * PRICE Returns the price per $100 face value of a security that pays
 * periodic interest
 * PRICE(settlement,maturity,rate,yld,redemption,frequency,basis) Settlement
 * is the security's settlement date. The security settlement date is the
 * date after the issue date when the security is traded to the buyer.
 * Maturity is the security's maturity date. The maturity date is the date
 * when the security expires. Rate is the security's annual coupon rate. Yld
 * is the security's annual yield. Redemption is the security's redemption
 * value per $100 face value. Frequency is the number of coupon payments per
 * year. For annual payments, frequency = 1; for semiannual, frequency = 2;
 * for quarterly, frequency = 4. Basis is the type of day count basis to
 * use.
 */
    internal fun calcPRICE(operands:Array<Ptg>):Ptg {
if (operands.size < 6)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcPRICE")
try
{
val sDate = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val rate = operands[2].doubleVal
val yld = operands[3].doubleVal
val redemption = operands[4].doubleVal
val frequency = operands[5].intVal
var basis = 0
if (operands.size > 6)
basis = operands[6].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4)
return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate))
return PtgErr(PtgErr.ERROR_NUM)
if (yld < 0 || rate < 0)
return PtgErr(PtgErr.ERROR_NUM)
if (redemption <= 0)
return PtgErr(PtgErr.ERROR_NUM)

val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()

 /*
             * A= coupDayBS E= coupDays N= coupNum DSC= coupDaysNC
             */
            val ops = arrayOfNulls<Ptg>(4)
ops[0] = operands[0]
ops[1] = operands[1]
ops[2] = operands[5]
ops[3] = PtgInt(basis)
val DSC = calcCoupDaysNC(ops).doubleVal // days from  settlement to next coupon
val E = calcCoupDays(ops).doubleVal // total # days in coupon period in which settlementfalls
val N = PtgCalculator.getLongValue(calcCoupNum(ops)).toDouble() // n is the number of coupons btwn settlement and maturity
val A = calcCoupDayBS(ops).doubleVal // days from beg. coupon period to settlement
val R = rate / frequency
val Y = 1 + yld / frequency
var result = redemption / Math.pow(Y, N - 1 + DSC / E)
var i = 1
while (i <= N)
{
result += (R * 100) / Math.pow(Y, i - 1 + DSC / E)
i++
}
result -= (100.0 * R * A / E)

val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcPRICE= " + result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * PRICEDISC Returns the price per $100 face value of a discounted security
 *
 *
 * PRICEDISC(settlement,maturity,discount,redemption,frequency,basis)
 * Settlement is the security's settlement date. Maturity is the security's
 * maturity date. Discount is the security's discount rate. Redemption is
 * the security's redemption value per $100 face value. Basis is the type of
 * day count basis to use.
 */
    internal fun calcPRICEDISC(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcPRICEDISC")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val discount = operands[2].doubleVal
val redemption = operands[3].doubleVal
var basis = 0
if (operands.size > 4)
basis = operands[4].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4)
return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate))
return PtgErr(PtgErr.ERROR_NUM)
if (redemption <= 0)
return PtgErr(PtgErr.ERROR_NUM)

val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()

 /*
             * Ptg[] ops; if (operands.length > 6) ops = new Ptg[4]; else ops = new
             * Ptg[3]; ops[0] = operands[0]; ops[1] = operands[1]; ops[2] =
             * operands[5]; if (operands.length > 6) ops[3] = operands[6];
             */
            val result = (redemption - ((discount * redemption * yearFrac(basis, settlementDate,
maturityDate))))
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcPRICEDISC= " + result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * PRICEMAT Returns the price per $100 face value of a security that pays
 * interest at maturity
 */
    internal fun calcPRICEMAT(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcPRICEMAT")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val iDate = DateConverter
.getCalendarFromNumber(operands[2].value) as GregorianCalendar
val rate = operands[3].doubleVal
val yld = operands[4].doubleVal
var basis = 0
if (operands.size > 5)
basis = operands[5].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate)) return PtgErr(PtgErr.ERROR_NUM)
if (rate < 0 || yld < 0) return PtgErr(PtgErr.ERROR_NUM)

val settlementDate = (DateConverter.getXLSDateVal(sDate)).toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate)).toLong()
val issueDate = (DateConverter.getXLSDateVal(iDate)).toLong()

val B = getDaysInYearFromBasis(basis, settlementDate, maturityDate)
val DSM = getDaysFromBasis(basis, settlementDate, maturityDate).toDouble()
val DIM = getDaysFromBasis(basis, issueDate, maturityDate).toDouble()
val A = getDaysFromBasis(basis, issueDate, settlementDate).toDouble()

var result = (100 + (DIM / B) * rate * 100.0) / (1 + ((DSM / B) * yld))
result -= (A / B) * rate * 100.0
val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcPRICEMAT= " + result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * PV(rate,nper,pmt,fv,type) Returns the present value of an investment Rate
 * is the interest rate per period. For example, if you obtain an automobile
 * loan at a 10 percent annual interest rate and make monthly payments, your
 * interest rate per month is 10%/12, or 0.83%. You would enter 10%/12, or
 * 0.83%, or 0.0083, into the formula as the rate. Nper is the total number
 * of payment periods in an annuity. For example, if you get a four-year car
 * loan and make monthly payments, your loan has 4*12 (or 48) periods. You
 * would enter 48 into the formula for nper. Pmt is the payment made each
 * period and cannot change over the life of the annuity. Typically, pmt
 * includes principal and interest but no other fees or taxes. For example,
 * the monthly payments on a $10,000, four-year car loan at 12 percent are
 * $263.33. You would enter -263.33 into the formula as the pmt. If pmt is
 * omitted, you must include the fv argument. Fv is the future value, or a
 * cash balance you want to attain after the last payment is made. If fv is
 * omitted, it is assumed to be 0 (the future value of a loan, for example,
 * is 0). For example, if you want to save $50,000 to pay for a special
 * project in 18 years, then $50,000 is the future value. You could then
 * make a conservative guess at an interest rate and determine how much you
 * must save each month. If fv is omitted, you must include the pmt
 * argument. Type is the number 0 or 1 and indicates when payments are due.
 * 0 or omitted At the end of the period 1 At the beginning of the period
 */
    internal fun calcPV(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcPV")
val rate = operands[0].doubleVal
val nper = operands[1].doubleVal
var pmt = 0.0
 // TODO: No specified error trapping?
        if (!(operands[2] is PtgMissArg))
pmt = operands[2].doubleVal
var fv = 0.0
var type = 0
if (operands.size > 3 && !(operands[3] is PtgMissArg))
fv = operands[3].doubleVal
if (operands.size > 4)
type = operands[4].intVal

val A = Math.pow(1 + rate, nper)
var B = pmt * (1 + rate * type)
B *= (A - 1) / rate
val result = (-fv - B) / A

 //double testresults= (result * A + pmt*(1+rate*type) * (A-1)/rate + fv);
        //testresults must==0  for Pv to be correct


        val pnum = PtgNumber(result)
if (DEBUG)
Logger.logInfo("Result from calcPV= " + result)
return pnum

}

/**
 * RATE Returns the interest rate per period of an annuity. RATE is
 * calculated by iteration and can have zero or more solutions. If the
 * successive results of RATE do not converge to within 0.0000001 after 20
 * iterations, RATE returns the #NUM! error value.
 *
 *
 * RATE(nper,pmt,pv,fv,type,guess)
 *
 *
 * Nper is the total number of payment periods in an annuity. Pmt is the
 * payment made each period and cannot change over the life of the annuity.
 * Typically, pmt includes principal and interest but no other fees or
 * taxes. If pmt is omitted, you must include the fv argument. Pv is the
 * present value  the total amount that a series of future payments is
 * worth now. Fv is the future value, or a cash balance you want to attain
 * after the last payment is made. If fv is omitted, it is assumed to be 0
 * (the future value of a loan, for example, is 0). Type is the number 0 or
 * 1 and indicates when payments are due. Guess is your guess for what the
 * rate will be.
 */
    internal fun calcRate(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcRate")
val nper = operands[0].doubleVal
val pmt = operands[1].doubleVal
var pv = 0.0
var fv = 0.0
var type = 0
var guess = 0.1
if (!(operands[2] is PtgMissArg))
pv = operands[2].doubleVal
if (operands.size > 3)
{
if (!(operands[3] is PtgMissArg))
fv = operands[3].doubleVal
}
if (operands.size > 4)
{
if (!(operands[4] is PtgMissArg))
type = operands[4].intVal
}
if (operands.size > 5)
guess = operands[5].doubleVal
 // validate params
        if (type != 0 && type != 1)
return PtgErr(PtgErr.ERROR_NUM)
if (pv == 0.0 && fv == 0.0)
return PtgErr(PtgErr.ERROR_NUM)

 // iterate over possible Rate values; value is correct when
        // f <= tolerance, defined as .00001%
        val TOLERANCE = 0.000000001
var bIsCorrect = false
var x0 = guess
var x1 = x0
val fx0:Double
val fprimex0:Double
 // iterate, using Newton's approximation
        var j = 0
while (j < 100 && !bIsCorrect)
{ // maximum 20 tries
 // Calculate f(x) = (a+ f*g*h + c)
            val R = Math.pow(1 + x0, nper)
val U = 1 / x0
val a = pv * R
val f = pmt * (1 + x0 * type)
val g = R - 1
val h = U
fx0 = a + f * g * U + fv
 //
            // (a + f*g*h + c)'
            //   = a' + f'gh + fg'h + fgh'
            val T = Math.pow(1 + x0, nper - 1)
val aprime = pv * nper * T
val fprime = pmt * type
val gprime = nper * T
val hprime = -1 * Math.pow(x0, -2.0)

fprimex0 = aprime + fprime * g * h + f * gprime * h + f * g * hprime
 // calculate x1, the next iteration
            x1 = x0 - fx0 / fprimex0
val delta = x1 - x0
bIsCorrect = (Math.abs(delta) <= TOLERANCE)
x0 = x1
j++
}

 /*		THIS CALCULATION IS OFF!!!
		//  x= rate
		//  f(x)= PV + PMT*((1-(1+x)^-NPER)/x)*(1+x)^TYPE + FV*(1+x)^-NPER = 0
		//  f'(x)=
		double x0, f0, trial, f;
		x0 = 0;
		f0 = pv * pmt + type + fv;
		trial = guess;

		// iterate, using Newton's approximation
		double delta;
		double fprime = 0; // derivative of f(x0)
		for (int j = 0; j < 100 && !bIsCorrect; j++) { // maximum 20 tries
			//trial= xl+delta*fl/(fl-fh);
			// Calculate f(x)
			double R = Math.pow(1 + trial, -nper);
			double T = Math.pow(1 + trial, type);
			f = pv + pmt * ((1 - R) / trial) * T + fv * R;
			// Calculate f'(x)
			double gofx = fv * R;
			double gprimex = fv * -nper * Math.pow(1 + trial, -nper - 1);
			double hofx = T;
			double hprimex = type * Math.pow(1 + trial, type - 1);
			double zofx = (1 - R) / trial;
			double zprimex = (1 + nper * Math.pow(1 + trial, -nper - 1))
					/ Math.pow(trial, 2);
			fprime = zprimex * hofx + zofx * hprimex + gprimex;
			delta = f / fprime;
			// testing!
			if (trial - delta <= 0)
				delta = trial / 2;
			trial -= delta;
			bIsCorrect = (Math.abs(delta) <= TOLERANCE);
		}
/**/
        if (!bIsCorrect)
return PtgErr(PtgErr.ERROR_NUM)

if (DEBUG)
Logger.logInfo("Result from calcRate= " + x1)
val pnum = PtgNumber(x1)
return pnum
}

/**
 * RECEIVED Returns the amount received at maturity for a fully invested
 * security
 */
    internal fun calcReceived(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcRECEIVED")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val investment = operands[2].doubleVal
val rate = operands[3].doubleVal
var basis = 0
if (operands.size > 4)
basis = operands[4].intVal

val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()

 //		double DSM = maturityDate - settlementDate;
 //		double B = getDaysInYearFromBasis(basis, settlementDate, maturityDate);
 //		double result = investment / (1 - (rate * DSM / B));
            val result = investment / (1 - (rate * yearFrac(basis, settlementDate, maturityDate)))
if (DEBUG)
Logger.logInfo("Result from calcRECEIVED= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * SLN Returns the straight-line depreciation of an asset for one period
 * SLN(cost,salvage,life) Cost is the initial cost of the asset. Salvage is
 * the value at the end of the depreciation (sometimes called the salvage
 * value of the asset). Life is the number of periods over which the asset
 * is depreciated (sometimes called the useful life of the asset).
 */
    internal fun calcSLN(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
val cost = operands[0].doubleVal
val salvage = operands[1].doubleVal
val life = operands[2].doubleVal
if (DEBUG)
debugOperands(operands, "calcSLN")
if (life == 0.0)
return PtgErr(PtgErr.ERROR_NUM)
val result = (cost - salvage) / life
if (DEBUG)
Logger.logInfo("Result from calcSLN= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * SYD Returns the sum-of-years' digits depreciation of an asset for a
 * specified period SYD(cost,salvage,life,per) Cost is the initial cost of
 * the asset. Salvage is the value at the end of the depreciation (sometimes
 * called the salvage value of the asset). Life is the number of periods
 * over which the asset is depreciated (sometimes called the useful life of
 * the asset). Per is the period and must use the same units as life.
 */
    internal fun calcSYD(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcSYD")
val cost = operands[0].doubleVal
val salvage = operands[1].doubleVal
val life = operands[2].doubleVal
val per = operands[3].doubleVal
if (life == 0.0)
return PtgErr(PtgErr.ERROR_NUM)
val A = (cost - salvage) * (life - per + 1) * 2.0
val B = life * (life + 1)
val result = A / B
if (DEBUG)
Logger.logInfo("Result from calcSYD= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * TBILLEQ Returns the bond-equivalent yield for a Treasury bill Settlement
 * is the Treasury bill's settlement date. The security settlement date is
 * the date after the issue date when the Treasury bill is traded to the
 * buyer. Maturity is the Treasury bill's maturity date. The maturity date
 * is the date when the Treasury bill expires. Discount is the Treasury
 * bill's discount rate.
 */
    internal fun calcTBillEq(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcTBILLEQ")
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val rate = operands[2].doubleVal

val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()
if (settlementDate >= maturityDate)
return PtgErr(PtgErr.ERROR_NUM)
if (maturityDate - settlementDate > 365)
return PtgErr(PtgErr.ERROR_NUM)
if (rate <= 0)
return PtgErr(PtgErr.ERROR_NUM)

val DSM = (maturityDate - settlementDate).toDouble()
val result:Double
if (DSM <= 182)
{
result = (365 * rate) / (360 - rate * DSM)
}
else
{
val A = DSM / 365
val B = rate * DSM

val C = (2 * A - 1) * B / (B - 360)
val D = Math.pow(A, 2.0) - C
result = (-2 * A + 2 * Math.sqrt(D)) / (2 * A - 1)
}
if (DEBUG)
Logger.logInfo("Result from calcTBILLEQ= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * TBILLPRICE Returns the price per $100 face value for a Treasury bill
 */
    internal fun calcTBillPrice(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcTBILLPRICE")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val rate = operands[2].doubleVal
val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()

if (settlementDate >= maturityDate)
return PtgErr(PtgErr.ERROR_NUM)
if (maturityDate - settlementDate > 365)
return PtgErr(PtgErr.ERROR_NUM)
if (rate <= 0)
return PtgErr(PtgErr.ERROR_NUM)

val DSM = (maturityDate - settlementDate).toDouble()
val result = 100 * (1 - (rate * DSM / 360))
if (DEBUG)
Logger.logInfo("Result from calcTBILLPRICE= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * TBILLYIELD Returns the yield for a Treasury bill
 *
 *
 * VDB Returns the depreciation of an asset for a specified or partial
 * period using a declining balance method
 */
    internal fun calcTBillYield(operands:Array<Ptg>):Ptg {
if (operands.size < 3)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcTBILLYIELD")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val price = operands[2].doubleVal

val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()
if (settlementDate >= maturityDate)
return PtgErr(PtgErr.ERROR_NUM)
if (maturityDate - settlementDate > 365)
return PtgErr(PtgErr.ERROR_NUM)
if (price <= 0)
return PtgErr(PtgErr.ERROR_NUM)

val DSM = (maturityDate - settlementDate).toDouble()
val result = ((100 - price) / price) * (360 / DSM)
if (DEBUG)
Logger.logInfo("Result from calcTBILLYIELD= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * VDB
 * Returns the depreciation of an asset for any period you specify,
 * including partial periods, using the double-declining balance method or
 * some other method you specify. VDB stands for variable declining balance.
 *
 *
 * VDB(cost,salvage,life,start_period,end_period,factor,no_switch)
 */
    internal fun calcVDB(operands:Array<Ptg>):Ptg {
if (operands.size < 5)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG) debugOperands(operands, "calcVDB")
val cost = operands[0].doubleVal
val salvage = operands[1].doubleVal
val life = operands[2].intVal
val start_period = operands[3].intVal
val end_period = operands[4].intVal
var factor = 2
if (operands.size > 5 && !(operands[5] is PtgMissArg))
factor = operands[5].intVal
var bNoSwitch = false
if (operands.size > 6)
bNoSwitch = PtgCalculator.getBooleanValue(operands[6])

if ((cost <= 0 || salvage <= 0 || life <= 0 ||
start_period < 0 || end_period < 0 ||
factor < 0 || end_period > life))
return PtgErr(PtgErr.ERROR_NUM)

var result = 0.0

val ops = arrayOfNulls<Ptg>(5)
ops[0] = PtgNumber(cost)
ops[1] = PtgNumber(salvage)
ops[2] = PtgInt(life)
ops[4] = PtgInt(factor)
if (bNoSwitch)
{    // just sum ddb
for (i in start_period + 1..end_period)
{
ops[3] = PtgInt(i)
result += calcDDB(ops).doubleVal
}
}
else
{  // switch to straight-line depreciation when dep. > ddb calc
var bSwitch = false
var A = 0.0
var i = start_period + 1
while (i <= end_period && !bSwitch)
{
val sl = (cost - A - salvage) / (life - i + 1)
ops[3] = PtgInt(i)
val ddb = calcDDB(ops).doubleVal
if (sl <= ddb)
{
A += ddb
i++
}
else
 // straight-line depreciation is greater than ddb; switch
                    bSwitch = true
}
result = A
 // use straight-line depreciation for rest of period
            for (j in i..end_period)
{
result += (cost - A - salvage) / (life - i + 1)
}
}

if (DEBUG) Logger.logInfo("Result from calcVDB= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * XIRR Returns the internal rate of return for a schedule of cash flows
 * that is not necessarily periodic
 */
    internal fun calcXIRR(operands:Array<Ptg>):Ptg {
if (operands.size < 2)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcXIRR")
var guess = .1
if (operands.size > 2)
guess = operands[2].doubleVal
 // convert references to Ptg[]s
        val values = PtgCalculator.getAllComponents(operands[0])
val dates = PtgCalculator.getAllComponents(operands[1])
if (DEBUG)
debugOperands(values, "calcXIRR")
if (DEBUG)
debugOperands(dates, "calcXIRR")
 /*
         * 'Newton-Raphson method: ' Given PV a function of X, determine the
         * value of X ' such that SUM[PV] = 0, using iteration. ' dPVdX =
         * derivative of PV with respect to X ' NB: if X = 1 + Rate, dX/dRate =
         * 1 ' ' Calculate PV and dPVdX using an arbitrary initial estimate of
         * X. ' Change X by -PV/dPVdX and calculate again. ' Repeat until X
         * changes by less than some arbitrary small amount.
         *
         * Let X = 1 + Rate
         * Let EXPi = -YearFrax(i)
         * PVi = Val(i) * X^(EXPi)
         * dPV/dX = Val(i) * EXPi * X^(EXPi- 1)
         *        = Val(i) * X^(EXPi) *EXPi / X
         * 		  = PVi * EXPi / X
         * dPVSum/dX = SUM[dPV/dX]
         * dPVSum/dX = 1/X * SUM[PVi * EXPi]
         *
         */
        if (values.size != dates.size)
return PtgErr(PtgErr.ERROR_NUM)
 // validate dates
        val date0 = dates[0].doubleVal
for (i in 1 until dates.size)
{
 // TODO: if not valid date, return PtgErr.ERROR_VALUE
            val `val` = PtgCalculator.getLongValue(dates[i])
if (`val` < date0)
return PtgErr(PtgErr.ERROR_NUM)
}
 // validate values: sum all outflows (= negative values) + inflows (=
        // positive values)
        var outflow = 0.0
var inflow = 0.0
 // get outflow (- values)
        for (i in values.indices)
{
val `val` = values[i].doubleVal
if (`val` < 0)
outflow += Math.abs(`val`)
else
inflow += `val`
}
if (outflow == 0.0 || inflow == 0.0)
return PtgErr(PtgErr.ERROR_NUM)

val n = values.size

 // iterate over possible irr values; value is correct when
        // f <= tolerance, defined as .000001%
        var bIsCorrect = false
val TOLERANCE = 0.00000001
var trial = guess
val x0:Double
val f0:Double
val f:Double
x0 = 1.0 //1+0 rate, x_0, lower bounds of guess
f0 = inflow - outflow // =sum of values when rate is 0
trial = 1 + guess // x_1, upper bounds
 // for secant method
        //  double f1= 0.0;
        //	double x1;
        //	double dx= xh-xl;
        var delta = 0.0
var fprime = 0.0 // derivative of f(x0)
for (i in 0 until n)
{
val `val` = values[i].doubleVal
val exp = (PtgCalculator.getLongValue(dates[i]) - date0) / 365
 //		f+= val*(Math.pow(xh, -exp));
            fprime += `val` * -exp // f' of x0
}
var j = 0
while (j < 100 && !bIsCorrect)
{ // maximum 100 tries
 // SECANT METHOD: trial= x0+dx*f0/(f0-f1);
            // NEWTON'S:
            f = 0.0
for (i in 0 until n)
{
val `val` = values[i].doubleVal
val exp = (dates[i].doubleVal - date0) / 365.0
f += `val` * (Math.pow(trial, -exp))
fprime += `val` * (Math.pow(trial, -exp)) * -exp // derivative of
 // f(trial)
            }
fprime = fprime / trial // final f' expression
delta = f / fprime
if (trial - delta <= 0)
delta = trial / 2
trial -= delta
 /*
             * secant method if (f < 0) { delta= xl-trial; xl= trial; fl= f; }
             * else { delta= xh-trial; xh= trial; fh= f; } dx= xh-xl;
             */
            bIsCorrect = (Math.abs(delta) <= TOLERANCE)
j++
}
if (!bIsCorrect)
return PtgErr(PtgErr.ERROR_NUM)

trial -= 1.0
if (DEBUG)
Logger.logInfo("Result from calcXIRR= " + trial)
val pnum = PtgNumber(trial)
return pnum
}

/**
 * XNPV Returns the net present value for a schedule of cash flows that is
 * not necessarily periodic
 */
    internal fun calcXNPV(operands:Array<Ptg>):Ptg {
if (operands.size < 2 || operands[1].components == null)
{
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcXNPV")
val rate = operands[0].doubleVal
 // convert references to Ptg[]s
        val values = PtgCalculator.getAllComponents(operands[1])
val dates = PtgCalculator.getAllComponents(operands[2])
if (DEBUG)
debugOperands(values, "calcXNPV")
if (DEBUG)
debugOperands(dates, "calcXNPV")

if (values.size != dates.size)
return PtgErr(PtgErr.ERROR_NUM)
 // TODO: validate dates
        val date0 = dates[0].doubleVal
for (i in 1 until dates.size)
{
 // TODO: if not valid date, return PtgErr.ERROR_VALUE
            val `val` = PtgCalculator.getLongValue(dates[i])
if (`val` < date0)
return PtgErr(PtgErr.ERROR_NUM)
}
 // validate values: sum all outflows (= negative values) + inflows (=
        // positive values)
        var outflow = 0.0
var inflow = 0.0
 // get outflow (- values)
        for (i in values.indices)
{
val `val` = values[i].doubleVal
if (`val` < 0)
outflow += Math.abs(`val`)
else
inflow += `val`
}
if (outflow == 0.0 || inflow == 0.0)
return PtgErr(PtgErr.ERROR_NUM)

val n = values.size
var result = 0.0
for (i in 0 until n)
{
val `val` = values[i].doubleVal
val exp = (dates[i].doubleVal - date0) / 365.0
result += `val` / (Math.pow(1 + rate, exp))
}
if (DEBUG)
Logger.logInfo("Result from calcXNPV= " + result)
val pnum = PtgNumber(result)
return pnum
}

/**
 * YIELD Returns the yield on a security that pays periodic interest
 *
 *
 * YIELD(settlement,maturity,rate,pr,redemption,frequency,basis) Settlement
 * is the security's settlement date. The security settlement date is the
 * date after the issue date when the security is traded to the buyer.
 * Maturity is the security's maturity date. The maturity date is the date
 * when the security expires. Rate is the security's annual coupon rate. Pr
 * is the security's price per $100 face value. Redemption is the security's
 * redemption value per $100 face value. Frequency is the number of coupon
 * payments per year. For annual payments, frequency = 1; for semiannual,
 * frequency = 2; for quarterly, frequency = 4. Basis (optional) is the type
 * of day count basis to use.
 */
    internal fun calcYIELD(operands:Array<Ptg>):Ptg {
if (operands.size < 6)
{ // not supported by function
return PtgErr(PtgErr.ERROR_NULL)
}
if (DEBUG)
debugOperands(operands, "calcINTRATE")
try
{
val sDate = DateConverter.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val rate = operands[2].doubleVal
val pr = operands[3].doubleVal
val redemption = operands[4].doubleVal
val frequency = operands[5].intVal
var basis = 0
if (operands.size > 6)
basis = operands[6].intVal

 //	 TODO: if dates are not valid, return #VALUE! error
            if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (!sDate.before(mDate)) return PtgErr(PtgErr.ERROR_NUM)
if (rate < 0) return PtgErr(PtgErr.ERROR_NUM)
if (pr <= 0 || redemption <= 0) return PtgErr(PtgErr.ERROR_NUM)

val settlementDate = (DateConverter.getXLSDateVal(sDate)).toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate)).toLong()
val ops:Array<Ptg>
ops = arrayOfNulls<Ptg>(4)
ops[0] = operands[0]
ops[1] = operands[1]
ops[2] = operands[5]
ops[3] = PtgInt(basis)
var result = 0.0
 /*
             * A= coupDayBS E= coupDays n= coupNum DSC= coupDaysNC DSR= getDaysFromBasis
             */
            val n = PtgCalculator.getLongValue(calcCoupNum(ops)).toDouble() // n is the number of coupons btwn settlement and maturity
val E = calcCoupDays(ops).doubleVal // total # days in coupon period in which settlementfalls
val A = calcCoupDayBS(ops).doubleVal // days from beg. coupon period to settlement

if (n <= 1)
{
val DSR = getDaysFromBasis(basis, settlementDate, maturityDate).toDouble()
val R = rate / frequency
val P = pr / 100
result = ((redemption / 100) + R) - (P + ((A / E) * R))
result /= P + ((A / E) * R)
result *= (frequency * E) / DSR
}
else
{
 // for n values > 1, must employ an iterated approach to find yield using formula for price
                val DSC = calcCoupDaysNC(ops).doubleVal // days from  settlement to next coupon
result = yieldIteration(DSC, E, n, A, rate, frequency, redemption, pr)
if (result == -1.0)
{// didn't find it
return PtgErr(PtgErr.ERROR_NUM)
}
}

val pnum = PtgNumber(result)
if (DEBUG) Logger.logInfo("Result from calcYIELD= " + result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

private fun yieldIteration(DSC:Double, E:Double, n:Double, A:Double, rate:Double, frequency:Int, redemption:Double, pr:Double):Double {
 /* Using Newton-Raphson, iterate over possible yield values; value is correct when
		* f <= tolerance, defined as .000001%
		// 100 iterations, formula for PRICE
		// each trial= f(trial)/f'(trial)
		// f(trial)= price calculation
		// f'(trial): (deep breath)
		 *
		 *  if f= price, and
		 *  G and F are given
		 *  C= redemption
		 *  B= 100*R
		 *  E= -(n-1+F)
		 *  Ei = -(i-1+F)
		 *  x= 1+trial/frequency
		 *
		 *  then,
		 *
		 *  ui(trial)= B*x^Ei
		 *  f'(ui(trial))= -Ei*B*x^(Ei-1)
		 * 			  = Ei*B*x^Ei*1/x
		 * 			  = B*x^Ei*1/x*Ei
		 * 			  = ui(trial)*Ei/x
		 *  f'(sum(ui(trial)))= sum(f'(ui))
		 * 			  = sum(ui(trial)*Ei/x
		 * 			  = 1/x*sum(ui(trial)*Ei)
		 *
		 *
		 *  f(trial)= C*x^E + sum(ui(trial)) - B*G
		 *  f'(trial)= f'(f(x)+sum(ui(trial))
		 * 			 = f'(x) + f'(sum(ui(trial))
		 * 			 = C*E*x^(E-1) + sum(ui(trial)*Ei/x)
		 * 			 = (C*E*x^E)*(1/x) + (1/x)*sum(ui(trial)*Ei)
		 * 			 = (C*E*x^E)*(1/x) + (1/x)*sum(B*x^Ei*Ei)
		 */
        val guess = .1
var bIsCorrect = false
val TOLERANCE = 0.00000001
var trial = guess
val f:Double
val F = DSC / E
val G = A / E
val R = rate / frequency
val Y:Double
val B = R * 100
val Exp = n - 1 + F

 // N-R iteration
        val fprime:Double // derivative of f(x0)
var delta = 0.0
var j = 0
while (j < 100 && !bIsCorrect)
{ // maximum 100 tries
 // f= f(trial)
            Y = 1 + trial / frequency
f = redemption / Math.pow(Y, Exp)
fprime = 0.0
var i = 1
while (i <= n)
{
f += B / Math.pow(Y, i - 1 + F)
fprime += (B / Math.pow(Y, i - 1 + F)) * (i - 1 + F)
i++
}
f -= (B * G)
 // pr-f =>0
            fprime /= Y        // final
fprime += redemption / Math.pow(Y, Exp) * Exp * (1 / Y)

 // N-R:  use f/fprime as iterative factor
            delta = (pr - f) / fprime
if (trial < delta)
trial = delta - trial
else
trial = trial - delta
 //			while (trial - delta <= 0) delta= delta/2;	// sanity check
 //			trial-= delta;
            bIsCorrect = (Math.abs(pr - f) <= TOLERANCE)
j++
}
if (bIsCorrect)
return (trial)
else
return -1.0
 /*
		MinYield = -1#
	      MaxYield = .Rate
	      If MaxYield = 0 Then MaxYield = 0.1
	      Do While CalculatedPrice(BondInfo, MaxYield) > .Price
	        MaxYield = MaxYield * 2
	      Loop


	      Yld = 0.5 * (MinYield + MaxYield)
	      For i = 1 To MaxIterations
	        Diff = CalculatedPrice(BondInfo, Yld) - .Price
	        If Abs(Diff) < Accuracy Then Exit For
	        'if calculated price is greater, correct yield is greater
	        If Diff > 0 Then MinYield = Yld Else MaxYield = Yld
	        Yld = 0.5 * (MinYield + MaxYield)
	      Next i
	    End If
	    BondYield = Yld
*/
    }

/**
 * YIELDDISC Returns the annual yield for a discounted security. For
 * example, a treasury bill
 */
    internal fun calcYieldDisc(operands:Array<Ptg>):Ptg {
if (operands.size < 4)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcYIELDDISC")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val pr = operands[2].doubleVal
val redemption = operands[3].doubleVal
var basis = 0
if (operands.size > 4)
basis = operands[4].intVal

val settlementDate = (DateConverter.getXLSDateVal(sDate)).toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate)).toLong()

if (pr <= 0 || redemption <= 0) return PtgErr(PtgErr.ERROR_NUM)
if (basis < 0 || basis > 4) return PtgErr(PtgErr.ERROR_NUM)
if (maturityDate <= settlementDate) return PtgErr(PtgErr.ERROR_NUM)

 //		double DSM = maturityDate - settlementDate;
 //		double B = getDaysInYearFromBasis(basis, settlementDate, maturityDate);
 //		double result = ((redemption - pr) / pr) * (B / DSM);
            val result = ((redemption - pr) / pr) / yearFrac(basis, settlementDate, maturityDate)
if (DEBUG)
Logger.logInfo("Result from calcYIELDDISC= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

/**
 * YIELDMAT Returns the annual yield of a security that pays interest at
 * maturity
 */
    internal fun calcYieldMat(operands:Array<Ptg>):Ptg {
if (operands.size < 5)
{
val perr = PtgErr(PtgErr.ERROR_NULL)
return perr
}
if (DEBUG)
debugOperands(operands, "calcYIELDMAT")
try
{
val sDate = DateConverter
.getCalendarFromNumber(operands[0].value) as GregorianCalendar
val mDate = DateConverter
.getCalendarFromNumber(operands[1].value) as GregorianCalendar
val iDate = DateConverter
.getCalendarFromNumber(operands[2].value) as GregorianCalendar
val rate = operands[3].doubleVal
val price = operands[4].doubleVal
var basis = 0
if (operands.size > 5)
basis = operands[5].intVal

val settlementDate = (DateConverter.getXLSDateVal(sDate))
.toLong()
val maturityDate = (DateConverter.getXLSDateVal(mDate))
.toLong()
val issueDate = (DateConverter.getXLSDateVal(iDate))
.toLong()

val B = getDaysInYearFromBasis(basis, settlementDate, maturityDate)
val DSM = getDaysFromBasis(basis, settlementDate, maturityDate).toDouble()
val DIM = getDaysFromBasis(basis, issueDate, maturityDate).toDouble()
val A = getDaysFromBasis(basis, issueDate, settlementDate).toDouble()

val result = (((1 + (DIM / B) * rate) - ((price / 100) + (A / B) * rate)) / ((price / 100) + (A / B) * rate)) * (B / DSM)
if (DEBUG) Logger.logInfo("Result from calcYIELDMAT= " + result)
val pnum = PtgNumber(result)
return pnum
}
catch (e:Exception) {}

return PtgErr(PtgErr.ERROR_VALUE)
}

internal fun debugOperands(operands:Array<Ptg>, f:String) {
if (DEBUG)
{
Logger.logInfo("Operands for " + f)
for (i in operands.indices)
{
val s = operands[i].string
if (!(operands[i] is PtgMissArg))
{
val v = operands[i].value.toString()
Logger.logInfo("\tOperand[" + i + "]=" + s + " " + v)
}
else
Logger.logInfo("\tOperand[" + i + "]=" + s + " is Missing")
}
}
}

 fun isLeapYear(year:Int):Boolean {
return (year % 400 == 0) || ((year % 100 != 0) && (year % 4 == 0))
}
}
