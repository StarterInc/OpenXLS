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

import io.starter.formats.XLS.FunctionNotSupportedException
import io.starter.formats.XLS.XLSRecord

import java.util.Locale

/**
 * Function Handler takes an array of PTG's with a header PtgFunc or PtgFuncVar, calcuates
 * those ptgs in the determined way, then return a relevant PtgValue
 *
 *
 * Descriptions of these functions are available on the msdn site,
 * http://msdn.microsoft.com/library/default.asp?url=/library/en-us/office97/html/S88F9.asp
 */


object FunctionHandler {


    /*
        Calculates the function and returns a relevant Ptg as a value
        This is the main entry method.

    */
    @Throws(FunctionNotSupportedException::class, CalculationException::class)
    fun calculateFunction(ptgs: Array<Ptg>): Ptg? {
        val funk: Ptg // the function identifier
        val operands: Array<Ptg> // the ptgs acted upon by the function
        val funkId = 0   //what function are we calling?

        funk = ptgs[0]
        // if ptgs are missing parent_recs, populate from funk
        val bpar = funk.parentRec
        if (bpar != null) {
            for (t in ptgs.indices) {
                if (ptgs[t].parentRec == null)
                    ptgs[t].parentRec = bpar
            }
        }

        val oplen = ptgs.size - 1
        operands = arrayOfNulls(oplen)
        System.arraycopy(ptgs, 1, operands, 0, oplen)
        if (funk.opcode.toInt() == 0x21 || funk.opcode.toInt() == 0x41 || funk.opcode.toInt() == 0x61) {  // ptgfunc
            return calculatePtgFunc(funk, funkId, operands)
        } else if (funk.opcode.toInt() == 0x22 || funk.opcode.toInt() == 0x42 || funk.opcode.toInt() == 0x62) { // ptgfuncvar
            return calculatePtgFuncVar(funk, funkId, operands)
        }
        return null
    }

    /*
        Keep the calculation of ptgfunc & ptgfuncvar seperate in case any differences show up
    */
    @Throws(FunctionNotSupportedException::class, CalculationException::class)
    fun calculatePtgFunc(funk: Ptg, funkId: Int, operands: Array<Ptg>): Ptg? {
        var funkId = funkId
        val pf = funk as PtgFunc
        funkId = pf.`val`
        return parse_n_calc(funk, funkId, operands)
    }


    /**
     * Keep the calculation of ptgfunc & ptgfuncvar seperate in case any differences show up
     *
     * @throws CalculationException
     */
    @Throws(FunctionNotSupportedException::class, CalculationException::class)
    fun calculatePtgFuncVar(funk: Ptg, funkId: Int, operands: Array<Ptg>): Ptg? {
        var funkId = funkId
        var operands = operands
        val pf = funk as PtgFuncVar
        funkId = pf.`val`
        // Handle Add-in Formulas - which have a name operand 1st
        if (funkId == FunctionConstants.xlfADDIN) {  // XL flag that formula is an add-in
            //	must pop the PtgNameX record to get the correct function id
            var s = ""
            var foundit = false
            if (operands[0] is PtgNameX) {
                val index = (operands[0] as PtgNameX).`val`
                s = pf.parentRec!!.sheet!!.workBook!!.getExternalName(index)
            } else if (operands[0] is PtgName) {
                s = (operands[0] as PtgName).storedName
            }
            if (s.startsWith("_xlfn.")) {    // Excel "new" functions
                s = s.substring(6)
            }
            if (Locale.JAPAN == Locale.getDefault()) {
                var y = 0
                while (y < FunctionConstants.jRecArr.size) {
                    if (s.equals(FunctionConstants.jRecArr[y][0], ignoreCase = true)) {
                        funkId = Integer.valueOf(FunctionConstants.jRecArr[y][1]).toInt()
                        y = FunctionConstants.jRecArr.size  // exit loop
                        foundit = true
                    }
                    y++
                }
            }
            if (!foundit) {
                var y = 0
                while (y < FunctionConstants.recArr.size) {    // Use FunctionConstants instead of PtFuncVar
                    if (s.equals(FunctionConstants.recArr[y][0], ignoreCase = true)) {
                        funkId = Integer.valueOf(FunctionConstants.recArr[y][1]).toInt()
                        y = FunctionConstants.recArr.size    // exit loop
                    }
                    y++
                }
            }
            if (funkId == 255)
            // it's not found
                throw FunctionNotSupportedException(s)

            //			now get rid of PtgNameX operand before calling function
            val ops = arrayOfNulls<Ptg>(operands.size - 1)
            System.arraycopy(operands, 1, ops, 0, operands.size - 1)
            operands = arrayOfNulls(ops.size)
            System.arraycopy(ops, 0, operands, 0, ops.size)
        }    // end KSC added
        return parse_n_calc(funk, funkId, operands)
    }


    /************************************************************************************
     * *
     * Your standard big case statement, calling methods based on what the funkid is. *
     * You will notice that these are seperated out into packages based on the MS     *
     * documentation (Link above).  Each package calls a different class full of      *
     * static method calls.   There are a lot of these :-)                            *
     * *
     * PLEASE:  Remove function from comment list when you enable it!!!!           *
     * *
     * @throws CalculationException
     */
    @Throws(FunctionNotSupportedException::class, CalculationException::class)
    fun parse_n_calc(function: Ptg, functionId: Int, operands: Array<Ptg>): Ptg? {
        var operands = operands
        var resultPtg: Ptg? = null
        val resultArrPtg: Array<Ptg>? = null

        when (functionId) {
            /********************************************
             * Database and List package functions    **
             */
            FunctionConstants.xlfDaverage -> resultPtg = DatabaseCalculator.calcDAverage(operands)

            FunctionConstants.xlfDcount -> resultPtg = DatabaseCalculator.calcDCount(operands)

            FunctionConstants.xlfDcounta -> resultPtg = DatabaseCalculator.calcDCountA(operands)

            FunctionConstants.xlfDget -> resultPtg = DatabaseCalculator.calcDGet(operands)

            FunctionConstants.xlfDmax -> resultPtg = DatabaseCalculator.calcDMax(operands)

            FunctionConstants.xlfDmin -> resultPtg = DatabaseCalculator.calcDMin(operands)

            FunctionConstants.xlfDproduct -> resultPtg = DatabaseCalculator.calcDProduct(operands)

            FunctionConstants.xlfDstdev -> resultPtg = DatabaseCalculator.calcDStdDev(operands)

            FunctionConstants.xlfDstdevp -> resultPtg = DatabaseCalculator.calcDStdDevP(operands)

            FunctionConstants.xlfDsum -> resultPtg = DatabaseCalculator.calcDSum(operands)

            FunctionConstants.xlfDvar -> resultPtg = DatabaseCalculator.calcDVar(operands)

            FunctionConstants.xlfDvarp -> resultPtg = DatabaseCalculator.calcDVarP(operands)

            /********************************************
             * Date and time functions         *********
             */
            FunctionConstants.xlfDate -> resultPtg = DateTimeCalculator.calcDate(operands)

            FunctionConstants.xlfDay -> resultPtg = DateTimeCalculator.calcDay(operands)

            FunctionConstants.xlfDays360 -> resultPtg = DateTimeCalculator.calcDays360(operands)

            FunctionConstants.xlfHour -> resultPtg = DateTimeCalculator.calcHour(operands)

            FunctionConstants.xlfMinute -> resultPtg = DateTimeCalculator.calcMinute(operands)

            FunctionConstants.xlfMonth -> resultPtg = DateTimeCalculator.calcMonth(operands)

            FunctionConstants.xlfYear -> resultPtg = DateTimeCalculator.calcYear(operands)

            FunctionConstants.xlfSecond -> resultPtg = DateTimeCalculator.calcSecond(operands)

            FunctionConstants.xlfTimevalue -> resultPtg = DateTimeCalculator.calcTimevalue(operands)

            FunctionConstants.xlfWeekday -> resultPtg = DateTimeCalculator.calcWeekday(operands)

            FunctionConstants.xlfWEEKNUM -> resultPtg = DateTimeCalculator.calcWeeknum(operands)

            FunctionConstants.xlfWORKDAY -> resultPtg = DateTimeCalculator.calcWorkday(operands)

            FunctionConstants.xlfYEARFRAC -> resultPtg = DateTimeCalculator.calcYearFrac(operands)

            FunctionConstants.xlfNow -> resultPtg = DateTimeCalculator.calcNow(operands)

            FunctionConstants.xlfTime -> resultPtg = DateTimeCalculator.calcTime(operands)

            FunctionConstants.xlfToday -> resultPtg = DateTimeCalculator.calcToday(operands)

            FunctionConstants.xlfDatevalue -> resultPtg = DateTimeCalculator.calcDateValue(operands)

            FunctionConstants.xlfEDATE -> resultPtg = DateTimeCalculator.calcEdate(operands)

            FunctionConstants.xlfEOMONTH -> resultPtg = DateTimeCalculator.calcEOMonth(operands)

            FunctionConstants.xlfNETWORKDAYS -> resultPtg = DateTimeCalculator.calcNetWorkdays(operands)

            /********************************************
             * DDE and External functions           ****
             */


            /********************************************
             * Engineering functions           *********
             */
            FunctionConstants.xlfBIN2DEC -> resultPtg = EngineeringCalculator.calcBin2Dec(operands)

            FunctionConstants.xlfBIN2HEX -> resultPtg = EngineeringCalculator.calcBin2Hex(operands)

            FunctionConstants.xlfBIN2OCT -> resultPtg = EngineeringCalculator.calcBin2Oct(operands)

            FunctionConstants.xlfDEC2BIN -> resultPtg = EngineeringCalculator.calcDec2Bin(operands)

            FunctionConstants.xlfDEC2HEX -> resultPtg = EngineeringCalculator.calcDec2Hex(operands)

            FunctionConstants.xlfDEC2OCT -> resultPtg = EngineeringCalculator.calcDec2Oct(operands)

            FunctionConstants.xlfHEX2BIN -> resultPtg = EngineeringCalculator.calcHex2Bin(operands)

            FunctionConstants.xlfHEX2DEC -> resultPtg = EngineeringCalculator.calcHex2Dec(operands)

            FunctionConstants.xlfHEX2OCT -> resultPtg = EngineeringCalculator.calcHex2Oct(operands)

            FunctionConstants.xlfOCT2BIN -> resultPtg = EngineeringCalculator.calcOct2Bin(operands)

            FunctionConstants.xlfOCT2DEC -> resultPtg = EngineeringCalculator.calcOct2Dec(operands)

            FunctionConstants.xlfOCT2HEX -> resultPtg = EngineeringCalculator.calcOct2Hex(operands)

            FunctionConstants.xlfCOMPLEX -> resultPtg = EngineeringCalculator.calcComplex(operands)

            FunctionConstants.xlfGESTEP -> resultPtg = EngineeringCalculator.calcGEStep(operands)

            FunctionConstants.xlfDELTA -> resultPtg = EngineeringCalculator.calcDelta(operands)

            FunctionConstants.xlfIMAGINARY -> resultPtg = EngineeringCalculator.calcImaginary(operands)

            FunctionConstants.xlfIMREAL -> resultPtg = EngineeringCalculator.calcImReal(operands)

            FunctionConstants.xlfIMARGUMENT -> resultPtg = EngineeringCalculator.calcImArgument(operands)

            FunctionConstants.xlfIMABS -> resultPtg = EngineeringCalculator.calcImAbs(operands)

            FunctionConstants.xlfIMDIV -> resultPtg = EngineeringCalculator.calcImDiv(operands)

            FunctionConstants.xlfIMCONJUGATE -> resultPtg = EngineeringCalculator.calcImConjugate(operands)

            FunctionConstants.xlfIMCOS -> resultPtg = EngineeringCalculator.calcImCos(operands)

            FunctionConstants.xlfIMSIN -> resultPtg = EngineeringCalculator.calcImSin(operands)

            FunctionConstants.xlfIMEXP -> resultPtg = EngineeringCalculator.calcImExp(operands)

            FunctionConstants.xlfIMSUB -> resultPtg = EngineeringCalculator.calcImSub(operands)

            FunctionConstants.xlfIMSUM -> resultPtg = EngineeringCalculator.calcImSum(operands)

            FunctionConstants.xlfIMPRODUCT -> resultPtg = EngineeringCalculator.calcImProduct(operands)

            FunctionConstants.xlfIMLN -> resultPtg = EngineeringCalculator.calcImLn(operands)

            FunctionConstants.xlfIMLOG10 -> resultPtg = EngineeringCalculator.calcImLog10(operands)

            FunctionConstants.xlfIMLOG2 -> resultPtg = EngineeringCalculator.calcImLog2(operands)

            FunctionConstants.xlfIMPOWER -> resultPtg = EngineeringCalculator.calcImPower(operands)

            FunctionConstants.xlfIMSQRT -> resultPtg = EngineeringCalculator.calcImSqrt(operands)

            FunctionConstants.xlfCONVERT -> resultPtg = EngineeringCalculator.calcConvert(operands)

            FunctionConstants.xlfERF -> resultPtg = EngineeringCalculator.calcErf(operands)


            /********************************************
             * Financial functions             *********
             */

            FunctionConstants.xlfDb -> resultPtg = FinancialCalculator.calcDB(operands)

            FunctionConstants.xlfDdb -> resultPtg = FinancialCalculator.calcDDB(operands)

            FunctionConstants.xlfPmt -> resultPtg = FinancialCalculator.calcPmt(operands)

            // KSC: Added
            FunctionConstants.xlfAccrintm -> resultPtg = FinancialCalculator.calcAccrintm(operands)

            FunctionConstants.xlfAccrint -> resultPtg = FinancialCalculator.calcAccrint(operands)

            FunctionConstants.xlfCoupDayBS -> resultPtg = FinancialCalculator.calcCoupDayBS(operands)

            FunctionConstants.xlfCoupDays -> resultPtg = FinancialCalculator.calcCoupDays(operands)

            FunctionConstants.xlfNpv -> resultPtg = FinancialCalculator.calcNPV(operands)

            FunctionConstants.xlfPv -> resultPtg = FinancialCalculator.calcPV(operands)

            FunctionConstants.xlfFv -> resultPtg = FinancialCalculator.calcFV(operands)

            FunctionConstants.xlfIpmt -> resultPtg = FinancialCalculator.calcIPMT(operands)

            FunctionConstants.xlfCumIPmt -> resultPtg = FinancialCalculator.calcCumIPmt(operands)

            FunctionConstants.xlfCumPrinc -> resultPtg = FinancialCalculator.calcCumPrinc(operands)

            FunctionConstants.xlfCoupNCD -> resultPtg = FinancialCalculator.calcCoupNCD(operands)

            FunctionConstants.xlfCoupDaysNC -> resultPtg = FinancialCalculator.calcCoupDaysNC(operands)

            FunctionConstants.xlfCoupPCD -> resultPtg = FinancialCalculator.calcCoupPCD(operands)

            FunctionConstants.xlfCoupNUM -> resultPtg = FinancialCalculator.calcCoupNum(operands)

            FunctionConstants.xlfDollarDE -> resultPtg = FinancialCalculator.calcDollarDE(operands)

            FunctionConstants.xlfDollarFR -> resultPtg = FinancialCalculator.calcDollarFR(operands)

            FunctionConstants.xlfEffect -> resultPtg = FinancialCalculator.calcEffect(operands)

            FunctionConstants.xlfRECEIVED -> resultPtg = FinancialCalculator.calcReceived(operands)

            FunctionConstants.xlfINTRATE -> resultPtg = FinancialCalculator.calcINTRATE(operands)

            FunctionConstants.xlfIrr -> resultPtg = FinancialCalculator.calcIRR(operands)

            FunctionConstants.xlfMirr -> resultPtg = FinancialCalculator.calcMIRR(operands)

            FunctionConstants.xlfXIRR -> resultPtg = FinancialCalculator.calcXIRR(operands)

            FunctionConstants.xlfXNPV -> resultPtg = FinancialCalculator.calcXNPV(operands)

            FunctionConstants.xlfRate -> resultPtg = FinancialCalculator.calcRate(operands)

            FunctionConstants.xlfYIELD -> resultPtg = FinancialCalculator.calcYIELD(operands)

            FunctionConstants.xlfPRICE -> resultPtg = FinancialCalculator.calcPRICE(operands)

            FunctionConstants.xlfPRICEDISC -> resultPtg = FinancialCalculator.calcPRICEDISC(operands)

            FunctionConstants.xlfPRICEMAT -> resultPtg = FinancialCalculator.calcPRICEMAT(operands)

            FunctionConstants.xlfDISC -> resultPtg = FinancialCalculator.calcDISC(operands)

            FunctionConstants.xlfNper -> resultPtg = FinancialCalculator.calcNPER(operands)

            FunctionConstants.xlfSln -> resultPtg = FinancialCalculator.calcSLN(operands)

            FunctionConstants.xlfSyd -> resultPtg = FinancialCalculator.calcSYD(operands)

            FunctionConstants.xlfDURATION -> resultPtg = FinancialCalculator.calcDURATION(operands)

            FunctionConstants.xlfMDURATION -> resultPtg = FinancialCalculator.calcMDURATION(operands)

            FunctionConstants.xlfTBillEq -> resultPtg = FinancialCalculator.calcTBillEq(operands)

            FunctionConstants.xlfTBillPrice -> resultPtg = FinancialCalculator.calcTBillPrice(operands)

            FunctionConstants.xlfTBillYield -> resultPtg = FinancialCalculator.calcTBillYield(operands)

            FunctionConstants.xlfYieldDisc -> resultPtg = FinancialCalculator.calcYieldDisc(operands)

            FunctionConstants.xlfYieldMat -> resultPtg = FinancialCalculator.calcYieldMat(operands)

            FunctionConstants.xlfPpmt -> resultPtg = FinancialCalculator.calcPPMT(operands)

            FunctionConstants.xlfFVSchedule -> resultPtg = FinancialCalculator.calcFVSCHEDULE(operands)

            FunctionConstants.xlfIspmt -> resultPtg = FinancialCalculator.calcISPMT(operands)

            FunctionConstants.xlfAmorlinc -> resultPtg = FinancialCalculator.calcAmorlinc(operands)

            FunctionConstants.xlfAmordegrc -> resultPtg = FinancialCalculator.calcAmordegrc(operands)

            FunctionConstants.xlfOddFPrice -> resultPtg = FinancialCalculator.calcODDFPRICE(operands)

            FunctionConstants.xlfOddFYield -> resultPtg = FinancialCalculator.calcODDFYIELD(operands)

            FunctionConstants.xlfOddLPrice -> resultPtg = FinancialCalculator.calcODDLPRICE(operands)

            FunctionConstants.xlfOddLYield -> resultPtg = FinancialCalculator.calcODDLYIELD(operands)

            FunctionConstants.xlfNOMINAL -> resultPtg = FinancialCalculator.calcNominal(operands)

            FunctionConstants.xlfVdb -> resultPtg = FinancialCalculator.calcVDB(operands)
            /********************************************
             * Information functions           *********
             */

            FunctionConstants.xlfCell -> resultPtg = InformationCalculator.calcCell(operands)

            FunctionConstants.xlfInfo -> resultPtg = InformationCalculator.calcInfo(operands)

            FunctionConstants.XLF_IS_NA -> resultPtg = InformationCalculator.calcIsna(operands)

            FunctionConstants.XLF_IS_ERROR -> resultPtg = InformationCalculator.calcIserror(operands)

            FunctionConstants.xlfIserr -> resultPtg = InformationCalculator.calcIserr(operands)

            FunctionConstants.xlfErrorType -> resultPtg = InformationCalculator.calcErrorType(operands)

            FunctionConstants.xlfNa -> resultPtg = InformationCalculator.calcNa(operands)

            FunctionConstants.xlfIsblank -> resultPtg = InformationCalculator.calcIsBlank(operands)

            FunctionConstants.xlfIslogical -> resultPtg = InformationCalculator.calcIsLogical(operands)

            FunctionConstants.xlfIsnontext -> resultPtg = InformationCalculator.calcIsNonText(operands)

            FunctionConstants.xlfIstext -> resultPtg = InformationCalculator.calcIsText(operands)

            FunctionConstants.xlfIsref -> resultPtg = InformationCalculator.calcIsRef(operands)

            FunctionConstants.xlfN -> resultPtg = InformationCalculator.calcN(operands)

            FunctionConstants.xlfIsnumber -> resultPtg = InformationCalculator.calcIsNumber(operands)

            FunctionConstants.xlfISEVEN -> resultPtg = InformationCalculator.calcIsEven(operands)

            FunctionConstants.xlfISODD -> resultPtg = InformationCalculator.calcIsOdd(operands)

            FunctionConstants.xlfType -> resultPtg = InformationCalculator.calcType(operands)

            /********************************************
             * Logical functions       *****
             */
            FunctionConstants.xlfAnd -> resultPtg = LogicalCalculator.calcAnd(operands)

            FunctionConstants.xlfFalse -> resultPtg = LogicalCalculator.calcFalse(operands)

            FunctionConstants.xlfTrue -> resultPtg = LogicalCalculator.calcTrue(operands)

            FunctionConstants.XLF_IS -> resultPtg = LogicalCalculator.calcIf(operands)

            FunctionConstants.xlfNot -> resultPtg = LogicalCalculator.calcNot(operands)

            FunctionConstants.xlfOr -> resultPtg = LogicalCalculator.calcOr(operands)

            FunctionConstants.xlfIFERROR -> resultPtg = LogicalCalculator.calcIferror(operands)

            /********************************************
             * Lookup and reference functions       *****
             */
            FunctionConstants.xlfAddress -> resultPtg = LookupReferenceCalculator.calcAddress(operands)

            FunctionConstants.xlfAreas -> resultPtg = LookupReferenceCalculator.calcAreas(operands)

            FunctionConstants.xlfChoose -> resultPtg = LookupReferenceCalculator.calcChoose(operands)

            FunctionConstants.xlfColumn -> {
                if (operands.size == 0) {
                    operands = arrayOfNulls(1)
                    operands[0] = function
                }
                resultPtg = LookupReferenceCalculator.calcColumn(operands)
            }

            FunctionConstants.xlfColumns -> resultPtg = LookupReferenceCalculator.calcColumns(operands)

            FunctionConstants.xlfHyperlink -> resultPtg = LookupReferenceCalculator.calcHyperlink(operands)


            FunctionConstants.xlfIndex -> resultPtg = LookupReferenceCalculator.calcIndex(operands)

            FunctionConstants.XLF_INDIRECT -> resultPtg = LookupReferenceCalculator.calcIndirect(operands)

            FunctionConstants.XLF_ROW -> {
                if (operands.size == 0) {
                    operands = arrayOfNulls(1)
                    operands[0] = function
                }
                resultPtg = LookupReferenceCalculator.calcRow(operands)
            }

            FunctionConstants.xlfRows -> resultPtg = LookupReferenceCalculator.calcRows(operands)

            FunctionConstants.xlfTranspose -> resultPtg = LookupReferenceCalculator.calcTranspose(operands)

            FunctionConstants.xlfLookup -> resultPtg = LookupReferenceCalculator.calcLookup(operands)

            FunctionConstants.xlfHlookup -> resultPtg = LookupReferenceCalculator.calcHlookup(operands)

            FunctionConstants.xlfVlookup -> resultPtg = LookupReferenceCalculator.calcVlookup(operands)

            FunctionConstants.xlfMatch -> resultPtg = LookupReferenceCalculator.calcMatch(operands)

            FunctionConstants.xlfOffset -> resultPtg = LookupReferenceCalculator.calcOffset(operands)

            /********************************************
             * Math & Trigonometry functions       *****
             */

            FunctionConstants.XLF_SUM -> resultPtg = MathFunctionCalculator.calcSum(operands)

            FunctionConstants.XLF_SUM_IF -> resultPtg = MathFunctionCalculator.calcSumif(operands)

            FunctionConstants.xlfSUMIFS -> resultPtg = MathFunctionCalculator.calcSumIfS(operands)

            FunctionConstants.xlfSumproduct -> resultPtg = MathFunctionCalculator.calcSumproduct(operands)

            FunctionConstants.xlfExp -> resultPtg = MathFunctionCalculator.calcExp(operands)

            FunctionConstants.xlfAbs -> resultPtg = MathFunctionCalculator.calcAbs(operands)

            FunctionConstants.xlfAcos -> resultPtg = MathFunctionCalculator.calcAcos(operands)

            FunctionConstants.xlfAcosh -> resultPtg = MathFunctionCalculator.calcAcosh(operands)

            FunctionConstants.xlfAsin -> resultPtg = MathFunctionCalculator.calcAsin(operands)

            FunctionConstants.xlfAsinh -> resultPtg = MathFunctionCalculator.calcAsinh(operands)

            FunctionConstants.xlfAtan -> resultPtg = MathFunctionCalculator.calcAtan(operands)

            FunctionConstants.xlfAtan2 -> resultPtg = MathFunctionCalculator.calcAtan2(operands)

            FunctionConstants.xlfAtanh -> resultPtg = MathFunctionCalculator.calcAtanh(operands)

            FunctionConstants.xlfCeiling -> resultPtg = MathFunctionCalculator.calcCeiling(operands)

            FunctionConstants.xlfCombin -> resultPtg = MathFunctionCalculator.calcCombin(operands)

            FunctionConstants.xlfCos -> resultPtg = MathFunctionCalculator.calcCos(operands)

            FunctionConstants.xlfCosh -> resultPtg = MathFunctionCalculator.calcCosh(operands)

            FunctionConstants.xlfDegrees -> resultPtg = MathFunctionCalculator.calcDegrees(operands)

            FunctionConstants.xlfEven -> resultPtg = MathFunctionCalculator.calcEven(operands)

            FunctionConstants.xlfFact -> resultPtg = MathFunctionCalculator.calcFact(operands)

            FunctionConstants.xlfDOUBLEFACT -> resultPtg = MathFunctionCalculator.calcFactDouble(operands)

            FunctionConstants.xlfFloor -> resultPtg = MathFunctionCalculator.calcFloor(operands)

            FunctionConstants.xlfGCD -> resultPtg = MathFunctionCalculator.calcGCD(operands)

            FunctionConstants.xlfInt -> resultPtg = MathFunctionCalculator.calcInt(operands)

            FunctionConstants.xlfLCM -> resultPtg = MathFunctionCalculator.calcLCM(operands)

            FunctionConstants.xlfMROUND -> resultPtg = MathFunctionCalculator.calcMRound(operands)

            FunctionConstants.xlfMmult -> resultPtg = MathFunctionCalculator.calcMMult(operands)

            FunctionConstants.xlfMULTINOMIAL -> resultPtg = MathFunctionCalculator.calcMultinomial(operands)

            FunctionConstants.xlfLn -> resultPtg = MathFunctionCalculator.calcLn(operands)

            FunctionConstants.xlfLog -> resultPtg = MathFunctionCalculator.calcLog(operands)

            FunctionConstants.xlfLog10 -> resultPtg = MathFunctionCalculator.calcLog10(operands)

            FunctionConstants.xlfMod -> resultPtg = MathFunctionCalculator.calcMod(operands)

            FunctionConstants.xlfOdd -> resultPtg = MathFunctionCalculator.calcOdd(operands)

            FunctionConstants.xlfPi -> resultPtg = MathFunctionCalculator.calcPi(operands)

            FunctionConstants.xlfPower -> resultPtg = MathFunctionCalculator.calcPower(operands)

            FunctionConstants.xlfProduct -> resultPtg = MathFunctionCalculator.calcProduct(operands)

            FunctionConstants.xlfQUOTIENT -> resultPtg = MathFunctionCalculator.calcQuotient(operands)

            FunctionConstants.xlfRadians -> resultPtg = MathFunctionCalculator.calcRadians(operands)

            FunctionConstants.xlfRand -> resultPtg = MathFunctionCalculator.calcRand(operands)

            FunctionConstants.xlfRANDBETWEEN -> resultPtg = MathFunctionCalculator.calcRandBetween(operands)

            FunctionConstants.xlfRoman -> resultPtg = MathFunctionCalculator.calcRoman(operands)

            FunctionConstants.xlfRound -> resultPtg = MathFunctionCalculator.calcRound(operands)

            FunctionConstants.xlfRounddown -> resultPtg = MathFunctionCalculator.calcRoundDown(operands)

            FunctionConstants.xlfRoundup -> resultPtg = MathFunctionCalculator.calcRoundUp(operands)

            FunctionConstants.xlfSign -> resultPtg = MathFunctionCalculator.calcSign(operands)

            FunctionConstants.xlfSin -> resultPtg = MathFunctionCalculator.calcSin(operands)

            FunctionConstants.xlfSinh -> resultPtg = MathFunctionCalculator.calcSinh(operands)

            FunctionConstants.xlfSqrt -> resultPtg = MathFunctionCalculator.calcSqrt(operands)

            FunctionConstants.xlfSQRTPI -> resultPtg = MathFunctionCalculator.calcSqrtPi(operands)

            FunctionConstants.xlfTan -> resultPtg = MathFunctionCalculator.calcTan(operands)

            FunctionConstants.xlfTanh -> resultPtg = MathFunctionCalculator.calcTanh(operands)

            FunctionConstants.xlfTrunc -> resultPtg = MathFunctionCalculator.calcTrunc(operands)

            /********************************************
             * Statistical functions       *****
             */

            FunctionConstants.XLF_COUNT -> resultPtg = StatisticalCalculator.calcCount(operands)

            FunctionConstants.xlfCounta -> resultPtg = StatisticalCalculator.calcCountA(operands)

            FunctionConstants.xlfCountblank -> resultPtg = StatisticalCalculator.calcCountBlank(operands)

            FunctionConstants.xlfCountif -> resultPtg = StatisticalCalculator.calcCountif(operands)

            FunctionConstants.xlfCOUNTIFS -> resultPtg = StatisticalCalculator.calcCountIfS(operands)

            FunctionConstants.XLF_MIN -> resultPtg = StatisticalCalculator.calcMin(operands)

            FunctionConstants.xlfMinA -> resultPtg = StatisticalCalculator.calcMinA(operands)

            FunctionConstants.XLF_MAX -> resultPtg = StatisticalCalculator.calcMax(operands)

            FunctionConstants.xlfMaxA -> resultPtg = StatisticalCalculator.calcMaxA(operands)

            FunctionConstants.xlfNormdist -> resultPtg = StatisticalCalculator.calcNormdist(operands)

            FunctionConstants.xlfNormsdist -> resultPtg = StatisticalCalculator.calcNormsdist(operands)

            FunctionConstants.xlfNormsinv -> resultPtg = StatisticalCalculator.calcNormsInv(operands)

            FunctionConstants.xlfNorminv -> resultPtg = StatisticalCalculator.calcNormInv(operands)

            FunctionConstants.XLF_AVERAGE -> resultPtg = StatisticalCalculator.calcAverage(operands)

            FunctionConstants.xlfAVERAGEIF -> resultPtg = StatisticalCalculator.calcAverageIf(operands)

            FunctionConstants.xlfAVERAGEIFS -> resultPtg = StatisticalCalculator.calcAverageIfS(operands)

            FunctionConstants.xlfAvedev -> resultPtg = StatisticalCalculator.calcAveDev(operands)

            FunctionConstants.xlfAverageA -> resultPtg = StatisticalCalculator.calcAverageA(operands)

            FunctionConstants.xlfMedian -> resultPtg = StatisticalCalculator.calcMedian(operands)

            FunctionConstants.xlfMode -> resultPtg = StatisticalCalculator.calcMode(operands)

            FunctionConstants.xlfQuartile -> resultPtg = StatisticalCalculator.calcQuartile(operands)

            FunctionConstants.xlfRank -> resultPtg = StatisticalCalculator.calcRank(operands)

            FunctionConstants.xlfStdev -> resultPtg = StatisticalCalculator.calcStdev(operands)

            FunctionConstants.xlfVar -> resultPtg = StatisticalCalculator.calcVar(operands)

            FunctionConstants.xlfVarp -> resultPtg = StatisticalCalculator.calcVarp(operands)

            FunctionConstants.xlfCovar -> resultPtg = StatisticalCalculator.calcCovar(operands)

            FunctionConstants.xlfCorrel -> resultPtg = StatisticalCalculator.calcCorrel(operands)

            FunctionConstants.xlfFrequency -> resultPtg = StatisticalCalculator.calcFrequency(operands)

            FunctionConstants.xlfLinest -> resultPtg = StatisticalCalculator.calcLineSt(operands)

            FunctionConstants.xlfSlope -> resultPtg = StatisticalCalculator.calcSlope(operands)

            FunctionConstants.xlfIntercept -> resultPtg = StatisticalCalculator.calcIntercept(operands)

            FunctionConstants.xlfPearson -> resultPtg = StatisticalCalculator.calcPearson(operands)

            FunctionConstants.xlfRsq -> resultPtg = StatisticalCalculator.calcRsq(operands)

            FunctionConstants.xlfSteyx -> resultPtg = StatisticalCalculator.calcSteyx(operands)

            FunctionConstants.xlfForecast -> resultPtg = StatisticalCalculator.calcForecast(operands)

            FunctionConstants.xlfTrend -> resultPtg = StatisticalCalculator.calcTrend(operands)

            FunctionConstants.xlfLarge -> resultPtg = StatisticalCalculator.calcLarge(operands)

            FunctionConstants.xlfSmall -> resultPtg = StatisticalCalculator.calcSmall(operands)

            /********************************************
             * Text functions                                *****
             */
            /*
			 * these DBCS functions are not working yet
		case FunctionConstants.xlfAsc:
			resultPtg= TextCalculator.calcAsc(operands);
			break;

		case FunctionConstants.xlfDbcs:
			resultPtg= TextCalculator.calcJIS(operands);
			break;
			*/
            FunctionConstants.xlfChar -> resultPtg = TextCalculator.calcChar(operands)

            FunctionConstants.xlfClean -> resultPtg = TextCalculator.calcClean(operands)

            FunctionConstants.xlfCode -> resultPtg = TextCalculator.calcCode(operands)

            FunctionConstants.xlfConcatenate -> resultPtg = TextCalculator.calcConcatenate(operands)

            FunctionConstants.xlfDollar -> resultPtg = TextCalculator.calcDollar(operands)

            FunctionConstants.xlfExact -> resultPtg = TextCalculator.calcExact(operands)

            FunctionConstants.xlfFind -> resultPtg = TextCalculator.calcFind(operands)

            // DBCS functions are not working 100% yet
            FunctionConstants.xlfFindb -> resultPtg = TextCalculator.calcFindB(operands)

            FunctionConstants.xlfFixed -> resultPtg = TextCalculator.calcFixed(operands)

            FunctionConstants.xlfLeft -> resultPtg = TextCalculator.calcLeft(operands)

            FunctionConstants.xlfLeftb -> resultPtg = TextCalculator.calcLeftB(operands)

            FunctionConstants.xlfLen -> resultPtg = TextCalculator.calcLen(operands)

            FunctionConstants.xlfLenb -> resultPtg = TextCalculator.calcLenB(operands)

            FunctionConstants.xlfLower -> resultPtg = TextCalculator.calcLower(operands)

            FunctionConstants.xlfUpper -> resultPtg = TextCalculator.calcUpper(operands)

            FunctionConstants.xlfMid -> resultPtg = TextCalculator.calcMid(operands)

            FunctionConstants.xlfProper -> resultPtg = TextCalculator.calcProper(operands)

            FunctionConstants.xlfReplace -> resultPtg = TextCalculator.calcReplace(operands)

            FunctionConstants.xlfRept -> resultPtg = TextCalculator.calcRept(operands)

            FunctionConstants.xlfRight -> resultPtg = TextCalculator.calcRight(operands)

            FunctionConstants.xlfSearch -> resultPtg = TextCalculator.calcSearch(operands)

            FunctionConstants.xlfSearchb -> resultPtg = TextCalculator.calcSearchB(operands)

            FunctionConstants.xlfSubstitute -> resultPtg = TextCalculator.calcSubstitute(operands)

            FunctionConstants.xlfT -> resultPtg = TextCalculator.calcT(operands)

            FunctionConstants.xlfTrim -> resultPtg = TextCalculator.calcTrim(operands)

            FunctionConstants.xlfText -> resultPtg = TextCalculator.calcText(operands)


            FunctionConstants.xlfValue -> resultPtg = TextCalculator.calcValue(operands)


            else -> {
                var s: String? = FunctionConstants.getFunctionString(functionId.toShort())
                if (s != null && s != "")
                    s = s.substring(0, s.length - 1)
                else
                    s = Integer.toHexString(functionId)
                //throw new FunctionNotSupportedException( (!.equals(""))?FunctionConstants.getFunctionString((short)funkId).substring(0, ):Integer.toHexString((int)funkId));
                throw FunctionNotSupportedException(s)    // 20081118 KSC: add a little more info ...
            }
        }// KSC: Clear out lookup caches!
        //			function.getParentRec().getWorkBook().getRefTracker().clearLookupCaches();
        // KSC: Clear out lookup caches!
        //			function.getParentRec().getWorkBook().getRefTracker().clearLookupCaches();


        return resultPtg
    }

    /****************************************************************************
     * *
     * The following section is made up of the calcuations for each of the     *
     * function types.  These map directly to the name of the function         *
     * declared in the header and is called from the coresponding switch       *
     * statement above.  ENJOY!                                                *
     * *
     */


    /****************************************
     * *
     * Excel function numbers              *
     * *
     */
    /*
	public static final int xlfCount    = 0;
	public static final int xlfIf	    = 1;
	public static final int xlfIsna     = 2;
	public static final int xlfIserror  = 3;
	public static final int xlfSum      = 4;
	public static final int xlfAverage  = 5;
	public static final int xlfMin      = 6;
	public static final int xlfMax      = 7;
	public static final int xlfRow      = 8;
	public static final int xlfColumn   = 9;
	public static final int xlfNa       = 10;
	public static final int xlfNpv      = 11;
	public static final int xlfStdev    = 12;
	public static final int xlfDollar   = 13;
	public static final int xlfFixed    = 14;
	public static final int xlfSin      = 15;
	public static final int xlfCos      = 16;
	public static final int xlfTan      = 17;
	public static final int xlfAtan     = 18;
	public static final int xlfPi       = 19;
	public static final int xlfSqrt     = 20;
	public static final int xlfExp      = 21;
	public static final int xlfLn       = 22;
	public static final int xlfLog10    = 23;
	public static final int xlfAbs      = 24;
	public static final int xlfInt      = 25;
	public static final int xlfSign     = 26;
	public static final int xlfRound    = 27;
	public static final int xlfLookup   = 28;
	public static final int xlfIndex    = 29;
	public static final int xlfRept     = 30;
	public static final int xlfMid      = 31;
	public static final int xlfLen      = 32;
	public static final int xlfValue    = 33;
	public static final int xlfTrue     = 34;
	public static final int xlfFalse    = 35;
	public static final int xlfAnd      = 36;
	public static final int xlfOr       = 37;
	public static final int xlfNot      = 38;
	public static final int xlfMod      = 39;
	public static final int xlfDcount   = 40;
	public static final int xlfDsum     = 41;
	public static final int xlfDaverage = 42;
	public static final int xlfDmin     = 43;
	public static final int xlfDmax     = 44;
	public static final int xlfDstdev   = 45;
	public static final int xlfVar      = 46;
	public static final int xlfDvar     = 47;
	public static final int xlfText     = 48;
	public static final int xlfLinest   = 49;
	public static final int xlfTrend    = 50;
	public static final int xlfLogest   = 51;
	public static final int xlfGrowth   = 52;
	public static final int xlfGoto     = 53;
	public static final int xlfHalt     = 54;
	public static final int xlfPv       = 56;
	public static final int xlfFv       = 57;
	public static final int xlfNper     = 58;
	public static final int xlfPmt      = 59;
	public static final int xlfRate     = 60;
	public static final int xlfMirr     = 61;
	public static final int xlfIrr      = 62;
	public static final int xlfRand     = 63;
	public static final int xlfMatch    = 64;
	public static final int xlfDate     = 65;
	public static final int xlfTime     = 66;
	public static final int xlfDay      = 67;
	public static final int xlfMonth    = 68;
	public static final int xlfYear     = 69;
	public static final int xlfWeekday  = 70;
	public static final int xlfHour     = 71;
	public static final int xlfMinute   = 72;
	public static final int xlfSecond   = 73;
	public static final int xlfNow      = 74;
	public static final int xlfAreas    = 75;
	public static final int xlfRows     = 76;
	public static final int xlfColumns  = 77;
	public static final int xlfOffset   = 78;
	public static final int xlfAbsref   = 79;
	public static final int xlfRelref   = 80;
	public static final int xlfArgument = 81;
	public static final int xlfSearch   = 82;
	public static final int xlfTranspose = 83;
	public static final int xlfError    = 84;
	public static final int xlfStep     = 85;
	public static final int xlfType     = 86;
	public static final int xlfEcho     = 87;
	public static final int xlfSetName  = 88;
	public static final int xlfCaller   = 89;
	public static final int xlfDeref    = 90;
	public static final int xlfWindows  = 91;
	public static final int xlfSeries   = 92;
	public static final int xlfDocuments = 93;
	public static final int xlfActiveCell = 94;
	public static final int xlfSelection = 95;
	public static final int xlfResult   = 96;
	public static final int xlfAtan2    = 97;
	public static final int xlfAsin     = 98;
	public static final int xlfAcos     = 99;
	public static final int xlfChoose   = 100;
	public static final int xlfHlookup  = 101;
	public static final int xlfVlookup  = 102;
	public static final int xlfLinks    = 103;
	public static final int xlfInput    = 104;
	public static final int xlfIsref    = 105;
	public static final int xlfGetFormula = 106;
	public static final int xlfGetName  = 107;
	public static final int xlfSetValue = 108;
	public static final int xlfLog      = 109;
	public static final int xlfExec     = 110;
	public static final int xlfChar     = 111;
	public static final int xlfLower    = 112;
	public static final int xlfUpper    = 113;
	public static final int xlfProper   = 114;
	public static final int xlfLeft     = 115;
	public static final int xlfRight    = 116;
	public static final int xlfExact    = 117;
	public static final int xlfTrim     = 118;
	public static final int xlfReplace  = 119;
	public static final int xlfSubstitute = 120;
	public static final int xlfCode     = 121;
	public static final int xlfNames    = 122;
	public static final int xlfDirectory = 123;
	public static final int xlfFind     = 124;
	public static final int xlfCell     = 125;
	public static final int xlfIserr    = 126;
	public static final int xlfIstext   = 127;
	public static final int xlfIsnumber = 128;
	public static final int xlfIsblank  = 129;
	public static final int xlfT        = 130;
	public static final int xlfN        = 131;
	public static final int xlfFopen    = 132;
	public static final int xlfFclose   = 133;
	public static final int xlfFsize    = 134;
	public static final int xlfFreadln  = 135;
	public static final int xlfFread    = 136;
	public static final int xlfFwriteln = 137;
	public static final int xlfFwrite   = 138;
	public static final int xlfFpos     = 139;
	public static final int xlfDatevalue = 140;
	public static final int xlfTimevalue = 141;
	public static final int xlfSln      = 142;
	public static final int xlfSyd      = 143;
	public static final int xlfDdb      = 144;
	public static final int xlfGetDef   = 145;
	public static final int xlfReftext  = 146;
	public static final int xlfTextref  = 147;
	public static final int xlfIndirect = 148;
	public static final int xlfRegister = 149;
	public static final int xlfCall     = 150;
	public static final int xlfAddBar   = 151;
	public static final int xlfAddMenu  = 152;
	public static final int xlfAddCommand = 153;
	public static final int xlfEnableCommand = 154;
	public static final int xlfCheckCommand = 155;
	public static final int xlfRenameCommand = 156;
	public static final int xlfShowBar  = 157;
	public static final int xlfDeleteMenu = 158;
	public static final int xlfDeleteCommand = 159;
	public static final int xlfGetChartItem = 160;
	public static final int xlfDialogBox = 161;
	public static final int xlfClean    = 162;
	public static final int xlfMdeterm  = 163;
	public static final int xlfMinverse = 164;
	public static final int xlfMmult    = 165;
	public static final int xlfFiles    = 166;
	public static final int xlfIpmt     = 167;
	public static final int xlfPpmt     = 168;
	public static final int xlfCounta   = 169;
	public static final int xlfCancelKey = 170;
	public static final int xlfInitiate = 175;
	public static final int xlfRequest  = 176;
	public static final int xlfPoke     = 177;
	public static final int xlfExecute  = 178;
	public static final int xlfTerminate = 179;
	public static final int xlfRestart  = 180;
	public static final int xlfHelp     = 181;
	public static final int xlfGetBar   = 182;
	public static final int xlfProduct  = 183;
	public static final int xlfFact     = 184;
	public static final int xlfGetCell  = 185;
	public static final int xlfGetWorkspace = 186;
	public static final int xlfGetWindow = 187;
	public static final int xlfGetDocument = 188;
	public static final int xlfDproduct = 189;
	public static final int xlfIsnontext = 190;
	public static final int xlfGetNote  = 191;
	public static final int xlfNote     = 192;
	public static final int xlfStdevp   = 193;
	public static final int xlfVarp     = 194;
	public static final int xlfDstdevp  = 195;
	public static final int xlfDvarp    = 196;
	public static final int xlfTrunc    = 197;
	public static final int xlfIslogical = 198;
	public static final int xlfDcounta  = 199;
	public static final int xlfDeleteBar = 200;
	public static final int xlfUnregister = 201;
	public static final int xlfUsdollar = 204;
	public static final int xlfFindb    = 205;
	public static final int xlfSearchb  = 206;
	public static final int xlfReplaceb = 207;
	public static final int xlfLeftb    = 208;
	public static final int xlfRightb   = 209;
	public static final int xlfMidb     = 210;
	public static final int xlfLenb     = 211;
	public static final int xlfRoundup  = 212;
	public static final int xlfRounddown = 213;
	public static final int xlfAsc      = 214;
	public static final int xlfDbcs     = 215;
	public static final int xlfRank     = 216;
	public static final int xlfAddress  = 219;
	public static final int xlfDays360  = 220;
	public static final int xlfToday    = 221;
	public static final int xlfVdb      = 222;
	public static final int xlfMedian   = 227;
	public static final int xlfSumproduct = 228;
	public static final int xlfSinh     = 229;
	public static final int xlfCosh     = 230;
	public static final int xlfTanh     = 231;
	public static final int xlfAsinh    = 232;
	public static final int xlfAcosh    = 233;
	public static final int xlfAtanh    = 234;
	public static final int xlfDget     = 235;
	public static final int xlfCreateObject = 236;
	public static final int xlfVolatile = 237;
	public static final int xlfLastError = 238;
	public static final int xlfCustomUndo = 239;
	public static final int xlfCustomRepeat = 240;
	public static final int xlfFormulaConvert = 241;
	public static final int xlfGetLinkInfo = 242;
	public static final int xlfTextBox  = 243;
	public static final int xlfInfo     = 244;
	public static final int xlfGroup    = 245;
	public static final int xlfGetObject = 246;
	public static final int xlfDb       = 247;
	public static final int xlfPause    = 248;
	public static final int xlfResume   = 251;
	public static final int xlfFrequency = 252;
	public static final int xlfAddToolbar = 253;
	public static final int xlfDeleteToolbar = 254;
	public static final int xlfADDIN	= 255;	// KSC: Added; Excel function ID for add-ins
	public static final int xlfResetToolbar = 256;
	public static final int xlfEvaluate = 257;
	public static final int xlfGetToolbar = 258;
	public static final int xlfGetTool  = 259;
	public static final int xlfSpellingCheck = 260;
	public static final int xlfErrorType = 261;
	public static final int xlfAppTitle = 262;
	public static final int xlfWindowTitle = 263;
	public static final int xlfSaveToolbar = 264;
	public static final int xlfEnableTool = 265;
	public static final int xlfPressTool = 266;
	public static final int xlfRegisterId = 267;
	public static final int xlfGetWorkbook = 268;
	public static final int xlfAvedev   = 269;
	public static final int xlfBetadist = 270;
	public static final int xlfGammaln  = 271;
	public static final int xlfBetainv  = 272;
	public static final int xlfBinomdist = 273;
	public static final int xlfChidist  = 274;
	public static final int xlfChiinv  = 275;
	public static final int xlfCombin   = 276;
	public static final int xlfConfidence = 277;
	public static final int xlfCritbinom = 278;
	public static final int xlfEven     = 279;
	public static final int xlfExpondist = 280;
	public static final int xlfFdist    = 281;
	public static final int xlfFinv     = 282;
	public static final int xlfFisher   = 283;
	public static final int xlfFisherinv = 284;
	public static final int xlfFloor    = 285;
	public static final int xlfGammadist =  286;
	public static final int xlfGammainv     = 287;
	public static final int xlfCeiling  = 288;
	public static final int xlfHypgeomdist = 289;
	public static final int xlfLognormdist = 290;
	public static final int xlfLoginv   = 291;
	public static final int xlfNegbinomdist = 292;
	public static final int xlfNormdist = 293;
	public static final int xlfNormsdist    = 294;
	public static final int xlfNorminv  = 295;
	public static final int xlfNormsinv = 296;
	public static final int xlfStandardize = 297;
	public static final int xlfOdd      = 298;
	public static final int xlfPermut   = 299;
	public static final int xlfPoisson  = 300;
	public static final int xlfTdist    = 301;
	public static final int xlfWeibull  = 302;
	public static final int xlfSumxmy2  = 303;
	public static final int xlfSumx2my2 = 304;
	public static final int xlfSumx2py2 = 305;
	public static final int xlfChitest  = 306;
	public static final int xlfCorrel   = 307;
	public static final int xlfCovar    = 308;
	public static final int xlfForecast = 309;
	public static final int xlfFtest    = 310;
	public static final int xlfIntercept = 311;
	public static final int xlfPearson  = 312;
	public static final int xlfRsq      = 313;
	public static final int xlfSteyx    = 314;
	public static final int xlfSlope    = 315;
	public static final int xlfTtest    = 316;
	public static final int xlfProb     = 317;
	public static final int xlfDevsq    = 318;
	public static final int xlfGeomean  = 319;
	public static final int xlfHarmean  = 320;
	public static final int xlfSumsq    = 321;
	public static final int xlfKurt     = 322;
	public static final int xlfSkew     = 323;
	public static final int xlfZtest    = 324;
	public static final int xlfLarge    = 325;
	public static final int xlfSmall    = 326;
	public static final int xlfQuartile = 327;
	public static final int xlfPercentile = 328;
	public static final int xlfPercentrank = 329;
	public static final int xlfMode     = 330;
	public static final int xlfTrimmean = 331;
	public static final int xlfTinv     = 332;
	public static final int xlfMovieCommand = 334;
	public static final int xlfGetMovie = 335;
	public static final int xlfConcatenate = 336;
	public static final int xlfPower    = 337;
	public static final int xlfPivotAddData = 338;
	public static final int xlfGetPivotTable = 339;
	public static final int xlfGetPivotField = 340;
	public static final int xlfGetPivotItem = 341;
	public static final int xlfRadians  = 342;
	public static final int xlfDegrees  = 343;
	public static final int xlfSubtotal = 344;
	public static final int xlfSumif    = 345;
	public static final int xlfCountif  = 346;
	public static final int xlfCountblank = 347;
	public static final int xlfScenarioGet = 348;
	public static final int xlfOptionsListsGet = 349;
	public static final int xlfIspmt    = 350;
	public static final int xlfDatedif  = 351;
	public static final int xlfDatestring = 352;
	public static final  int xlfNumberstring = 353;
	public static final int xlfRoman    = 354;
	public static final int xlfOpenDialog = 355;
	public static final int xlfSaveDialog = 356;
	public static final int xlfViewGet  = 357;
	public static final int xlfGetPivotData = 358;
	public static final int xlfHyperlink = 359;
	public static final int xlfPhonetic     = 360;
	public static final int xlfAverageA     = 361;
	public static final int xlfMaxA     = 362;
	public static final int xlfMinA     = 363;
	public static final int xlfStDevPA  = 364;
	public static final int xlfVarPA    = 365;
	public static final int xlfStDevA   = 366;
	public static final int xlfVarA     = 367;
	// KSC: ADD-IN formulas - use any index; name must be present in FunctionConstants.addIns
	// Financial Formulas
	public static final int xlfAccrintm= 368;
	public static final int xlfAccrint= 369;
	public static final int xlfCoupDayBS= 370;
	public static final int xlfCoupDays= 371;
	public static final int xlfCumIPmt= 372;
	public static final int xlfCumPrinc= 373;
	public static final int xlfCoupNCD= 374;
	public static final int xlfCoupDaysNC= 375;
	public static final int xlfCoupPCD= 376;
	public static final int xlfCoupNUM= 377;
	public static final int xlfDollarDE= 378;
	public static final int xlfDollarFR= 379;
	public static final int xlfEffect= 380;
	public static final int xlfINTRATE= 381;
	public static final int xlfXIRR= 382;
	public static final int xlfXNPV= 383;
	public static final int xlfYIELD= 384;
	public static final int xlfPRICE= 385;
	public static final int xlfPRICEDISC= 386;
	public static final int xlfPRICEMAT= 387;
	public static final int xlfDURATION= 388;
	public static final int xlfMDURATION= 389;
	public static final int xlfTBillEq= 390;
	public static final int xlfTBillPrice= 391;
	public static final int xlfTBillYield= 392;
	public static final int xlfYieldDisc= 393;
	public static final int xlfYieldMat= 394;
	public static final int xlfFVSchedule= 395;
	public static final int xlfAmorlinc= 396;
	public static final int xlfAmordegrc= 397;
	public static final int xlfOddFPrice= 398;
	public static final int xlfOddLPrice= 399;
	public static final int xlfOddFYield= 400;
	public static final int xlfOddLYield= 401;
	public static final int xlfNOMINAL= 402;
	public static final int xlfDISC= 403;
	public static final int xlfRECEIVED= 404;
	// Engineering Formulas
	public static final int xlfBIN2DEC= 405; 
	public static final int xlfBIN2HEX= 406; 
	public static final int xlfBIN2OCT= 407; 
	public static final int xlfDEC2BIN= 408; 
	public static final int xlfDEC2HEX= 409; 
	public static final int xlfDEC2OCT= 410; 
	public static final int xlfHEX2BIN= 411; 
	public static final int xlfHEX2DEC= 412; 
	public static final int xlfHEX2OCT= 413; 
	public static final int xlfOCT2BIN= 414; 
	public static final int xlfOCT2DEC= 415; 
	public static final int xlfOCT2HEX= 416;
	public static final int xlfCOMPLEX= 417;
	public static final int xlfGESTEP= 	418;
	public static final int xlfDELTA= 	419;
	public static final int xlfIMAGINARY= 420;
	public static final int xlfIMABS=	421;
	public static final int xlfIMDIV=	422;
	public static final int xlfIMCONJUGATE= 423;
	public static final int xlfIMCOS=	424;
	public static final int xlfIMSIN=	425;
	public static final int xlfIMREAL=	426;
	public static final int xlfIMEXP=	427;
	public static final int xlfIMSUB=	428;
	public static final int xlfIMSUM=	429;
	public static final int xlfIMPRODUCT= 430;
	public static final int xlfIMLN=	431;
	public static final int xlfIMLOG10= 432;
	public static final int xlfIMLOG2=	433;
	public static final int xlfIMPOWER=	434;
	public static final int xlfIMSQRT=	435;
	public static final int xlfIMARGUMENT= 436;
	public static final int xlfCONVERT= 437;
	// Math Add-in Formulas
	public static final int xlfDOUBLEFACT= 438;
	public static final int xlfGCD=		439;
	public static final int xlfLCM=		440;
	public static final int xlfMROUND=	441;
	public static final int xlfMULTINOMIAL= 442;
	public static final int xlfQUOTIENT=	443;
	public static final int xlfRANDBETWEEN= 444;
	public static final int xlfSERIESSUM=	445;
	public static final int xlfSQRTPI=		446;
*/
}