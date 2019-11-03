package io.starter.OpenXLS

/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */

import io.starter.formats.XLS.CellNotFoundException
import io.starter.formats.XLS.FormulaNotFoundException
import io.starter.formats.XLS.FunctionNotSupportedException
import io.starter.formats.XLS.WorkSheetNotFoundException
import io.starter.toolkit.Logger
import org.junit.Test

import java.io.BufferedOutputStream
import java.io.File
import java.io.FileOutputStream

/**
 * This Class Demonstrates the functionality of of OpenXLS Formula manipulation.
 */

class FormulaTest {

    @Test
    fun testFormulaParse() {
        val t = testformula()
        t.testFormula()
    }

    @Test
    fun testHandlerFunctions() {
        val t = testformula()
        t.testHandlerFunctions()
    }

    @Test
    fun testMultiChange() {
        val t = testformula()
        t.testMultiChange()
    }

    @Test
    fun changeSingleCellLoc() {
        val t = testformula()
        t.changeSingleCellLoc()
    }

}

/**
 * Test the manipulation of Formulas within a worksheet.
 */
internal class testformula {
    var book: WorkBookHandle? = null
    var sheet: WorkSheetHandle? = null
    var sheetname = "Sheet1"
    var wd = System.getProperty("user.dir") + "/docs/samples/Formulas/"
    var finpath = wd + "testFormula.xls"

    var sht: WorkSheetHandle? = null

    /**
     * thrash multiple changes to formula references and recalc
     * ------------------------------------------------------------
     */
    fun testMultiChange() {
        try {
            Logger.logInfo("Testing multiple changes to formula references and recalc")
            val wbx = WorkBookHandle()
            val sheet1 = wbx.getWorkSheet(0)
            sheet1.add(100.123, "A1")
            sheet1.add(200.123, "A2")
            val cx = sheet1.add("=sum(A1*A2)", "A3")
            Logger.logInfo(cx.toString())
            Logger.logInfo("start setting 100k vals")
            for (t in 0..99999) {
                sheet1.getCell("A1").setVal(Math.random() * 10000)
                sheet1.getCell("A2").setVal(Math.random() * 10000)
                val calced = cx!!.`val`
                Logger.logInfo(calced!!.toString())
            }
            Logger.logInfo("done setting 100k vals")
            wbx.write(FileOutputStream(File(
                    wd + "testFormulas_out.xls")), WorkBookHandle.FORMAT_XLS)

        } catch (ex: Exception) {
            Logger.logErr("testFormulas.testMultiChange: $ex")
        }

    }

    /**
     * Demonstrates Dynamic Formula Calculation
     */
    fun testCalculation() {
        try {
            this.openSheet(finpath, sheetname)
            // c4 + d4 = f4
            val mycell1 = sheet!!.getCell("C4")
            val mycell2 = sheet!!.getCell("D4")
            val myformulacell = sheet!!.getCell("F4")

            // output the calculated values
            val form = myformulacell.formulaHandle
            io.starter.toolkit.Logger.log(form.calculate()!!.toString())

            // change the values then recalc
            mycell1.setVal(99)
            mycell2.setVal(420)
            io.starter.toolkit.Logger.log(form.calculate()!!.toString())

            testWrite("testCalculation_out.xls")
        } catch (e: CellNotFoundException) {
            io.starter.toolkit.Logger.log("cell not found$e")
        } catch (e: FormulaNotFoundException) {
            io.starter.toolkit.Logger.log("No formula to change$e")
        } catch (e: Exception) {
            Logger.logErr("TestFormulas failed.", e)
        }

    }

    /**
     * Move a Cell Reference within a Formula
     */
    fun changeSingleCellLoc() {
        try {
            this.openSheet(finpath, sheetname)
            val mycell = sheet!!.getCell("A10")
            val form = mycell.formulaHandle
            form.changeFormulaLocation("A3", "G10")
            testWrite("testChangeSingleCellLoc_out.xls")
        } catch (e: CellNotFoundException) {
            io.starter.toolkit.Logger.log("cell not found$e")
        } catch (e: FormulaNotFoundException) {
            io.starter.toolkit.Logger.log("No formula to change$e")
        }

    }

    /**
     * Move a Cell range reference within a Formula
     */
    fun testHandlerFunctions() {
        try {
            this.openSheet(finpath, sheetname)
            val mycell = sheet!!.getCell("E8")
            val myhandle = mycell.formulaHandle
            val b = myhandle.changeFormulaLocation("A1:B2", "D1:D28")
            testWrite("testHandlerFunctions_out.xls")
        } catch (e: CellNotFoundException) {
            io.starter.toolkit.Logger.log("cell not found$e")
        } catch (e: FormulaNotFoundException) {
            io.starter.toolkit.Logger.log("No formula to change$e")
        }

    }

    /**
     * Add a cell to a Cell range reference within a Formula
     */
    fun testCellHandlerFunctions() {
        try {
            this.openSheet(finpath, sheetname)
            val mycell = sheet!!.getCell("E8")
            val secondcell = sheet!!.getCell("D19")
            val myhandle = mycell.formulaHandle
            val b = myhandle.addCellToRange("A1:B2", secondcell)
            testWrite("testCellHandlerFunctions_out.xls")
        } catch (e: CellNotFoundException) {
            io.starter.toolkit.Logger.log("cell not found$e")
        } catch (e: FormulaNotFoundException) {
            io.starter.toolkit.Logger.log("No formula to change$e")
        }

    }

    /**
     * Run tests
     */
    fun testFormula() {
        try {
            val finpath = wd + "testFormula.xls"
            val sheetname = "Sheet1"
            this.openSheet(finpath, sheetname)
            sheet!!.removeRow(2, true)
            testWrite("testFormula_out.xls")
        } catch (e: Exception) {
            io.starter.toolkit.Logger
                    .log("Exception in testFORMULA.testFormulaSeries(): $e")
        }

    }

    /**
     * Demonstrates calculation of formulas
     *
     *
     * Jan 19, 2010
     *
     * @param fs
     * @param sh
     */
    fun testFormulaCalc(fs: String, sh: String) {
        val book = WorkBookHandle(fs)
        sheetname = sh
        try {
            sht = book.getWorkSheet(sheetname)
        } catch (e: Exception) {
            Logger.logErr("TestFormulas failed.", e)
        }

        var f: FormulaHandle? = null
        var i: Double? = null

        /************************************
         * Formula Parse test
         */
        if (sheetname.equals("Sheet1", ignoreCase = true)) {
            try {

                // one ref & ptgadd
                sht!!.add(null, "A1")
                var c = sht!!.getCell("A1")
                c.setFormula("b1+5")
                f = c.formulaHandle
                i = f!!.calculate() as Double

                // two refs & ptgadd
                sht!!.add(null, "A2")
                c = sht!!.getCell("A2")
                c.setFormula("B1+ A1")
                f = c.formulaHandle
                i = f!!.calculate() as Double

                // ptgsub
                f.setFormula("B1 - 5")
                i = f.calculate() as Double

                // ptgmul
                f.setFormula("D1 * F1")
                i = f.calculate() as Double

                // ptgdiv
                f.setFormula("E1 / F1")
                i = f.calculate() as Double

                // ptgpower
                f.setFormula("E1 ^ F1")
                i = f.calculate() as Double

                f.setFormula("E1 > F1")
                var b = f.calculate() as Boolean

                f.setFormula("E1 >= F1")
                b = f.calculate() as Boolean

                f.setFormula("E1 < F1")
                b = f.calculate() as Boolean

                f.setFormula("E1 <= F1")
                b = f.calculate() as Boolean

                f.setFormula("Pi()")
                i = f.calculate() as Double

                f.setFormula("LOG(10,2)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                f.setFormula("ROUND(32.443,1)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                f.setFormula("MOD(45,6)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                f.setFormula("DATE(1998,2,4)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                f.setFormula("SUM(1998,2,4)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                f.setFormula("IF(TRUE,1,0)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                f.setFormula("ISERR(\"test\")")
                b = f.calculate() as Boolean
                io.starter.toolkit.Logger.log(b.toString())

                // many operand ptgfuncvar
                f.setFormula("SUM(12,3,2,4,5,1)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                // test with a sub-calc
                f.setFormula("IF((1<2),1,0)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                f.setFormula("IF((1<2),MOD(45,6),1)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                f.setFormula("IF((1<2),if((true),8,1),1)")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

                f.setFormula("IF((SUM(23,2,3,4)<12),if((true),8,1),DATE(1998,2,4))")
                i = f.calculate() as Double
                io.starter.toolkit.Logger.log(i.toString())

            } catch (e: CellNotFoundException) {
                Logger.logErr("TestFormulas failed.", e)
            } catch (e: FunctionNotSupportedException) {
                Logger.logErr("TestFormulas failed.", e)
            } catch (e: Exception) {
                Logger.logErr("TestFormulas failed.", e)
            }

            testWrite("testCalcFormulas_out.xls")
        }
    }

    fun openSheet(finp: String, sheetnm: String) {
        book = WorkBookHandle(finp)
        try {
            sheet = book!!.getWorkSheet(sheetnm)
        } catch (e: WorkSheetNotFoundException) {
            io.starter.toolkit.Logger.log("couldn't find worksheet$e")
        }

    }

    fun testWrite(fname: String) {
        try {
            val f = java.io.File(wd + fname)
            val fos = FileOutputStream(f)
            val bbout = BufferedOutputStream(fos)
            book!!.write(bbout)
            bbout.flush()
            fos.close()
        } catch (e: java.io.IOException) {
            Logger.logInfo("IOException in Tester.  $e")
        }

    }

}