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
import io.starter.toolkit.Logger

import java.util.Stack


/**
 * Formula Calculator.
 *
 *
 * Translates an excel calc stack into a value.  The stack is in a modified
 * reverse polish notation.  For details please look into the Excel 97 reference for BIFF8.
 *
 *
 * Formula Calculators do not exist per formula, rather it uses a factory pattern.
 * Put a formula in & calculate it.
 *
 *
 * Actual calculation methods exist within the OperatorPtg's themselves, so this just handles
 * the grunt work of parsing and passing ptgs back and forth along with formatting the output.
 *
 * @see Formula
 */
object FormulaCalculator {

    private val DEBUG = false // just use this to see ptg calcing

    /**
     * Calculates the value of calcStac  This is handled by
     * running through the stack, adding operands to tempstack until
     * an operator PTG is found. At that point pass the relevant ptg's from
     * tempstack into the calculate method of the operator PTG.  The operator
     * ptg should return a valid value PTG.
     */
    @Throws(FunctionNotSupportedException::class)
    fun calculateFormula(expression: Stack<*>): Any {
        val sz = expression.size
        var stck = arrayOfNulls<Ptg>(sz)
        stck = expression.toTypedArray() as Array<Ptg>
        val calcStack = Stack()
        for (t in 0 until sz) { // flip the stack TODO: investigate why needed
            calcStack.add(0, stck[t])
        }
        val tempstack = Stack()
        while (!calcStack.isEmpty()) {// loop while there are Ptgs
            handlePtg(calcStack, tempstack)
        }
        val finalptg = tempstack.pop() as Ptg
        return finalptg.value
    }


    /**
     * Calculates the final Ptg result of calcStack.  This is handled by
     * running through the stack, adding operands to tempstack until
     * an operator PTG is found. At that point pass the relevant ptg's from
     * tempstack into the calculate method of the operator PTG.  The operator
     * ptg should return a valid value PTG.
     */
    @Throws(FunctionNotSupportedException::class)
    fun calculateFormulaPtg(expression: Stack<*>): Ptg {
        val sz = expression.size
        var stck = arrayOfNulls<Ptg>(sz)
        stck = expression.toTypedArray() as Array<Ptg>
        val calcStack = Stack()
        for (t in 0 until sz) { // flip the stack TODO: investigate why needed
            calcStack.add(0, stck[t])
        }
        val tempstack = Stack()
        while (!calcStack.isEmpty()) {// loop while there are Ptgs
            handlePtg(calcStack, tempstack)
        }
        return tempstack.pop()
    }

    /**
     * This is a very similar method to the handle ptg method in formula parser.
     * Instead of creating a tree however it calculates in the order recommended by
     * the book of knowledge (excel developers guide).  That is, FILO.  First In Last Out.
     * We also don't really care about things like parens, they are just for display purposes.
     */
    @Throws(FunctionNotSupportedException::class)
    internal fun handlePtg(newstck: Stack<*>, vals: Stack<*>) {
        var p = newstck.pop() as Ptg
        var x = 0
        var t = 0
        if (p.isOperator || p.isControl || p.isFunction) {
            // Get rid of the parens ptgs
            if (p.isControl && !vals.isEmpty()) {
                if (p.opcode.toInt() == 0x15) { // its a parens!
                    // the parens is already pop'd so just return and it is gone...
                    return
                }
                // we didn't use it, back it goes.
                if (DEBUG) Logger.logInfo("opr: $p")
            }
            // make sure we have the correct amount popped back in..
            if (p.isBinaryOperator) t = 2
            if (p.isUnaryOperator) t = 1
            if (p.isStandAloneOperator) t = 0
            if (p.opcode.toInt() == 0x22 || p.opcode.toInt() == 0x42 || p.opcode.toInt() == 0x62) {
                t = p.numParams
            }// it's a ptgfunkvar!
            if (p.opcode.toInt() == 0x21 || p.opcode.toInt() == 0x41 || p.opcode.toInt() == 0x61) {
                t = p.numParams
            }// guess that ptgfunc is not only one..

            val vx = arrayOfNulls<Ptg>(t)
            while (x < t) {
                vx[t - 1 - x] = vals.pop() as Ptg
                x++
            }// get'em

            // QUITE AN IMPORTANT LINE... FYI. -jm
            try {
                p = p.calculatePtg(vx)
            } catch (e: CalculationException) {
                p = PtgErr(e.errorCode)
                if (e.name == "#CIR_ERR!") {
                    p.isCircularError = true
                }
            }

            /* useful for debugging*/
            if (DEBUG) {
                var adr = ""
                if (p.parentRec != null)
                    adr = "addr: " + p.parentRec.cellAddress
                Logger.logInfo("$adr val: $p")
            }
            vals.push(p)// push it back on the stack

        } else if (p.isOperand) {

            if (DEBUG) Logger.logInfo("opr: $p")

            vals.push(p)

        } else if (p is PtgAtr) {

            // this is probably just a space at this point, don't output error message

        } else {
            throw FunctionNotSupportedException("WARNING: Calculating Formula failed: Unsupported/Incorrect Ptg Type: 0x" + p.opcode + " " + p.string)
        }
    }
}