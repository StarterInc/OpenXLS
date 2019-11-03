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
package io.starter.formats.XLS

import io.starter.formats.XLS.formulas.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger
import java.util.Stack
import java.util.Vector

/**
 *
 */
class ExpressionParser : java.io.Serializable {
    companion object {
        /**
         *
         */
        private const val serialVersionUID = 4745215965823234010L
        private val DEBUGLEVEL = 0
        /*
     * All of the operand values
     *
     * Section of binary operator PTG's. These pop the two
     * top values out of a stack and perform an operation on
     * them before pushing back in
     */
        // really "special" one, read all about it.
        val ptgExp: Short = 0x1
        val ptgAdd: Short = 0x3
        val ptgSub: Short = 0x4
        val ptgMlt: Short = 0x5
        val ptgDiv: Short = 0x6
        val ptgPower: Short = 0x7
        val ptgConcat: Short = 0x8
        val ptgLT: Short = 0x09
        val ptgLE: Short = 0x0a
        val ptgEQ: Short = 0x0b
        val ptgGE: Short = 0x0c
        val ptgGT: Short = 0x0d
        val ptgNE: Short = 0x0e
        val ptgIsect: Short = 0x0f
        val ptgUnion: Short = 0x10
        val ptgRange: Short = 0x11
        // End of binary operator PTG's

        // Unary Operator tokens
        val ptgUPlus: Short = 0x12
        val ptgUMinus: Short = 0x13                    // todo
        val ptgPercent: Short = 0x14                    // todo

        // Controls
        val ptgParen: Short = 0x15
        val ptgAtr: Short = 0x19
        // End of Controls

        // Constant operators
        val ptgMissArg: Short = 0x16
        val ptgStr: Short = 0x17
        val ptgEndSheet: Short = 0x1b
        val ptgErr: Short = 0x1c
        val ptgBool: Short = 0x1d
        val ptgInt: Short = 0x1e
        val ptgNum: Short = 0x1f
        // End of Constant Operators

        val ptgArray: Short = 0x20
        val ptgFunc: Short = 0x21
        val ptgFuncVar: Short = 0x22
        val ptgName: Short = 0x23
        val ptgRef: Short = 0x24
        val ptgArea: Short = 0x25
        val ptgMemArea: Short = 0x26
        val ptgMemErr: Short = 0x27
        val ptgMemFunc: Short = 0x29
        val ptgRefErr: Short = 0x2a
        val ptgAreaErr: Short = 0x2b
        val ptgRefN: Short = 0x2c
        val ptgAreaN: Short = 0x2d
        val ptgNameX: Short = 0x39
        val ptgRef3d: Short = 0x3a
        val ptgArea3d: Short = 0x3b
        val ptgRefErr3d: Short = 0x3c

        // who knows, added to fix broken Named ranges -jm 03/26/04
        val ptgAreaErr3d: Short = 0x3d
        val ptgMemAreaA: Short = 0x66
        val ptgMemAreaNV: Short = 0x4e
        val ptgMemAreaN: Short = 0x2e

        /**
         * Parse the byte array, create component Ptg's and insert
         * them into a stack.
         *
         *
         *
         *
         * Feb 8, 2010
         *
         * @param function
         * @param rec
         * @return
         */
        fun parseExpression(function: ByteArray, rec: XLSRecord): Stack<*> {
            return ExpressionParser.parseExpression(function, rec, function.size)
        }

        /**
         * Parse the byte array, create component Ptg's and insert them into
         * a stack.
         *
         *
         * Feb 8, 2010
         *
         * @param function
         * @param rec
         * @param expressionLen
         * @return
         */
        fun parseExpression(function: ByteArray, rec: XLSRecord, expressionLen: Int): Stack<*> {
            var expressionLen = expressionLen
            val stack = Stack()
            var ptg: Short = 0x0
            var ptgLen = 0
            var hasArrays = false
            /*
         * Not really needed
         * //boolean hasPtgExtraMem= false;
         * //PtgMemArea pma= null;
         */
            val arrayLocs = CompatibleVector()
            if (expressionLen > function.size)
                expressionLen = function.size // deal with out of spec formulas
            // (testJapanese:Open25.xls) -jm
            // KSC: shared formula changes for peformance: now
            // PtgRefN's/PtgAreaN's are instantiated and
            // reference-tracked (of a sort) ...

            // iterate the expression and create Ptgs.
            run {
                var i = 0
                while (i < expressionLen) {
                    // check if the 40 bit is set, is it a Array class?
                    if (function[i] and 0x40 == 0x40) {
                        // rec is a value class
                        // we need to strip the high-order bits and set the 0x20 bit
                        ptg = (function[i] or 0x20 and 0x3f).toShort()
                    } else {
                        // the bit is already set, just strip the high order bits
                        // rec may be an array class. need to figure rec one out.
                        ptg = (function[i] and 0x3f).toShort()
                    }
                    when (ptg) {

                        ptgExp -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgExp Located")
                            if (i == 0) {// MUST BE THE ONLY PTG in the formula expression
                                val px = PtgExp()
                                ptgLen = px.length
                                val b = ByteArray(ptgLen)
                                if (ptgLen + i <= function.size)
                                    System.arraycopy(function, i, b, 0, ptgLen)
                                px.parentRec = rec
                                px.init(b)
                                stack.push(px)
                                break
                            }
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgStr Located")
                            var x = i
                            x += 1 // move past the opcode to the cch
                            ptgLen = function[x] and 0xff // this is the cch
                            val theGrbit = function[x + 1].toShort()// this is the grbit;
                            if (theGrbit and 0x1 == 0x1) {
                                // unicode string
                                ptgLen = ptgLen * 2
                            }
                            ptgLen += 3 // include the PtgId, cch, & grbit;
                            val pst = PtgStr()
                            var b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pst.init(b)
                            pst.parentRec = rec
                            stack.push(pst)
                        }
                        // ptgStr is one of the only ptg's that varies in length, so
                        // there is some special handling
                        // going on for it.
                        ptgStr -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgStr Located")
                            var x = i
                            x += 1
                            ptgLen = function[x] and 0xff
                            val theGrbit = function[x + 1].toShort()
                            if (theGrbit and 0x1 == 0x1) {
                                ptgLen = ptgLen * 2
                            }
                            ptgLen += 3
                            val pst = PtgStr()
                            var b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pst.init(b)
                            pst.parentRec = rec
                            stack.push(pst)
                        }
                        /* */

                        ptgMemAreaA -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgMemAreaA Located" + function[i])
                            x = i
                            x += 5 // move past the opcode & reserved to the cce
                            ptgLen = ByteTools.readShort(function[x].toInt(), function[x + 1].toInt()).toInt() // this
                            // is
                            // the
                            // cce
                            ptgLen += 7 // include the PtgId, cce, & reserv;
                            val pmema = PtgMemAreaA()
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pmema.init(b)
                            pmema.parentRec = rec
                            stack.push(pmema)
                        }

                        ptgMemAreaN -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgMemAreaN Located" + function[i])
                            val pmemn = PtgMemAreaN()
                            ptgLen = pmemn.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pmemn.init(b)
                            pmemn.parentRec = rec
                            stack.push(pmemn)
                        }

                        ptgMemAreaNV -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgMemAreaNV Located" + function[i])
                            x = i
                            x += 5 // move past the opcode & reserved to the cce
                            ptgLen = ByteTools.readShort(function[x].toInt(), function[x + 1].toInt()).toInt() // this
                            // is
                            // the
                            // cce
                            ptgLen += 7 // include the PtgId, cce, & reserv;
                            val pmemv = PtgMemAreaNV()
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pmemv.init(b)
                            pmemv.parentRec = rec
                            stack.push(pmemv)
                        }

                        // ptgMemArea also varies in length...
                        ptgMemArea -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgMemArea Located" + function[i])
                            ptgLen = 7
                            val pmem = PtgMemArea()
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pmem.init(b)
                            // now grab the rest of the "extra data" that defines the
                            // ptgmemarea
                            // these are separate ptgs (PtgArea, PtgRef's ... plus
                            // PtgUnions)
                            // that comprise the PtgMemArea coordinates
                            pmem.parentRec = rec
                            i += ptgLen // after PtgMemArea record, get subexpression
                            ptgLen = pmem.getnTokens()
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pmem.setSubExpression(b)
                            stack.push(pmem)
                        }

                        ptgMemFunc -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgMemFunc Located")
                            val pmemf = PtgMemFunc()
                            x = i
                            x += 1 // move past the opcode to the cce
                            ptgLen = ByteTools.readShort(function[x].toInt(), function[x + 1].toInt()).toInt() // this
                            // is
                            // the
                            // cce
                            ptgLen += 3 // include the PtgId, cce, & reserv;
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pmemf.parentRec = rec
                            pmemf.init(b)
                            stack.push(pmemf)
                        }

                        ptgInt -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgInt Located")
                            val pi = PtgInt()
                            ptgLen = pi.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pi.init(b)
                            pi.parentRec = rec
                            stack.push(pi)
                        }

                        ptgErr -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgErr Located")
                            val perr = PtgErr()
                            ptgLen = perr.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            perr.init(b)
                            perr.parentRec = rec
                            stack.push(perr)
                        }

                        ptgNum -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgNum Located")
                            val pnum = PtgNumber()
                            ptgLen = pnum.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pnum.init(b)
                            pnum.parentRec = rec
                            stack.push(pnum)
                        }

                        ptgBool -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgBool Located")
                            val pboo = PtgBool()
                            ptgLen = pboo.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pboo.init(b)
                            pboo.parentRec = rec
                            stack.push(pboo)
                        }

                        ptgName -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgName Located")
                            val pn = PtgName()
                            ptgLen = pn.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pn.parentRec = rec
                            pn.init(b)
                            pn.addListener()
                            stack.push(pn)
                            val chk = i + ptgLen
                            if (chk < function.size) {
                                if (function[i + ptgLen].toInt() == 0x0) {
                                    if (DEBUGLEVEL > 1)
                                        Logger.logWarn("Undocumented Name Record mystery byte encountered in Formula: ")
                                    i++
                                }
                            }
                        }

                        ptgNameX -> {
                            if (DEBUGLEVEL > 1)
                                Logger.logInfo("ptgNameX Located")
                            if (DEBUGLEVEL > 0)
                                Logger.logWarn("referencing external spreadsheets unsupported.")
                            val pnx = PtgNameX()
                            ptgLen = pnx.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pnx.init(b)
                            pnx.parentRec = rec
                            pnx.addListener()
                            stack.push(pnx)
                        }

                        ptgRef -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgRef Located ")
                            val pt = PtgRef()
                            ptgLen = pt.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pt.parentRec = rec // parent rec must be set before init
                            pt.init(b)
                            pt.addToRefTracker()
                            stack.push(pt)
                        }

                        ptgArray -> {
                            hasArrays = true
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgArray Located ")
                            val pa = PtgArray()
                            ptgLen = 8 // 7 len + id
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pa.init(b) // setArrVals(b); // 20090820 KSC: b represents base
                            // record not array values
                            val ingr = Integer.valueOf(stack.size) // constant value
                            // array for
                            // PtgArray
                            // appears at
                            // end of stack
                            // see hasArrays
                            // below
                            arrayLocs.add(ingr)
                            stack.push(pa)
                        }

                        ptgRefN -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgRefN Located ")
                            val ptn = PtgRefN(false)
                            ptgLen = ptn.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            ptn.parentRec = rec // parent rec must be set before init
                            ptn.init(b)
                            if (rec.opcode == XLSConstants.SHRFMLA)
                                ptn.addToRefTracker()
                            stack.push(ptn)
                        }

                        ptgArea -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgArea Located ")
                            val pg = PtgArea()
                            ptgLen = pg.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pg.parentRec = rec // parent rec must be set before init
                            pg.init(b)
                            pg.addToRefTracker()
                            stack.push(pg)
                        }

                        ptgArea3d -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgArea3d Located ")
                            val pg3 = PtgArea3d()
                            ptgLen = pg3.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pg3.init(b, rec) // we need this to init the sub-ptgs correctly
                            pg3.addListener()
                            pg3.addToRefTracker()
                            stack.push(pg3)
                        }

                        ptgAreaN -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgAreaN Located ")
                            val pgn = PtgAreaN()
                            ptgLen = pgn.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pgn.parentRec = rec
                            pgn.init(b)
                            if (rec.opcode == XLSConstants.SHRFMLA) {
                                pgn.addToRefTracker()
                            }
                            stack.push(pgn)
                        }

                        ptgAreaErr3d -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgAreaErr3d Located")
                            val ptfa = PtgAreaErr3d()
                            ptgLen = ptfa.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            ptfa.parentRec = rec
                            ptfa.init(b)
                            // ptfa.addToRefTracker();
                            stack.push(ptfa)
                        }

                        ptgRefErr3d -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgRefErr3d Located")
                            val ptfr = PtgRefErr3d()
                            ptgLen = ptfr.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            ptfr.parentRec = rec
                            ptfr.init(b)
                            // ptfr.addToRefTracker();
                            stack.push(ptfr)
                        }

                        ptgMemErr -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgMemErr Located")
                            val pm = PtgMemErr()
                            ptgLen = pm.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pm.parentRec = rec
                            pm.init(b)
                            stack.push(pm)
                        }

                        ptgRefErr -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgRefErr Located")
                            val pr = PtgRefErr()
                            ptgLen = pr.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pr.parentRec = rec // parent rec must be set before init
                            pr.init(b)
                            stack.push(pr)
                        }

                        ptgEndSheet -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgEndSheet Located")
                            val prs = PtgEndSheet()
                            ptgLen = prs.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            prs.init(b)
                            prs.parentRec = rec
                            stack.push(prs)
                        }

                        ptgRef3d -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgRef3d Located")
                            val pr3 = PtgRef3d()
                            ptgLen = pr3.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pr3.parentRec = rec
                            pr3.init(b)
                            pr3.addListener()
                            pr3.addToRefTracker()
                            stack.push(pr3)
                            // if an External Link i.e. defined in another workbook,
                            // flag formula as such
                            if (pr3.isExternalLink && rec.opcode == XLSConstants.FORMULA)
                                (rec as Formula).setIsExternalRef(true)
                        }
                        /*
                 * PtgAtr is another one of the ugly size-changing ptg's
                 */
                        ptgAtr -> {
                            val pat = PtgAtr(0.toByte())
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgAtr Located")
                            ptgLen = pat.length
                            if (function[i + 1] and 0x4 == 0x4) {
                                ptgLen = ByteTools
                                        .readShort(function[i + 2].toInt(), function[i + 3].toInt()).toInt()
                                ptgLen++ // one extra for some undocumented reason
                                ptgLen = ptgLen * 2 // seems to be two bytes per...
                                ptgLen += 4 // add the cch & grbit
                            }
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pat.init(b)
                            pat.init()
                            pat.parentRec = rec
                            stack.push(pat)
                        }

                        ptgFunc -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgFunc Located")
                            val ptf = PtgFunc()
                            ptgLen = ptf.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            ptf.init(b)
                            ptf.parentRec = rec
                            stack.push(ptf)
                        }

                        ptgFuncVar -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgFuncVar Located")
                            val ptfv = PtgFuncVar()
                            ptgLen = ptfv.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)

                            ptfv.init(b)
                            ptfv.parentRec = rec
                            if (ptfv.functionId.toInt() == FunctionConstants.XLF_INDIRECT) {
                                /*
                         * TESTING NEW WAY:
                         * // New way does not account for expanded shared formula
                         * references, unfortunately so keep original for new
                         *
                         * Stack indirectStack= new Stack();
                         * int z= stack.size()-1;
                         * int nparams= 1;
                         * for (; z > 0 && nparams > 0; z--) {
                         * Ptg p= (Ptg) stack.get(z);
                         * if (p instanceof PtgAtr) {
                         * continue;
                         * }
                         * if(p.getIsOperator()||p.getIsControl()||p.getIsFunction()
                         * ){
                         * if(p.getIsControl() ){
                         * if(p.getOpcode() == 0x15) { // its a parens!
                         * // the parens is already pop'd so just return and it is
                         * gone...
                         * continue;
                         * }
                         * }
                         * int t= 0;
                         * // make sure we have the correct amount popped back in..
                         * if (p.getIsBinaryOperator()) t=2;
                         * if (p.getIsUnaryOperator()) t=1;
                         * if (p.getIsStandAloneOperator()) t=0;
                         * if (p.getOpcode() == 0x22 || p.getOpcode() == 0x42 ||
                         * p.getOpcode() == 0x62){t=p.getNumParams();}// it's a
                         * ptgfunkvar!
                         * if (p.getOpcode() == 0x21 || p.getOpcode() == 0x41 ||
                         * p.getOpcode() == 0x61){t=p.getNumParams();}// guess that
                         * ptgfunc is not only one..
                         * nparams+=t-1;
                         * if (nparams==0)
                         * break;
                         * } else {
                         * nparams--;
                         * if (nparams==0)
                         * break;
                         * }
                         * }
                         * indirectStack.addAll(stack.subList(z, stack.size()));
                         * indirectStack.push(ptfv);
                         * rec.getWorkBook().addIndirectFormulaStack(indirectStack);
                         * // must save and calculate indirect reference AFTER all
                         * formulas/cells have been added ...
                         * // original is below
                         * /
                         **/
                                /**/
                                if (rec.opcode == XLSConstants.FORMULA) {
                                    (rec as Formula).setContainsIndirectFunction(true)
                                } else if (rec.opcode == XLSConstants.SHRFMLA) {
                                    (rec as Shrfmla).setContainsIndirectFunction(true)
                                }
                                /**/
                            }
                            stack.push(ptfv)
                        }

                        ptgAdd -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgAdd Located")
                            val pad = PtgAdd()
                            ptgLen = pad.length
                            b = ByteArray(ptgLen)
                            // if((ptgLen+i) <= function.length)
                            System.arraycopy(function, i, b, 0, ptgLen)
                            pad.init(b)
                            pad.parentRec = rec
                            stack.push(pad)
                        }

                        ptgMissArg -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgMissArg Located")
                            val pmar = PtgMissArg()
                            ptgLen = pmar.length
                            b = ByteArray(ptgLen)
                            // if((ptgLen+i) <= function.length)
                            System.arraycopy(function, i, b, 0, ptgLen)
                            pmar.init(b)
                            pmar.parentRec = rec
                            stack.push(pmar)
                        }

                        ptgSub -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgSub Located")
                            val psb = PtgSub()
                            ptgLen = psb.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            psb.init(b)
                            psb.parentRec = rec
                            stack.push(psb)
                        }

                        ptgMlt -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgMlt Located")
                            val pml = PtgMlt()
                            ptgLen = pml.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pml.init(b)
                            pml.parentRec = rec
                            stack.push(pml)
                        }

                        ptgDiv -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgDiv Located")
                            val pdiv = PtgDiv()
                            ptgLen = pdiv.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pdiv.init(b)
                            pdiv.parentRec = rec
                            stack.push(pdiv)
                        }

                        ptgUPlus -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgUPlus Located")
                            val puplus = PtgUPlus()
                            ptgLen = puplus.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            puplus.init(b)
                            puplus.parentRec = rec
                            stack.push(puplus)
                        }

                        ptgUMinus -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgUminus Located")
                            val puminus = PtgUMinus()
                            ptgLen = puminus.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            puminus.init(b)
                            puminus.parentRec = rec
                            stack.push(puminus)
                        }

                        ptgPercent -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgPercent Located")
                            val pperc = PtgPercent()
                            ptgLen = pperc.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pperc.init(b)
                            pperc.parentRec = rec
                            stack.push(pperc)
                        }

                        ptgPower -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgPower Located")
                            val pow = PtgPower()
                            ptgLen = pow.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pow.init(b)
                            pow.parentRec = rec
                            stack.push(pow)
                        }

                        ptgConcat -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgConcat Located")
                            val pcon = PtgConcat()
                            ptgLen = pcon.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pcon.init(b)
                            pcon.parentRec = rec
                            stack.push(pcon)
                        }

                        ptgLT -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgLT Located")
                            val plt = PtgLT()
                            ptgLen = plt.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            plt.init(b)
                            plt.parentRec = rec
                            stack.push(plt)
                        }

                        ptgLE -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgLE Located")
                            val ple = PtgLE()
                            ptgLen = ple.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            ple.init(b)
                            ple.parentRec = rec
                            stack.push(ple)
                        }

                        ptgEQ -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgEQ Located")
                            val peq = PtgEQ()
                            ptgLen = peq.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            peq.init(b)
                            peq.parentRec = rec
                            stack.push(peq)
                        }

                        ptgGE -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgGE Located")
                            val pge = PtgGE()
                            ptgLen = pge.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pge.init(b)
                            pge.parentRec = rec
                            stack.push(pge)
                        }

                        ptgGT -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgGT Located")
                            val pgt = PtgGT()
                            ptgLen = pgt.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pgt.init(b)
                            pgt.parentRec = rec
                            stack.push(pgt)
                        }

                        ptgNE -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgNE Located")
                            val pne = PtgNE()
                            ptgLen = pne.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)

                            pne.init(b)
                            pne.parentRec = rec
                            stack.push(pne)
                        }

                        ptgIsect -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgIsect Located")
                            val pist = PtgIsect()
                            ptgLen = pist.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)

                            pist.init(b)
                            pist.parentRec = rec
                            stack.push(pist)
                        }

                        ptgUnion -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgUnion Located")
                            val pun = PtgUnion()
                            ptgLen = pun.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pun.init(b)
                            pun.parentRec = rec
                            stack.push(pun)
                        }

                        ptgRange -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("ptgRange Located")
                            val pran = PtgRange()
                            ptgLen = pran.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)
                            pran.init(b)
                            pran.parentRec = rec
                            stack.push(pran)
                        }

                        ptgParen -> {
                            if (DEBUGLEVEL > 5)
                                Logger.logInfo("PtgParens Located")
                            val pp = PtgParen()
                            ptgLen = pp.length
                            b = ByteArray(ptgLen)
                            if (ptgLen + i <= function.size)
                                System.arraycopy(function, i, b, 0, ptgLen)

                            pp.init(b)
                            pp.parentRec = rec
                            stack.push(pp)
                        }

                        else -> {
                            val pmy = PtgMystery()
                            ptgLen = function.size - i
                            b = ByteArray(ptgLen)
                            if (DEBUGLEVEL > 0)
                                Logger.logWarn("Unsupported Formula Function: 0x"
                                        + Integer.toHexString(ptg.toInt()) + " length: " + ptgLen)
                            System.arraycopy(function, i, b, 0, ptgLen)
                            pmy.init(b)
                            pmy.parentRec = rec
                            stack.push(pmy)
                        }
                    }// hasPtgExtraMem= true; // has a PtgExtraMem structure
                    // after end of parsed expression: The PtgExtraMem structure
                    // specifies a range that corresponds to a PtgMemArea as
                    // specified in RgbExtra.
                    // pma= pmem; // save for later
                    i += ptgLen
                }
            }
            if (hasArrays && rec is Formula) { // Array Recs handle extra
                // data differently
                // array data is appended to end of expression
                // for each array in the function list,
                // get saved ptgArray var (stored in stack var),
                // grab data and parse array components
                var startPos = expressionLen
                for (i in arrayLocs.indices) {
                    val ingr = arrayLocs.elementAt(i) as Int
                    val parr = stack.elementAt(ingr) as PtgArray

                    // have to assume that remaining data all goes for this
                    // ptgarray
                    // since length is variable and can only be ascertained by
                    // parsing
                    // if multiple arrays are present, actual array length will
                    // be returned via setArrVals
                    val b = ByteArray(function.size - startPos) // get "extra"
                    // array
                    // data
                    System.arraycopy(function, startPos, b, 0, b.size)
                    try {
                        parr.parentRec = rec
                        startPos += parr.setArrVals(b)
                    } catch (e: Exception) {// TODO: this needs to be caught due to
                        // "name" records being parsed
                        // incorrectly. The problem has to do
                        // with the lenght of the name record
                        // not including the extra 7 bytes of
                        // space. Temporary fix for infoteria
                        if (DEBUGLEVEL > 0)
                            Logger.logInfo("ExpressionParser.parseExpression: Array: $e")
                    }

                }
            } /*
         * no need to keep PtgExtraMem as can regenerate easily else
         * if (hasPtgExtraMem && rec instanceof Formula) {
         * //The PtgExtraMem structure specifies a range that
         * corresponds to a PtgMemArea as specified in RgbExtra.)
         * // count (2 bytes): An unsigned integer that specifies
         * the areas within the range.
         * // array (variable): An array of Ref8U that specifies the
         * range. The number of elements MUST be equal to count.
         * pma.setPostExpression(function, expressionLen);
         * }
         */
            if (DEBUGLEVEL > 5)
                Logger.logInfo("finished formula")
            return stack

        }

        /*
     * Returns the ptg that matches the string location sent to
     * it.
     * rec can either be in the format "C5" or a range, such as
     * "C4:D9"
     *
     */
        @Throws(FormulaNotFoundException::class)
        fun getPtgsByLocation(loc: String, expression: Stack<*>): List<Ptg> {
            val lv = Vector<Ptg>()
            for (i in expression.indices) {
                val o = expression.elementAt(i) ?: throw FormulaNotFoundException(
                        "Couldn't get Ptg at: $loc")
                if (o is Byte) {
                    // do nothing
                } else if (o is Ptg) {
                    var lo: String? = o.location
                    if (lo == null)
                        lo = "none"
                    var comp = loc
                    if (loc.indexOf("!") > -1) { // the sheet is referenced
                        if (lo.indexOf("!") == -1) { // and the ptg does not have
                            // sheet referenced
                            comp = loc.substring(loc.indexOf("!") + 1)
                        }
                    }

                    if (comp.equals(lo, ignoreCase = true)) {
                        lv.add(o)
                    } else {
                        // try fq location
                        lo = o.toString()
                        if (loc.equals(lo, ignoreCase = true)) {
                            lv.add(o)

                        } else if (o is PtgRef3d) {// gotta look into the
                            // first & last
                            // already checked
                        } else if (o is PtgArea) {// gotta look into the
                            // first & last
                            val first = o.firstPtg
                            val last = o.lastPtg
                            if (first!!.location.equals(loc, ignoreCase = true))
                                lv.add(first)
                            if (last!!.location.equals(loc, ignoreCase = true))
                                lv.add(last)
                        }
                    }
                }
            }
            return lv
        }

        /**
         * returns the position in the expression stack for the ptg associated with this location
         *
         * @param loc        String
         * @param expression
         * @return
         * @throws FormulaNotFoundException
         */
        @Throws(FormulaNotFoundException::class)
        fun getExpressionLocByLocation(loc: String, expression: Stack<*>): Int {

            for (i in expression.indices) {
                val o = expression.elementAt(i) ?: throw FormulaNotFoundException(
                        "Couldn't get Ptg at: $loc")
                if (o is Byte) {
                    // do nothing
                } else if (o is Ptg) {
                    var lo = o.location
                    if (loc.equals(lo, ignoreCase = true)) {
                        return i
                    }
                    // try full location
                    lo = o.toString()
                    if (loc.equals(lo, ignoreCase = true)) {
                        return i
                    }
                    if (o is PtgArea) {// gotta look into the first & last
                        val first = o.firstPtg
                        val last = o.lastPtg
                        if (first!!.location.equals(loc, ignoreCase = true))
                            return i
                        if (last!!.location.equals(loc, ignoreCase = true))
                            return i
                    }
                }
            }
            return -1
        }

        /**
         * returns the position in the expression stack for the desired ptg
         *
         * @param ptg        Ptg to lookk up
         * @param expression
         * @return
         * @throws FormulaNotFoundException
         */
        @Throws(FormulaNotFoundException::class)
        fun getExpressionLocByPtg(ptg: Ptg, expression: Stack<*>): Int {

            for (i in expression.indices) {
                val o = expression.elementAt(i) ?: throw FormulaNotFoundException(
                        "Couldn't get Ptg at: $ptg")
                if (o is Byte) {
                    // do nothing
                } else if (o is Ptg) {
                    if (o == ptg)
                        return i
                }
            }
            return -1
        }

        /**
         * getCellRangePtgs handles locating which cells are refereced in an expression stack.
         *
         *
         * Essentially the use is we can check a formula if it refereces a cell that is moving, then we have
         * the ability to manipulate these ranges in whatever way makes sense.
         *
         * @return an array of ptgs that are location based (ptgRef, PtgArea)
         * @expression = a Stack of ptgs that represent an excel calculation.
         */
        @Throws(FormulaNotFoundException::class)
        fun getCellRangePtgs(expression: Stack<*>): Array<Ptg> {
            val ret = Vector()
            for (i in expression.indices) {
                val o = expression.elementAt(i) ?: throw FormulaNotFoundException("Couldn't get Ptg at: $i")
                if (o is Byte) {
                    // do nothing
                } else if (o is Ptg) {
// handle shared formula range
                    if (o is PtgExp) {
                        val lox = o.location
                        val ref = PtgRef()
                        ref.parentRec = o.parentRec // must be done
                        // before
                        // setLocation
                        ref.location = lox
                        ret.add(ref)
                    } else if (o is PtgRefErr || o is PtgAreaErr3d) {
                        ret.add("#REF!")
                    } else if (o is PtgMemFunc) {
                        // Ptg[] p=
                        // getCellRangePtgs(((PtgMemFunc)part).getSubExpression());
                        val p = o.components
                        for (z in p.indices)
                            ret.add(p[z])
                    } else {
                        val lox = o.location
                        if (lox != null)
                            ret.add(o)
                    }
                }
            }
            val retp = arrayOfNulls<Ptg>(ret.size)
            return ret.toTypedArray() as Array<Ptg>
        }
    }

}