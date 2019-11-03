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

import io.starter.formats.XLS.ExpressionParser
import io.starter.formats.XLS.FunctionNotSupportedException
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList
import java.util.Stack

/**
 * PtgMemArea is an optimization of referenced areas.  Sweet!
 * *******************************************************************************
 * NOTE:  Below is from documentation but DOES NOT APPEAR to be what happens in actuality;
 * PtgMemArea token is followed by several ptg reference-types plus ptgunion(s), ends with a PtgParen.
 * The cce field is the length of all of these following tokens.
 * These following Ptgs are set and parsed in .setPostRecord
 *
 *
 *
 *
 *
 *
 * Like most optimizations it really sucks.  It is also one of the few Ptg's that
 * has a variable length.
 *
 *
 * Format of length section
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           (reserved)     4       Whatever it may be
 * 2           cce			   2	   length of the reference subexpression
 *
 * Format of reference Subexpression
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0			cref		2			The number of rectangles to follow
 * 2			rgref		var			An Array of rectangles
 *
 * Format of Rectangles
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rwFirst     2       The First row of the reference
 * 2           rwLast     2       The Last row of the reference
 * 4           ColFirst    1       (see following table)
 * 6           ColLast    1       (see following table)
</pre> *
 *
 * @see Ptg
 *
 * @see Formula
 */
open class PtgMemArea : GenericPtg() {
    internal var cce = 0
    internal var subexpression: Stack<*>? = null
    //PtgRef[] comps = null; not used anymore; see note below

    /**
     * Returns all of the cells of this range as PtgRef's.
     * This includes empty cells, values, formulas, etc.
     * Note the setting of parent-rec requires finding the cell
     * the PtgRef refer's to.  If that is null then the PtgRef
     * will exist, just with a null value.  This could cause issues when
     * programatically populating cells.
     *
     *
     * NOTE: now obtaining component ptgs is done in populateValues as it is
     * a more complex operation than simply gathering all referenced ptgs
     */
    override/*if(comps!=null) // cache
        		return comps;

        ArrayList v = new ArrayList();
        try {
    	        for (int i= 0; i < ptgs.length; i++) {
    	        	if (ptgs[i] instanceof PtgArea) {
    	        		Ptg[] ps= ((PtgRef) ptgs[i]).getComponents();
    	        		for (int j= 0; j< ps.length; j++)
    	        			v.add(ps[j]);
    	        	} else if (ptgs[i] instanceof PtgRef) {
    	        		v.add((PtgRef) ptgs[i]);
    	        	} else { // it's a PtgName
    	        		Ptg[] pcomps= ((PtgName) ptgs[i]).getComponents();
    	        		for (int j= 0; j<pcomps.length; j++)
    	        			v.add(pcomps[j]);
    	        	}
    		    }
        }catch (Exception e){Logger.logInfo("calculating formula range value in PtgArea failed: " + e);}
        comps = new PtgRef[v.size()];
        v.toArray(comps);z
        */ var components: Array<Ptg>? = null
        internal set

    override val isOperand: Boolean
        get() = true

    // really no need to keep postexpression as can regenerate easily ...
    internal var postExp: ByteArray? = null

    /**
     * retrieves the PtgExtraMem structure that is located at the end of the function record
     *
     * @return
     */
    // first count # refs (excluding ptgUnion, range, etc )
    // input cce + cce*Ref8U (8 bytes) describing reference
    // rw first
    // col first
    // a single ref; repeat
    // rw last
    // col last
    // a range
    // rw last
    // col last
    /* KSC: TESTING
		if (!Arrays.equals(recbytes, postExp))
			io.starter.toolkit.Logger.log("ISSUE!!!");*/ val postRecord: ByteArray
        get() {
            var cce: Short = 0
            for (i in subexpression!!.indices) {
                val p = subexpression!![i] as Ptg
                if (p is PtgRef) {
                    cce++
                }
            }
            val b = ByteTools.shortToLEBytes(cce)
            val recbytes = ByteArray(cce * 8 + 2)
            recbytes[0] = b[0]
            recbytes[1] = b[1]
            var pos = 2
            var i = 0
            while (i < subexpression!!.size && pos + 7 < recbytes.size) {
                val p = subexpression!![i] as Ptg
                if (p is PtgRef) {
                    val rc = p.rowCol
                    System.arraycopy(ByteTools.shortToLEBytes(rc[0].toShort()), 0, recbytes, pos, 2)
                    System.arraycopy(ByteTools.shortToLEBytes(rc[1].toShort()), 0, recbytes, pos + 4, 2)
                    if (rc.size == 2) {
                        System.arraycopy(ByteTools.shortToLEBytes(rc[0].toShort()), 0, recbytes, pos + 2, 2)
                        System.arraycopy(ByteTools.shortToLEBytes(rc[1].toShort()), 0, recbytes, pos + 6, 2)
                    } else {
                        System.arraycopy(ByteTools.shortToLEBytes(rc[2].toShort()), 0, recbytes, pos + 2, 2)
                        System.arraycopy(ByteTools.shortToLEBytes(rc[3].toShort()), 0, recbytes, pos + 6, 2)
                    }
                    pos += 8
                }
                i++
            }
            return recbytes
        }

    internal var refsheets = ArrayList()

    /**
     * generate the bytes necessary to describe this PtgMemArea;
     * extra data described by getPostRecord is necessary for completion
     *
     * @see getPostRecord
     */
    override// bytes 1-4 are unused
    val record: ByteArray
        get() {
            var len = 0
            for (i in subexpression!!.indices) {
                val p = subexpression!![i] as Ptg
                len += p.record.size
            }
            cce = len
            val rec = ByteArray(len + 7)
            val b = ByteTools.shortToLEBytes(cce.toShort())
            rec[0] = 0x26
            rec[5] = b[0]
            rec[6] = b[1]
            var offset = 7
            for (i in subexpression!!.indices) {
                val p = subexpression!![i] as Ptg
                System.arraycopy(p.record, 0, rec, offset, p.record.size)
                offset += p.record.size
            }
            record = rec
            return record
        }

    override val length: Int
        get() = cce + 7

    override val value: Any?
        get() {
            try {
                val dub: DoubleArray?
                try {
                    dub = PtgCalculator.getDoubleValueArray(components!!)
                } catch (e: CalculationException) {
                    return null
                }

                var result = 0.0
                for (i in dub!!.indices) {
                    result += dub[i]
                }
                return result
            } catch (e: FunctionNotSupportedException) {
                Logger.logWarn("Function Unsupported error in PtgMemFunction: $e")
                return null
            }

        }


    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        this.populateVals()
    }

    fun getnTokens(): Int {
        return cce
    } // KSC:

    /**
     * sets the bytes to describe the subexpression (set of ptgs)
     * and parses the subexpression
     *
     * @param b
     */
    fun setSubExpression(b: ByteArray) {
        var retbytes = ByteArray(7)
        System.arraycopy(record, 0, retbytes, 0, record.size)
        retbytes = ByteTools.append(b, retbytes)
        record = retbytes
        this.populateVals()
        // TODO:
    }

    /**
     * sets the PtgExtraMem structure which is appended to the end of the function array
     *
     * @param b
     */
    fun setPostExpression(b: ByteArray, expressionLen: Int) {
        val len = b.size - expressionLen
        postExp = ByteArray(len)
        System.arraycopy(b, expressionLen, postExp!!, 0, len)
        /*
 * 		parsing PtgExtraMem - not really necessary
 *
        if (b.length > expressionLen+2) {
	 		int z= expressionLen;
			int count= ByteTools.readShort(b[z++], b[z++]);
			for (int i= 0; i < count && z+7 < b.length; i++) {
				int[] rc= new int[4];
				rc[0]= ByteTools.readShort(b[z++], b[z++]);	// rw first
				rc[2]= ByteTools.readShort(b[z++], b[z++]);	// rw last
				rc[1]= ByteTools.readShort(b[z++], b[z++]);// col first
				rc[3]= ByteTools.readShort(b[z++], b[z++]); // col last
				ExcelTools.formatRangeRowCol(rc);
			}
		}*/
    }

    internal open fun populateVals() {
        // 1st byte = ID, next 4 are ignored
        // cce= size of following sub-expressions
        cce = ByteTools.readShort(record[5].toInt(), record[6].toInt()).toInt()
        // this is not really correct!
        if (record.size > 7) {
            val subexp: ByteArray
            subexp = ByteArray(cce)
            System.arraycopy(record, 7, subexp, 0, cce)
            subexpression = ExpressionParser.parseExpression(subexp, this.parentRec)
            // subexpression stack in form of:  REFERENCE, REFERENCE, OP [,REFERENCE, OP] ...
            // op can be one of:  PtgUnion [,] PtgIsect [ ] or PtgRange [:]
            // calculate subexpression to obtain ptgs
            try {
                val o = FormulaCalculator.calculateFormula(this.subexpression!!)
                val components = ArrayList()
                if (o != null && o is Array<Ptg>) {
                    this.components = o
                    for (i in this.components!!.indices) {
                        if (!refsheets.contains((this.components!![i] as PtgRef).sheetName))
                            refsheets.add((this.components!![i] as PtgRef).sheetName)
                        if (this.components!![i] is PtgArea) {
                            val p = this.components!![i].components
                            for (j in p.indices)
                                components.add(p[j])
                        } else
                            components.add(this.components!![i])
                    }
                } else {    // often a single reference surrounded by parens
                    for (i in subexpression!!.indices) {
                        try {
                            val pr = subexpression!![i] as PtgRef
                            if (!refsheets.contains(pr.sheetName))
                                refsheets.add(pr.sheetName)
                            if (pr is PtgArea) {
                                val pa = pr.components
                                for (j in pa!!.indices)
                                    components.add(pa[j])
                            } else
                                components.add(pr)
                        } catch (e: Exception) {
                        }

                    }
                }
                this.components = arrayOfNulls(components.size)
                components.toTypedArray()
            } catch (e: Exception) {
                Logger.logErr("PtgMemArea init: $e")
            }

            //int z= subexpression.size();
            // to get # of references (PtgRefs) = stack size/2 + 1
        }
    }

    /**
     * traverse through expression to retrieve set of ranges
     * either discontiguous union (,), intersected ( ) or regular range (:)
     */
    override fun toString(): String {
        return FormulaParser.getExpressionString(subexpression!!).substring(1)    // avoid "="
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -6869393084367355874L
    }


}
