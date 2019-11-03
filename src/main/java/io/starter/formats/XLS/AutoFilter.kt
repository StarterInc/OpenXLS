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
package io.starter.formats.XLS


import io.starter.OpenXLS.ExcelTools
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.ArrayList


/**
 * Auto filter controls the auto-filter capabilities of Excel.  It has numerous sub-records and references
 *
 *
 * p><pre>
 * offset  name        size    contents
 * ---
 * 4       iEntry          2       Index of the active autofilter
 * 6       grbit           2       Option flags
 * 8       doper1          10      DOPER struct for the first filter condition
 * 18      doper2          10      DOPER struct for the second filter condition
 * 28      rgch            var     String storage for vtString DOPER
 *
</pre> *
 */
/* TODO: ************************************************************************************************************
 * Creation of new AutoFilters, removal of existing (see Boundsheet)
 *
 * fix:  iEntry is not always the index of the column
 * APPARENTLY if there are more than two autofilters, iEntry is the index of the column
 * if there is only 1 autofilter, iEntry is 0 --- dependent upon Obj record????
 *
 * Finish:  setTop10, setVal, setVal2: verify all is correctly done ....
 * *******************************************************************************************************************
 */
class AutoFilter : io.starter.formats.XLS.XLSRecord() {

    internal var iEntry: Short = 0    // *** this is the column number ***  oops!  not always!!!
    internal var doper1: Doper? = null
    internal var doper2: Doper? = null

    // booleans used here for memory space/grbit fields, these are really 1/0 values.  whas the diff?
    internal var wJoin: Boolean = false// true if custom filter conditions are ORed
    internal var fSimple1: Boolean = false// true if the first condition is a simple equality;,
    internal var fSimple2: Boolean = false // trueif the second condition is a simple equality;
    internal var fTop10: Boolean = false // true if the condition is a Top10 autofilter
    internal var fTop: Boolean = false // true if the top 10 AutoFilter shows the top itemsl 0 if it shows the bottom items
    internal var fPercent: Boolean = false // true if the Top 10 AutoFilter shows percentage, 0 if it shows items
    internal var wTop10: Short = 0 //The number of items to show (from 1-500)

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, /* iEntry */
            0, 0, /* grbit */
            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, /* unused Doper1 */
            0, 0, 0, 0, 0, 0, 0, 0, 0, 0)/* unused Doper2 */

    /**
     * return the column number (0-based) that this AutoFilter references
     *
     * @return int column number
     */
    /**
     * set the column number (0-based) that this AutoFilter references
     *
     * @param int col - 0-based column number
     */
    var col: Int
        get() = iEntry.toInt()
        set(col) {
            iEntry = col.toShort()
            update()
        }

    /**
     * return the value of the first comparison of this AutoFilter, if any
     *
     * @return Object value
     * @see getVal2
     */
    override val `val`: Any?
        get() = doper1!!.toString()

    /**
     * returns the value of the second comparison of this AutoFilter, if any
     *
     * @return Object value
     * @see getVal
     */
    val val2: Any?
        get() = doper2!!.toString()

    /**
     * get the operator associated with this AutoFilter
     * <br></br>NOTE: this will return the operator in the first condition if this AutoFilter contains two conditions
     * <br></br>Use getOp2 to retrieve the second condition operator
     *
     * @return String operator: one of "=", ">", ">=", "<>", "<", "<="
     * @see getOp2
     */
    val op: String
        get() = doper1!!.comparisonOperator

    /**
     * get the operator associated with the second condtion of this AutoFilter, if any
     * <br></br>NOTE: this will return the operator in the second condition if this AutoFilter contains two conditions
     *
     * @return String operator: one of "=", ">", ">=", "<>", "<", "<=", or null, if no second condition
     * @see getOp
     */
    val op2: String?
        get() = if (doper2 !is UnusedDoper) doper2!!.comparisonOperator else null

    /**
     * returns true if this AutoFitler is set to filter blank rows
     *
     * @return boolean
     */
    val isFilterBlanks: Boolean
        get() = fSimple1 && !fSimple2 && !fTop && !fTop10 && doper1 is NoBlanksDoper

    /**
     * returns true if this AutoFitler is set to filter all non-blank rows
     *
     * @return boolean
     */
    val isFilterNonBlanks: Boolean
        get() = fSimple1 && !fSimple2 && !fTop && !fTop10 && doper1 is AllBlanksDoper

    /**
     * returns true if this is a top-10-type AutoFilter
     *
     * @return true if this is a top-10-type AutoFilter, false otherwise
     */
    val isTop10: Boolean
        get() = fTop10 && wTop10 > 0

    /**
     * initialize the AutoFilter record
     */
    override fun init() {
        super.init()
        iEntry = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        //parse out grbit flags
        val grbit = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        this.decodeGrbit(grbit)
        //parse both DOPERs
        doper1 = parseDoper(this.getBytesAt(4, 10)!!)
        doper2 = parseDoper(this.getBytesAt(14, 10)!!)
        // parse string data, if any, appended to end of data record
        if (this.getData()!!.size > 25) {
            val rgch = this.getBytesAt(25, this.getData()!!.size - 25)
            var pos = 0
            if (doper1 is StringDoper) {
                (doper1 as StringDoper).setString(rgch, pos)
                pos += (doper1 as StringDoper).cch    // set position pointer to next string data, if any
                pos++
            }
            if (doper2 is StringDoper) {
                (doper2 as StringDoper).setString(rgch, pos)
            }
        }
    }

    /**
     * Parse out the grbit
     *
     * @param grbt
     */
    private fun decodeGrbit(grbit: Short) {    // top 500: grbit= -1488
        wJoin = grbit and 0x3 == 0x3    // 0= AND, 3= OR
        fSimple1 = grbit and 0x4 == 0x4
        fSimple2 = grbit and 0x8 == 0x8
        fTop10 = grbit and 0x10 == 0x10
        fTop = grbit and 0x20 == 0x20
        fPercent = grbit and 0x40 == 0x40
        wTop10 = (grbit and 0xFF80 shr 7).toShort()
    }

    /**
     * encode the grbit from the source flags
     *
     * @return short encoded grbit
     */
    private fun encodeGrbit(): Short {
        var grbit: Short = 0
        if (wJoin) grbit = 3    // Or	0= AND
        grbit = ByteTools.updateGrBit(grbit, fSimple1, 2)
        grbit = ByteTools.updateGrBit(grbit, fSimple2, 3)
        grbit = ByteTools.updateGrBit(grbit, fTop10, 4)
        grbit = ByteTools.updateGrBit(grbit, fTop, 5)
        grbit = ByteTools.updateGrBit(grbit, fPercent, 6)
        grbit = (wTop10 shl 7 or grbit).toShort()
        return grbit
    }

    /**
     * Take an array of 10 bytes and parse out the doper
     *
     * @param doperBytes
     * @return
     */
    private fun parseDoper(doperBytes: ByteArray): Doper? {
        var retDoper: Doper? = null
        when (doperBytes[0]) {
            0x0 -> retDoper = UnusedDoper(doperBytes)
            0x2 -> retDoper = RKDoper(doperBytes)
            0x4 -> retDoper = IEEEDoper(doperBytes)
            0x6 -> retDoper = StringDoper(doperBytes)
            0x8 -> retDoper = ErrorDoper(doperBytes)
            0xC -> retDoper = AllBlanksDoper(doperBytes)
            0xE -> retDoper = NoBlanksDoper(doperBytes)
        }
        return retDoper
    }

    /**
     * commit all the flags, dopers, etc to the byte rec array
     */
    fun update() {
        var data = ByteArray(25)
        var b = ByteTools.shortToLEBytes(iEntry)
        System.arraycopy(b, 0, data, 0, 2)
        b = ByteTools.shortToLEBytes(encodeGrbit())
        System.arraycopy(b, 0, data, 2, 2)
        System.arraycopy(doper1!!.record, 0, data, 4, 10)
        System.arraycopy(doper2!!.record, 0, data, 14, 10)
        if (doper1 is StringDoper)
        // append rgch bytes
            data = ByteTools.append((doper1 as StringDoper).strBytes, data)
        if (doper2 is StringDoper)
        // append rgch bytes
            data = ByteTools.append((doper2 as StringDoper).strBytes, data)
        setData(data)
    }

    /**
     * Evaluates this AutoFilter's condition for each cell in the indicated column
     * over all rows in the sheet; if the condition is not met, the row is set to hidden
     * <br></br>NOTE: since there may be other conditions (other AutoFilters, for instance)
     * setting the row to it's hidden state, this method **will not**
     * set the row to unhidden if the condition passes.
     */
    fun evaluate() {
        val val1: Any?
        var val2: Any? = null
        val1 = getVal(doper1)    // get the doper value from the 1st doper/comparison, if any
        val hasDoper2 = doper2 !is UnusedDoper
        if (hasDoper2)
            val2 = getVal(doper2)

        var op1 = "="
        var op2 = ""
        var passes = true
        if (!fSimple1) {
            op1 = doper1!!.comparisonOperator
        }
        if (!fSimple2 && hasDoper2) {
            op2 = doper2!!.comparisonOperator
        }

        // TODO:
        // above/below average?  ooxml only?
        // date/time comparisons????
        // begins with, ends with ...
        if (fTop10) {
            if (fTop) { // ascending
                evaluateTopN()
            } else
            // descending
                evaluateBottomN()
        } else if (doper1 is AllBlanksDoper || doper1 is NoBlanksDoper) {
            val filterBlanks = doper1 is NoBlanksDoper
            val n = this.sheet!!.numRows
            for (i in 0 until n) {
                var r = this.sheet!!.getRowByNumber(i)
                if (r == null) {// it's blank
                    // create a blank cell and then set to hidden?
                    if (filterBlanks) {
                        this.sheet!!.addValue("", ExcelTools.formatLocation(intArrayOf(i, iEntry.toInt())))
                        r = this.sheet!!.getRowByNumber(i)
                        r!!.isHidden = true
                    }
                } else {    //row is not blank, check to see if cell is blank
                    try {
                        val c = r.getCell(iEntry)
                        if (c is Blank && filterBlanks)
                            r.isHidden = true
                        else if (!filterBlanks)
                            r.isHidden = true
                    } catch (e: NullPointerException) {
                        // NPE= blank
                        if (filterBlanks)
                            r.isHidden = true
                    } catch (e: CellNotFoundException) {

                    }

                }
            }
            val rows = this.sheet!!.rows
            if (!filterBlanks) { // easy; everything that is NOT BLANK is hidden
                for (i in 1 until rows.size)
                    rows[i].isHidden = true
            } else { // everything that is blank is hidden

            }
        } else { // all other criteria are based upon operator comparisons ...
            val rows = this.sheet!!.rows
            for (i in 1 until rows.size) {
                try {
                    val c = rows[i].getCell(iEntry)
                    passes = io.starter.formats.XLS.formulas.Calculator.compareCellValue(c, val1, op1)
                    if (hasDoper2 && (wJoin || !wJoin && passes))
                        passes = io.starter.formats.XLS.formulas.Calculator.compareCellValue(c, val2, op2)
                    if (!passes)
                        rows[i].isHidden = true
                } catch (e: Exception) {
                }
                // just keep evaluation
            }
        }
    }


    /**
     * gets the Object Value of the Doper
     * either a Double, String or Boolean type
     *
     * @param d - Doper
     * @return Object value of the Doper
     */
    private fun getVal(d: Doper?): Any? {
        var `val`: Any? = null
        if (d is ErrorDoper) {
            if (d.isBooleanVal)
                `val` = java.lang.Boolean.valueOf(d.booleanVal)
            else
                `val` = d.getErrVal()    // error doper
        } else if (d is StringDoper)
        // string comparison
            `val` = d.toString()
        else if (d is RKDoper)
            `val` = d.getVal()
        else if (d is IEEEDoper)
            `val` = d.getVal()
        return `val`
    }

    /**
     * evaluates a Top-N condition, hiding rows that are not in the top N values of the column
     *
     * @see evaluateBottomN    for descending or Bottom-N evaluation
     */
    private fun evaluateTopN() {
        // must go thru 1+ times as must gather up values then go back and set hidden ...
        // identifies top n values then displays ALL rows that contain those values
        val top10 = ArrayList()
        val n = if (!fPercent) wTop10 else this.sheet!!.numRows / wTop10
        val maxVals = DoubleArray(n)
        for (i in 0 until n) maxVals[i] = java.lang.Double.NEGATIVE_INFINITY
        var curmin = java.lang.Double.NEGATIVE_INFINITY
        val rows = this.sheet!!.rows
        for (i in 1 until rows.size) {
            try {
                val c = rows[i].getCell(iEntry)
                val `val` = c.dblVal
                var insertionpoint = -1
                if (`val` >= curmin) {
                    for (j in 0 until n) {    // see where new value falls
                        if (`val` == maxVals[j]) { // then no need to move values around; just add index to set of indexes
                            var idxs: String? = top10.get(j)
                            if (idxs == null)
                                idxs = i + ", "
                            else
                                idxs += i + ", "
                            top10.set(j, idxs)
                            insertionpoint = -1    // so don't add below
                            break
                        } else if (`val` > maxVals[j]) {
                            if (insertionpoint == -1 || maxVals[j] < maxVals[insertionpoint]) {
                                insertionpoint = j    // overwrite point
                            }
                        }
                    }
                    if (insertionpoint >= 0) {
                        if (top10.size > insertionpoint)
                            top10.removeAt(insertionpoint)    // replace below
                        val idxs = i + ", "
                        top10.add(insertionpoint, idxs)
                        maxVals[insertionpoint] = `val`
                        // reestablish curmin
                        curmin = java.lang.Double.MAX_VALUE
                        for (j in 0 until n)
                            curmin = Math.min(maxVals[j], curmin)
                    }
                } // othwerwise doesn't meet criteia

            } catch (e: Exception) {
            }
            // just keep evaluation
        }
        // get master list of non-hidden rows
        var nonhiddenIdxs = ""
        for (i in top10.indices) {
            nonhiddenIdxs += top10.get(i) as String
        }
        // now that have list of rows which SHOULDN'T be hidden, set all else to hidden
        for (j in 1 until rows.size) {
            val idx = j + ", "
            if (nonhiddenIdxs.indexOf(idx) == -1)
            // not found!
                rows[j].isHidden = true
        }
    }

    /**
     * evaluates a Top-N condition in descending order, in other words, a Bottom-N evaluation,
     * hiding rows that are not in the bottom N values of the column
     *
     * @see evaluateTopN    for ascending or Top-N evaluation
     */
    private fun evaluateBottomN() {
        // must go thru 1+ times as must gather up values then go back and set hidden ...
        // identifies bottom n values then displays ALL rows that contain those values
        val bottomN = ArrayList()
        val n = if (!fPercent) wTop10 else this.sheet!!.numRows / wTop10
        val minVals = DoubleArray(n)
        for (i in 0 until n) minVals[i] = java.lang.Double.POSITIVE_INFINITY
        var curmax = java.lang.Double.POSITIVE_INFINITY
        val rows = this.sheet!!.rows
        for (i in 1 until rows.size) {
            try {
                val c = rows[i].getCell(iEntry)
                val `val` = c.dblVal
                var insertionpoint = -1
                if (`val` <= curmax) {
                    for (j in 0 until n) {    // see where new value falls
                        if (`val` == minVals[j]) { // then no need to move values around; just add index to set of indexes
                            var idxs: String? = bottomN.get(j)
                            if (idxs == null)
                                idxs = i + ", "
                            else
                                idxs += i + ", "
                            bottomN.set(j, idxs)
                            insertionpoint = -1    // so don't add below
                            break
                        } else if (`val` < minVals[j]) {
                            if (insertionpoint == -1 || minVals[j] > minVals[insertionpoint]) {
                                insertionpoint = j    // overwrite point
                            }
                        }
                    }
                    if (insertionpoint >= 0) {
                        if (bottomN.size > insertionpoint)
                            bottomN.removeAt(insertionpoint)    // replace below
                        val idxs = i + ", "
                        bottomN.add(insertionpoint, idxs)
                        minVals[insertionpoint] = `val`
                        // reestablish curmax
                        curmax = java.lang.Double.NEGATIVE_INFINITY
                        for (j in 0 until n)
                            curmax = Math.max(minVals[j], curmax)
                    }
                } // othwerwise doesn't meet criteia

            } catch (e: Exception) {
            }
            // just keep evaluation
        }
        // get master list of non-hidden rows
        var nonhiddenIdxs = ""
        for (i in bottomN.indices) {
            nonhiddenIdxs += bottomN.get(i) as String
        }
        // now that have list of rows which SHOULDN'T be hidden, set all else to hidden
        for (j in 1 until rows.size) {
            val idx = j + ", "
            if (nonhiddenIdxs.indexOf(idx) == -1)
            // not found!
                rows[j].isHidden = true
        }
    }

    /**
     * return a string representation of this autofilter
     */
    override fun toString(): String {
        var op1 = "="
        var op2 = ""
        val hasDoper2 = doper2 != null && doper2 !is UnusedDoper
        if (fTop10) {
            return if (fTop) {
                if (fPercent) "Top $wTop10%" else "Top $wTop10 Items"
            } else {
                if (fPercent) "Bottom $wTop10%" else "Bottom $wTop10 Items"
            }
        } else if (!fSimple1) {
            op1 = doper1!!.comparisonOperator
        } else {
            if (doper1 is AllBlanksDoper)
                return "Non Blanks"
            else if (doper1 is NoBlanksDoper)
                return "Blanks"
        }

        if (!fSimple2 && hasDoper2) {
            op2 = doper2!!.comparisonOperator
        }
        return if (!hasDoper2) op1 + doper1!!.toString()!! else op1 + doper1!!.toString() + (if (wJoin) " OR " else " AND ") + op2 + doper2!!.toString()
    }

    /**
     * Update the record before streaming
     *
     * @see io.starter.formats.XLS.XLSRecord.preStream
     */
    override fun preStream() {
        // no need to update unless things have changed ... this.update();
    }

    /**
     * Sets the custom comparison of this AutoFilter via a String operator and an Object value
     * <br></br>Only those records that meet the equation (column value) <operator> value will be shown
     * <br></br>e.g show all rows where column value >= 2.99
     *
     * Object value can be of type
     *
     * String
     * <br></br>Boolean
     * <br></br>Error
     * <br></br>a Number type object
     *
     * String operator may be one of: "=", ">", ">=", "<>", "<", "<="
     *
     * @param Object val - value to set
     * @param String op - operator
     * @see setVal2
    </operator> */
    fun setVal(`val`: Any, op: String) {
        val doperRec = ByteArray(10)
        if (`val` is String) {// doper1 should be a string
            val s = `val`
            if (!s.startsWith("!")) { // it's not an error val
                doperRec[0] = 0x6
                doper1 = StringDoper(doperRec)
                (doper1 as StringDoper).setString(s)
            } else { // it's an error doper
                doperRec[0] = 0x8
                doperRec[2] = 1    // fError
                doper1 = ErrorDoper(doperRec)
            }
        } else if (`val` is Boolean) {
            doperRec[0] = 0x8
            doperRec[2] = 0    // fError
            doperRec[3] = (if (`val`.booleanValue()) 1 else 0).toByte()
            doper1 = ErrorDoper(doperRec)
        } else {   // assume a Number object
            doperRec[0] = 0x2
            try {
                var d = 0.0
                if (`val` is Double)
                    d = `val`.toDouble()
                else if (`val` is Int)
                    d = `val`.toDouble()
                else
                    throw NumberFormatException("Unable to convert to Numeric Object" + `val`.javaClass)
                doper1 = RKDoper(doperRec)
                (doper1 as RKDoper).setVal(d)
                //doper1= new IEEEDoper(doperRec);
            } catch (e: Exception) {
                Logger.logErr("AutoFilter.setVal: error setting value to $`val`:$e")
            }

        }
        doper1!!.comparisonOperator = op
        fSimple1 = "=" == op
        update()
    }


    /**
     * Sets the custom comparison of the second condition of this AutoFilter via a String operator and an Object value
     * <br></br>This method sets the second condition of a two-condition filter
     *
     * Only those records that meet the equation:
     * <br></br> first condiition AND/OR (column value) <operator> Value will be shown
     * <br></br>e.g show all rows where (column value) <= 1.99 AND (column value) >= 2.99
    </operator> *
     * Object value can be of type
     *
     * String
     * <br></br>Boolean
     * <br></br>Error
     * <br></br>a Number type object
     *
     * String operator may be one of: "=", ">", ">=", "<>", "<", "<="
     *
     * @param Object  val - value to set
     * @param String  op - operator
     * @param boolean AND - true if two conditions should be AND'ed, false if OR'd
     * @see setVal2
     */
    fun setVal2(`val`: Any, op: String, AND: Boolean) {
        val doperRec = ByteArray(10)
        if (`val` is String) {// doper1 should be a string
            val s = `val`
            if (!s.startsWith("!")) { // it's not an error val
                doperRec[0] = 0x6
                doper2 = StringDoper(doperRec)
                (doper1 as StringDoper).setString(s)
            } else { // it's an error doper
                doperRec[0] = 0x8
                doperRec[2] = 1    // fError
                doper2 = ErrorDoper(doperRec)
            }
        } else if (`val` is Boolean) {
            doperRec[0] = 0x8
            doperRec[2] = 0    // fError
            doperRec[3] = (if (`val`.booleanValue()) 1 else 0).toByte()
            doper2 = ErrorDoper(doperRec)
        } else {   // assume a Number object
            doperRec[0] = 0x2
            try {
                var d = 0.0
                if (`val` is Double)
                    d = `val`.toDouble()
                else if (`val` is Int)
                    d = `val`.toDouble()
                else
                    throw NumberFormatException("Unable to convert to Numeric Object" + `val`.javaClass)
                doper2 = RKDoper(doperRec)
                (doper2 as RKDoper).setVal(d)
                //doper1= new IEEEDoper(doperRec);
            } catch (e: Exception) {
                Logger.logErr("AutoFilter.setVal: error setting value to $`val`:$e")
            }

        }
        doper2!!.comparisonOperator = op
        fSimple2 = "=" == op
        wJoin = !AND
        update()
    }

    /**
     * sets this AutoFilter to be a Top-n type of filter
     * <br></br>Top-n filters only show the Top n values or percent in the column
     * <br></br>n can be from 1-500, or 0 to turn off Top 10 filtering
     *
     * @param int     n - 0-500
     * @param boolean percent - true if show Top-n percent; false to show Top-n items
     * @param boolean top10 - true if show Top 10 (items or percent), false to show Bottom N (items or percent)
     */
    fun setTop10(n: Int, percent: Boolean, top10: Boolean) {
        if (n == 0) {
            fTop = false
            fTop10 = false
            fPercent = false
            wTop10 = 0
            // TODO: set fSimple1?  remove dopers??
        } else if (n > 0 && n <= 500) {
            fTop = top10
            fTop10 = true
            wTop10 = n.toShort()
            fPercent = percent
            fSimple1 = false    // true if 1st condition is simple equality
            fSimple2 = false    // true if 2nd condition is simple equality
            doper1 = IEEEDoper(byteArrayOf(4, 6, 0, 0, 0, 0, 0, 0, 0, 0))
            (doper1 as IEEEDoper).setVal(n.toDouble())
            doper2 = UnusedDoper(byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
        } else
            Logger.logErr("AutoFilter.setTop10: value $n must be between 0 and 500")
        update()
    }

    /**
     * sets this AutoFilter to filter all blank rows
     */
    fun setFilterBlanks() {
        fSimple1 = true
        fSimple2 = false
        fTop = false
        fTop10 = false
        fPercent = false
        wTop10 = 0
        doper1 = NoBlanksDoper(byteArrayOf(14, 5, 0, 0, 0, 0, 0, 0, 0, 0))
        doper2 = UnusedDoper(byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
        update()
    }

    /**
     * sets this AutoFilter to filter all non-blank rows
     */
    fun setFilterNonBlanks() {
        fSimple1 = true
        fSimple2 = false
        fTop = false
        fTop10 = false
        fPercent = false
        wTop10 = 0
        doper1 = AllBlanksDoper(byteArrayOf(12, 2, 0, 0, 0, 0, 0, 1, 0, 0))
        doper2 = UnusedDoper(byteArrayOf(0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
        update()
    }


    /**
     * Doper (Database OPER Structure) are parsed definitions of the AutoFilter
     *
     *
     * 10-byte parsed definitions that appear in the Custom AutoFilter dialog box
     *
     * There are several sub-types of Dopers:
     * <br></br>String
     * <br></br>IEEE Floating Point
     * <br></br>RK
     * <br></br>Error or Boolean
     * <br></br>All Blanks
     * <br></br>All non-blanksfs
     * <br></br>Unused -- placeholder doper
     *
     * @see Excel bible page 284
     */
    private open inner class Doper protected constructor(rec: ByteArray) {
        var vt: Byte = 0
        var grbitSign: Byte = 0
        var record: ByteArray
            internal set

        /**
         * Returns the comparison operator for this doper, ie '>' '=', etc
         *
         * @return
         */
        /**
         * sets the custom comparison operator to one of:
         * <br></br>"<"
         * <br></br>"="
         * <br></br>"<="
         * <br></br>">"
         * <br></br>"<>"
         * <br></br>">="
         *
         * @param String op - custom operator
         */
        var comparisonOperator: String
            get() {
                when (grbitSign) {
                    1 -> return "<"
                    2 -> return "="
                    3 -> return "<="
                    4 -> return ">"
                    5 -> return "<>"
                    6 -> return ">="
                }
                return ""
            }
            set(op) {
                if ("<" == op)
                    grbitSign = 1
                if ("=" == op)
                    grbitSign = 2
                if ("<=" == op)
                    grbitSign = 3
                if (">" == op)
                    grbitSign = 4
                if ("<>" == op)
                    grbitSign = 5
                if (">=" == op)
                    grbitSign = 6
                record[1] = grbitSign
            }

        init {
            vt = rec[0]
            grbitSign = rec[1]    // comparison code
            record = rec
        }

        override fun toString(): String? {
            return null
        }
    }

    /**
     * A doper representing an unused value/filter condition unused.
     */
    private inner class UnusedDoper(rec: ByteArray) : Doper(rec) {

        override fun toString(): String? {
            return "Unused"
        }

    }

    /**
     * A doper representing an RK number
     *
     * Dopers define an AutoFilter value using a 10-byte doperRec
     * <br></br>For all dopers, doperRec[0]=vt or code
     * <br></br>doperRec[1]=comparison operator
     *
     * For RK Dopers,
     * doperRec[2]->[6] = rk number, 6-9= reserved
     */
    private inner class RKDoper(rec: ByteArray) : Doper(rec) {
        internal var `val`: Double = 0.toDouble()

        init {
            val b = ByteArray(4)
            System.arraycopy(record, 2, b, 0, 4)
            `val` = Rk.parseRkNumber(b)    // parse bytes in Rk-number format into a double value
        }

        override fun toString(): String? {
            return `val`.toString()
        }

        /**
         * set the double value for this Rk-type Doper record
         *
         * @param double d - double value to set
         */
        fun setVal(d: Double) {
            `val` = d
            val b = Rk.getRkBytes(d)
            System.arraycopy(b, 0, record, 2, 4)
        }

        /**
         * return the double value associated with this Rk-type Doper record
         *
         * @return double value
         */
        fun getVal(): Double {
            return `val`
        }
    }

    /**
     * A doper representing a IEEE number
     *
     * Dopers define an AutoFilter value using a 10-byte doperRec
     * <br></br>For all dopers, doperRec[0]=vt or code
     * <br></br>doperRec[1]=comparison operator
     *
     * For IEEE Dopers,
     * doperRec[2->9] = IEEE floating point number
     */
    private inner class IEEEDoper
    /**
     * create an IEEEDoper Object from doper record bytes
     * (10 bytes as part of the AutoFilter record)
     *
     * @param byte[] rec 10 byte doper record
     */
    (rec: ByteArray) : Doper(rec) {
        internal var `val`: Double = 0.toDouble()

        init {
            val b = ByteArray(8)
            System.arraycopy(record, 2, b, 0, 8)
            `val` = ByteTools.eightBytetoLEDouble(b)        // TODO: is this correct??
        }

        /**
         * return the double value associated with this IEEE-type Doper record
         *
         * @return double value
         */
        fun getVal(): Double {
            return `val`
        }

        /**
         * set the double value for this IEEE-type Doper record
         *
         * @param double d - double value to set
         */
        fun setVal(d: Double) {
            `val` = d
            val b = ByteTools.doubleToLEByteArray(`val`)
            System.arraycopy(b, 0, record, 2, 8)
        }

        override fun toString(): String? {
            return `val`.toString()
        }
    }

    /**
     * A doper representing a String
     *
     * Dopers define an AutoFilter value using a 10-byte doperRec
     * <br></br>For all dopers, doperRec[0]=vt or code
     * <br></br>doperRec[1]=comparison operator
     *
     * For String Dopers,
     * doperRec[6]= cch or String length
     * <br></br>The actual string data is appended to the end of the entire AutoFilter data record
     * <br></br>
     * NOTE: strings must be in normal or default encoding
     */
    private inner class StringDoper(rec: ByteArray) : Doper(rec) {
        internal var s: String

        /**
         * returns cch, the length of this string data
         *
         * @return int
         */
        val cch: Int
            get() = record[6].toInt()

        /**
         * returns the byte array representing the String referenced by this StringDoper
         *
         * @return byte[]
         */
        val strBytes: ByteArray
            get() = s.toByteArray()

        /**
         * sets the string for this StringDoper from a byte array and an index into said array
         * <br></br>amount to read from byteArray is stored in cch, bit 6 of the 10-byte doperRec
         *
         * @param byte[] rgch - source byte array for string
         * @param int    start - start index into the byte array
         */
        fun setString(rgch: ByteArray?, start: Int) {
            val cch = record[6].toInt()
            val stringbytes = ByteArray(cch)
            System.arraycopy(rgch!!, start, stringbytes, 0, cch)
            s = String(stringbytes)
        }

        /**
         * s the string for this StringDoper from a String Object
         *
         * @param String s
         */
        fun setString(s: String) {
            this.s = s
            val b = s.toByteArray()
            record[6] = b.size.toByte()
        }

        /**
         * returns the String referenced by this StringDoper
         */
        override fun toString(): String? {
            return s
        }
    }

    /**
     * A doper representing an Error or Boolean
     *
     * Dopers define an AutoFilter value using a 10-byte doperRec
     * <br></br>For all dopers, doperRec[0]=vt or code
     * <br></br>doperRec[1]=comparison operator
     *
     * For Error or Boolean Dopers,
     * doperRec[2]= fError;
     * if 0, doperRec[3]= boolean value
     * 				 if 1, doperRec[3]= Error value
     * Error Values:
     * <br></br>	0	= #NULL!
     * <br></br>	0x7	= #DIV/0!
     * <br></br>	0xF	= #VALUE!
     * <br></br>	0x17= #REF!
     * <br></br>	0x1D= #NAME?
     * <br></br>	0x24= #NUM!
     * <br></br>	0x2A= #N/A
     */
    private inner class ErrorDoper(rec: ByteArray) : Doper(rec) {
        /**
         * returns the boolean value if this is a type Boolean Doper
         *
         * @return boolean
         */
        var booleanVal: Boolean = false
            internal set
        internal var errVal: Int = 0

        /**
         * returns true if this is a Error Doper
         *
         * @return boolean true if this is a Error Doper
         */
        val isErrVal: Boolean
            get() = record[2].toInt() == 1

        /**
         * returns true if this is a Boolean Doper
         *
         * @return boolean true if this is a boolean Doper
         */
        val isBooleanVal: Boolean
            get() = record[2].toInt() == 0

        init {
            if (record[2].toInt() == 0)
            // boolean doper
                booleanVal = record[3].toInt() != 0
            else
                errVal = record[3].toInt()    // see above vals
        }

        /**
         * interprets the error code located in doperRec[2] into a String Error Value
         * e.g. "#NULL!"
         *
         * @return String error Value
         */
        fun getErrVal(): String {
            if (record[2].toInt() == 1) {
                when (errVal) {
                    0 -> return "#NULL!"
                    0x7 -> return "#DIV/0!"
                    0xF -> return "#VALUE!"
                    0x17 -> return "#REF!"
                    0x1D -> return "#NAME?"
                    0x24 -> return "#NUM!"
                    0x2A -> return "#N/A"
                }
            }
            return ""
        }


        /**
         * returns a String representation of this Error or Boolean Doper
         * <br></br>If this is an Error Doper, returns one of the Error Values
         * <br></br>If this is a booelan Doper, returns "true" or "false"
         *
         * @return String representation
         * @see AutoFilter.getErrVal
         */
        override fun toString(): String? {
            if (isErrVal)
                return getErrVal()
            return if (booleanVal) "true" else "false"
        }
    }

    /**
     * A doper representing an all blanks selection.
     * <br></br>bytes are unused other than identifier
     */
    private inner class AllBlanksDoper(rec: ByteArray) : Doper(rec) {

        override fun toString(): String? {
            return "All Blanks"
        }
    }

    /**
     * A doper representing an all blanks selection.
     * <br></br>bytes are unused other than identifier
     */
    private inner class NoBlanksDoper(rec: ByteArray) : Doper(rec) {

        override fun toString(): String? {
            return "No Blanks"
        }
    }

    companion object {

        private val serialVersionUID = -5228830347211523997L

        /**
         * creates a new AutoFilter record
         */
        // 20090630 KSC: don't use static PROTOTYPE_BYTES as changes are propogated
        val prototype: XLSRecord?
            get() {
                val af = AutoFilter()
                af.opcode = XLSConstants.AUTOFILTER
                af.setData(af.PROTOTYPE_BYTES)
                af.init()
                return af
            }
    }


}
