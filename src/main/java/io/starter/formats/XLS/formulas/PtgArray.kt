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
package io.starter.formats.XLS.formulas

import io.starter.formats.XLS.ExpressionParser
import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger

import java.io.UnsupportedEncodingException

/*
 * ARRAY CONSTANT followed by 7 reserved bytes.
 *
 * The token value for ptgArray consists of the array
 * dimensions and the array values
 *
 * ptgArray differs from most other operand tokens in that
 * the token value doesn't follow the token type.
 *
 * Instead, the token value is appended to the saved parsed
 * expression, immediately following the last token.
 *
 * Offset Name Size Contents
 * ---------------------------------------------------------
 * --
 * 0 nc 1 number of columns -1 in array constant (0 = 256)
 * 1 nr 2 number of rows -1 in array constant
 * 3 rgval var the array vals (k+1)*(nr+1) length
 *
 *
 * The format of the token value is shown in the following
 * table.
 *
 * The number of values in the array constant is equal to
 * the product of the array dimensions, (nc+1)*(nr+1_
 *
 * Each value is either an 8-byte IEEE fp numbr or a string.
 * The two formats for these values are shown in the
 * following tables.
 *
 *
 * IEEE FP Number
 * Offset Name Size Contents
 * ---------------------------------------------------------
 * --
 * 0 grbit 1 =01h
 * 1 num 8 IEEE FP number
 *
 * String
 * Offset Name Size Contents
 * ---------------------------------------------------------
 * --
 * 0 grbit 1 =02h
 * 1 cch 1 Length of the String
 * 2 rgch var the string.
 *
 * If a formula contains more than one array constant, the
 * token values for the array constants are appended to the
 * saved
 * parsed expression in order: first the values for the
 * first array constant,
 * then the values for the second array constant, etc.
 *
 * If a formula contains very long array constants, the
 * FORMULA, ARRAY, or NAME record contaniing the parsed
 * expression
 * may overflow into CONTINUE records. In such cases, an
 * individual array value is NEVER SPLIT between records,
 * but record boundaries are established between adjacent
 * array values.
 *
 * The reference class ptgArray never appears in an Excel
 * formula, only the ptgArrayV and ptgArrayA classes are
 * used.
 *
 *
 * @see Ptg
 *
 * @see Formula
 *
 */
// 20090119-22 KSC: Many, many changes changes
class PtgArray : GenericPtg(), Ptg {
    /**
     * returns the 0-based number of columns in this array
     * number of columns is the amount of elements before the semicolon (if present)
     * a,b,c;d,e,f; ....
     *
     * @return
     */
    var numberOfColumns = -1
        internal set
    /**
     * returns the 0-based number of rows in this array
     * if nr>1 then the array is in the form of:
     * a,b,c;d,e,f; .... where the semicolons delineate rows
     *
     * @return
     */
    var numberOfRows = -1
        internal set
    /**
     * these are the bytes appended to the formula token array, after all other ptg's
     *
     * @return
     */
    var postRecord: ByteArray? = null
        internal set
    internal var arrVals = CompatibleVector()
    internal var isIntermediary = false                    // 20090824
    // KSC:
    // true
    // if
    // this
    // PtgArray
    // is
    // only
    // part
    // of
    // a
    // calcualtion
    // process;
    // if
    // so,
    // apparently
    // can
    // have
    // more
    // than
    // 256
    // columns
    // [BugTracker
    // 2683]

    override val isOperand: Boolean
        get() = true

    /**
     * return the first 8 bytes of the ptgArray record
     * this represents the id byte and 7 reserved bytes
     *
     * @return
     */
    val preRecord: ByteArray
        get() = record

    val `val`: Int
        get() = -1

    override// 20090820 KSC: value = entire array instead of 1st value;
    // desired value is determined by cell position as compared
    // to current formula; see Formula.calculate [BugTracker
    // 2683]
    // return elementAt(0).getValue(-1); // default: return 1st
    // value
    val value: Any?
        get() = string

    /*
     * returns the string value of the name
     *
     * @see io.starter.formats.XLS.formulas.Ptg#getValue()
     */
    override// if it's a single value, just return val
    // retstr += "}";
    // retVal = retstr.substring(0,retstr.length()-1);
    val string: String
        get() {
            var retVal: Any? = null
            val p = this.components
            var retstr = ""
            if (numberOfColumns == 0 && numberOfRows == 0) {
                for (i in p!!.indices) {
                    if (i != 0)
                        retstr += ","
                    retstr += p[i].value.toString()
                }
            } else {
                retstr = ""
                var loc = 0
                for (x in 0 until numberOfRows + 1) {
                    if (x != 0)
                        retstr += ";"
                    for (i in 0 until numberOfColumns + 1) {
                        if (i != 0)
                            retstr += ","
                        retstr += p!![loc++].value.toString()
                    }
                }
            }
            retVal = retstr
            return "{$retVal}"
        }

    override val textString: String
        get() = string

    /**
     * Returns the second section of bytes for the PtgArray.
     * These are the bytes that are split off the end of the
     * formula
     *
     * public void getComponentBytes(){
     *
     * }
     * //public void updateRecord(){
     * //} */

    /**
     * Override due to mystery extra byte
     * occasionally found in ptgName recs.
     */
    override/*
         * 20090820 KSC: really want record length not rgval length,
         * which now is separate [BugTracker 2683]
         * if (rgval!=null)
         * return rgval.length;
         */ val length: Int
        get() = 8

    override// it's a range
    val components: Array<Ptg>?
        get() {
            var retVals = arrayOfNulls<Ptg>(arrVals.size)
            for (i in arrVals.indices) {
                val o = arrVals.elementAt(i)
                if (o is Double) {
                    val pnum = PtgNumber(o.toDouble())
                    retVals[i] = pnum
                } else if (o is Boolean) {
                    val pb = PtgBool(o.booleanValue())
                    retVals[i] = pb
                } else {
                    if (FormulaParser.isRef(o as String) || FormulaParser.isRange(o)) {
                        val pa = PtgArea3d()
                        pa.parentRec = this.parentRec
                        pa.useReferenceTracker = true
                        pa.location = o
                        val pacomps = pa.components
                        val temp = arrayOfNulls<Ptg>(retVals.size - 1 + pacomps!!.size)
                        System.arraycopy(retVals, 0, temp, 0, retVals.size - 1)
                        System.arraycopy(pacomps, 0, temp, retVals.size - 1, pacomps.size)
                        retVals = temp
                    } else {
                        val pstr = PtgStr(o)
                        retVals[i] = pstr
                    }
                }
            }
            return retVals
        }

    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        this.populateVals()
    }

    private fun populateVals() {
        if (record.size > 8) { // means that array data has already been
            // appeneded to end of record array; store
            // in rgvals
            postRecord = ByteArray(record.size - 8)
            System.arraycopy(record, 8, postRecord!!, 0, postRecord!!.size) // save post
            // array=
            // nc, nr +
            // array
            // data
        }
        if (postRecord != null) {
            // clear out record array:0= id 1-7=reserved
            val b = ByteArray(8)
            b[0] = record[0]
            record = b
            this.parseArrayComponents()
        } // otherwise, it's just the initial input of the 1st 8 bytes
        // record - see Formula
    }

    /**
     * given "extra info" at end of formula expression, parse array values
     */
    fun parseArrayComponents() {
        var nitems = 0
        arrVals.clear() // 20090820 KSC: makes sense to! [BugTracker 2683]
        if (!isIntermediary)
        // 20090824 KSC: sometimes an intermediary ptgarry
        // can have more than 256 columns [BugTracker
        // 2683]
            numberOfColumns = postRecord!![0] and 0xFF // number of columns
        numberOfRows = ByteTools.readShort(postRecord!![1].toInt(), postRecord!![2].toInt()).toInt() // number of rows
        try {
            // (nc+1)*(nr+1) compoments
            var i = 3
            while (i < postRecord!!.size) { // 20090820 KSC: post array
                // contains nc & nr so i
                // should be initially 3
                // instead of 0 [BugTracker
                // 2683]
                if (postRecord!![i].toInt() == 0) { // empty value
                    i++
                    i += 8
                    arrVals.add("") // TODO: Empty Constant should be null?
                } else if (postRecord!![i].toInt() == 0x1) { // its a number
                    i++
                    val barr = ByteArray(8)
                    System.arraycopy(postRecord!!, i, barr, 0, 8)
                    val `val` = ByteTools.eightBytetoLEDouble(barr)
                    val d = `val`
                    arrVals.add(d)
                    i = i + 8
                } else if (postRecord!![i].toInt() == 0x2) { // its a string
                    val strLen = ByteTools
                            .readShort(postRecord!![i + 1].toInt(), postRecord!![i + 2].toInt()).toInt()
                    i += 3
                    val grbt = postRecord!![i++].toInt()
                    val barr = ByteArray(strLen)
                    System.arraycopy(postRecord!!, i, barr, 0, strLen)
                    var strVal = ""
                    try {
                        if (grbt and 0x1 == 0x1) {
                            strVal = String(barr, XLSConstants.UNICODEENCODING)
                        } else {
                            strVal = String(barr, XLSConstants.DEFAULTENCODING)
                        }
                    } catch (e: UnsupportedEncodingException) {
                        Logger.logInfo("decoding formula string in array failed: $e")
                    }

                    arrVals.add(strVal)
                    i += strLen
                } else if (postRecord!![i].toInt() == 0x4) { // its a boolean
                    if (postRecord!![++i].toInt() == 0)
                        arrVals.add(java.lang.Boolean.valueOf(false))
                    else
                        arrVals.add(java.lang.Boolean.valueOf(true))
                    i = i + 8
                } else if (postRecord!![i].toInt() == 0x10) { // it's an error value
                    val errCode = postRecord!![++i].toInt()
                    when (errCode) {
                        0 -> arrVals.add("#NULL!")
                        0x7 -> arrVals.add("#DIV/0!")
                        0x0F -> arrVals.add("#VALUE!")
                        0x17 -> arrVals.add("#REF!")
                        0x1D -> arrVals.add("#NAME!")
                        0x24 -> arrVals.add("#NUM!")
                        0x2A -> arrVals.add("#N/A!")
                    }
                    i = i + 8
                }
                nitems++
                if (nitems == (numberOfColumns + 1) * (numberOfRows + 1)) { // Finished with this
                    // array!
                    val length = i
                    i = postRecord!!.size
                    // length may be less than rgval.length for cases of more
                    // than one array parameter
                    // see ExpressionParser.parseExpression
                    if (postRecord!!.size != length) {// then truncate both record +
                        // rgval
                        val tmp = ByteArray(length)
                        System.arraycopy(postRecord!!, 0, tmp, 0, length)
                        postRecord = tmp
                    }
                }
            }
        } catch (e: Exception) {
            Logger.logErr("Error Processing Array Formula: $e")
            return
        }

    }

    fun setVal(arrStr: String) {
        var arrStr = arrStr
        // remove the initial { and ending }
        arrStr = arrStr.substring(1, arrStr.length - 1)
        if (arrStr.indexOf("{") != -1) { // SHOULDN'T -- see
            // FormulaParser.getPtgsFromFormulaString
            Logger.logErr("PtgArray.setVal: Multiple Arrays Encountered")
        }

        // parse all array strings into rows, cols
        var rows: Array<String>? = null
        var cols: Array<Array<String>>? = null
        // split rows
        rows = arrStr.split(";".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
        cols = arrayOfNulls(rows!!.size)
        for (i in rows!!.indices) {
            val s = rows[i].split(",".toRegex()).toTypedArray() // include empty strings
            cols[i] = s
        }
        var databytes = ByteArray(11)
        databytes[0] = 0x60 // 20h=tArrayR, 40h=tArrayV, 60h=tArrayA
        isIntermediary = false // init value
        if (cols[0].size >= 255) { // 20090824 KSC: apparently sometimes an
            // intermediary calculations step can
            // include > 256 array elements ...
            isIntermediary = true
            numberOfColumns = cols[0].size - 1
        }
        databytes[8] = (cols[0].size - 1 and 0xFF).toByte() // nc-1 // 20090819
        // KSC: placed
        // in wrong pt
        // of record:
        // was [1]
        // [BugTracker
        // 2683]
        // databytes[8] = (byte)((cols[0].length-1)); // nc-1 //
        // 20090819 KSC: placed in wrong pt of record: was [1]
        // [BugTracker 2683]
        System.arraycopy(ByteTools
                .shortToLEBytes((rows.size - 1).toShort()), 0, databytes, 9, 2) // nr-1
        // //
        // 20090819
        // KSC:
        // placed
        // in
        // wrong
        // pt
        // of
        // record:
        // was
        // 2,3
        // [BugTracker
        // 2683]
        // iterate the array and fill out the data section
        for (j in rows.indices) {
            for (i in 0 until cols[0].size) {
                val valbytes = this.valuesIntoByteArray(cols[j][i])
                databytes = ByteTools.append(valbytes, databytes)
            }
        }
        // populate primary values for rec
        record = databytes
        this.init(databytes)
    }

    /**
     * Turns a vector of values into a byte array representation for the data section of this record
     *
     * @param compVect
     * @return
     */
    private fun valuesIntoByteArray(constVal: String?): ByteArray {
        var databytes = ByteArray(0)
        var thisElement = ByteArray(9)

        try { // number?
            val d = Double(constVal!!)
            thisElement[0] = 0x1 // id for number value
            val b = ByteTools.toBEByteArray(d)
            System.arraycopy(b, 0, thisElement, 1, b.size)
            databytes = ByteTools.append(thisElement, databytes)
        } catch (ee: NumberFormatException) {
            try {
                if (constVal!!.equals("true", ignoreCase = true) || constVal.equals("false", ignoreCase = true)) {
                    val bb = java.lang.Boolean.valueOf(constVal)
                    thisElement[0] = 0x4 // id for boolean value
                    thisElement[1] = (if (bb.booleanValue()) 1 else 0).toByte()
                } else if (constVal == null || constVal == "") { // emtpy
                    // or
                    // null
                    // value
                    thisElement[0] = 0x0 // id for empty value
                } else if (constVal[0] == '#') { // it's an error value
                    thisElement[0] = 0x10 // id for error value
                    var errCode = 0
                    if (constVal == "#NULL!")
                        errCode = 0
                    else if (constVal == "#DIV/0!")
                        errCode = 0x7
                    else if (constVal == "#VALUE!")
                        errCode = 0x0F
                    else if (constVal == "#REF!")
                        errCode = 0x17
                    else if (constVal == "#NAME!")
                        errCode = 0x1D
                    else if (constVal == "#NUM!")
                        errCode = 0x24
                    else if (constVal == "#N/A!" || constVal == "#N/A"
                            || constVal == "N/A")
                        errCode = 0x2A
                    thisElement[1] = errCode.toByte()
                } else { // assume string
                    thisElement = ByteArray(3)
                    try {
                        thisElement = ByteArray(4)
                        thisElement[0] = 0x2 // id for string
                        val b = constVal.toByteArray(charset(XLSConstants.UNICODEENCODING))
                        System.arraycopy(ByteTools
                                .shortToLEBytes(b.size.toShort()), 0, thisElement, 1, 2)
                        thisElement[3] = 1 // compressed= 0, uncompressed= 1
                        // (16-bit chars)
                        thisElement = ByteTools.append(b, thisElement)
                    } catch (z: UnsupportedEncodingException) {
                        Logger.logWarn("encoding formula array:$z")
                    }

                }
                databytes = ByteTools.append(thisElement, databytes)
            } catch (ex: Exception) {
                Logger.logWarn("PtgArray.valuesIntoByteArray:  error parsing array element:$ex")
            }

        }

        return databytes
    }

    /*
     * not used
     * public int getLength(byte[] b){
     * int co = b[1];
     * int rw = ByteTools.readShort(b[2], b[3]);
     * rw++; // appears that rows are not ordinal here...
     * int numrecs = co*rw;
     * int len = 4;
     * int loc = 4;
     * for (int i=0;i<=numrecs;i++){
     * if (b[len] == 0x1){ // its a number
     * len += 9;
     * }else{
     * len += b[len+1] + 2;
     * }
     * }
     * length = len;
     * return length;
     * }
     */

    override fun toString(): String {
        return this.string
    }

    /**
     * sets the array components values for this PtgArray
     * returns the actual array components length
     *
     * @see ExpressionParser.parseExpression
     */
    fun setArrVals(by: ByteArray): Int {
        postRecord = by
        if (postRecord != null) {
            // clear out record array:0= id 1-7=reserved
            val b = ByteArray(8)
            b[0] = record[0]
            record = b
            this.parseArrayComponents()
        }
        return postRecord!!.size

    }

    fun getArrVals(): ByteArray? {
        return postRecord
    }

    /**
     * returns a ptg at the specified location.  Assumes that it is a one-dimensional
     * array.  If you need a multidimensional array please use the other elementAt(int,int)method
     *
     * @param loc
     * @return
     */
    fun elementAt(loc: Int): Ptg {
        val p = this.components
        return p!![loc]
    }

    fun elementAt(col: Int, row: Int): Ptg? {
        val p = this.components
        try {
            var loc = 0
            for (i in 0 until row) {
                loc += numberOfColumns // 20090816 KSC: why +1???? +1); [BugTracker 2683]
            }
            loc += col
            return elementAt(loc)
        } catch (e: ArrayIndexOutOfBoundsException) {
            Logger.logErr("PtgArray.elementAt: error retrieving value at ["
                    + row + "," + col + "]: " + e)
        }

        return null
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 4416140231168551393L
    }

}
