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

import io.starter.OpenXLS.ExcelTools
import io.starter.formats.XLS.*
import io.starter.toolkit.Logger

import java.util.ArrayList

/**
 * LookupReferenceCalculator is a collection of static methods that operate
 * as the Microsoft Excel function calls do.
 *
 *
 * All methods are called with an array of ptg's, which are then
 * acted upon.  A Ptg of the type that makes sense (ie boolean, number)
 * is returned after each method.
 */

object LookupReferenceCalculator {

    /**
     * ADDRESS
     * Creates a cell address as text, given specified row and column numbers.
     *
     *
     * Syntax
     * ADDRESS(row_num,column_num,abs_num,a1,sheet_text)
     * Row_num   is the row number to use in the cell reference.
     * Column_num   is the column number to use in the cell reference.
     * Abs_num   specifies the type of reference to return.
     *
     *
     * Abs_num	 Returns this type of reference
     * 1 or omitted	 Absolute
     *
     *
     * 2 	 Absolute row; relative column
     *
     *
     * 3	 Relative row; absolute column
     *
     *
     * 4	 Relative
     *
     *
     * A1   is a logical value that specifies the A1 or R1C1 reference style.
     * If a1 is TRUE or omitted, ADDRESS returns an A1-style reference;
     * if FALSE, ADDRESS returns an R1C1-style reference.
     *
     *
     * Sheet_text   is text specifying the name of the worksheet to be
     * used as the external reference. If sheet_text is omitted, no sheet name is used.
     *
     *
     * Examples
     *
     *
     * ADDRESS(2,3) equals "$C$2"
     *
     *
     * ADDRESS(2,3,2) equals "C$2"
     *
     *
     * ADDRESS(2,3,2,FALSE) equals "R2C[3]"
     *
     *
     * ADDRESS(2,3,1,FALSE,"[Book1]Sheet1") equals "[Book1]Sheet1!R2C3"
     *
     *
     * ADDRESS(2,3,1,FALSE,"EXCEL SHEET") equals "'EXCEL SHEET'!R2C3"
     */
    internal fun calcAddress(operands: Array<Ptg>): Ptg {
        if (operands.size < 2) return PtgErr(PtgErr.ERROR_VALUE)
        // deal with floating point refs
        var rx1 = operands[0].value.toString()
        if (rx1.indexOf(".") > -1)
            rx1 = rx1.substring(0, rx1.indexOf("."))
        val row = Integer.valueOf(rx1).toInt()
        // deal with floating point refs
        var cx1 = operands[1].value.toString()
        if (cx1.indexOf(".") > -1)
            cx1 = cx1.substring(0, cx1.indexOf("."))

        val col = Integer.valueOf(cx1).toInt()
        var abs_num = 1
        var ref_style = true
        var sheettext = ""
        if (operands.size > 2) {
            if (operands[2].value != null) { //checking for a ptgmissarg
                abs_num = (operands[2].value as Int).toInt()
            }
        }
        if (operands.size > 3) {
            if (operands[3].value != null) { //checking for a ptgmissarg
                val b = java.lang.Boolean.valueOf(operands[3].value.toString())
                ref_style = b.booleanValue()
            }
        }
        if (operands.size > 4) {
            if (operands[4].value != null) { //checking for a ptgmissarg
                sheettext = operands[4].value.toString() + "!"
            }
        }
        var loc = ""
        val colstr = ExcelTools.getAlphaVal(col - 1)
        if (ref_style) {
            if (abs_num == 1) {
                loc = "$$colstr$$row"
            } else if (abs_num == 2) {
                loc = "$colstr$$row"
            } else if (abs_num == 3) {
                loc = "$$colstr$row"
            } else if (abs_num == 4) {
                loc = colstr + row
            }
        } else {
            if (abs_num == 1) {
                loc = "R" + row + "C" + col // this is transposed with abs_num 4.  Error in Excel
            } else if (abs_num == 2) {
                loc = "R" + row + "C[" + col + "]"
            } else if (abs_num == 3) {
                loc = "R[$row]C$col"
            } else if (abs_num == 4) {
                loc = "R[$row]C[$col]"

            }
        }
        loc = sheettext + loc
        return PtgStr(loc)

    }

    /**
     * AREAS
     * Returns the number of areas in a reference. An area is a range of contiguous cells or a single cell.
     *
     *
     * Reference   is a reference to a cell or range of cells and can refer to multiple areas.
     * If you want to specify several references as a single argument, then you must include extra sets of parentheses so that Microsoft Excel will not interpret the comma as a field separator.
     *
     *
     * NOTE: this appears to be correct given Excel information but logic is not 100% known
     * e.g. =AREAS(B2:D4 D3) = 1
     * =AREAS(B2:D4 E3) gives a #NULL! error - why?
     * =AREAS(B2:D4,E3) = 2
     */
    internal fun calcAreas(operands: Array<Ptg>): Ptg {
        val ref = operands[0]
        val s = ref.toString()
        val areas = s.split(",(?=([^'|\"]*'[^'|\"]*'|\")*[^'|\"]*$)".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
        return PtgNumber(areas.size.toDouble())
    }

    /**
     * CHOOSE
     * Chooses a value from a list of values
     *
     *
     * Note, this function does not support one specific use-case.  That is choosing a ptgref
     * and using that ptgref to complete a ptgarea.  Example
     * =SUM(E6:CHOOSE(3,G4,G5,G6))
     * =SUM(E6:G6)
     */
    internal fun calcChoose(operands: Array<Ptg>): Ptg {
        if (operands.size < 2) return PtgErr(PtgErr.ERROR_VALUE)
        var o = operands[0].value
        try {
            val dd = Double(o.toString()) // this can be non-integer, so truncate it if so...
            val i = dd.toInt()
            if (i > operands.size + 1 || i < 1) return PtgErr(PtgErr.ERROR_REF)
            o = operands[i].value
            val d = o as Double
            return PtgNumber(d)
        } catch (ex: Exception) {
            PtgErr(PtgErr.ERROR_VALUE)
        }

        return PtgStr(o.toString())

    }

    /**
     * COLUMN
     * Returns the column number of a reference
     */
    internal fun calcColumn(operands: Array<Ptg>): Ptg {
        if (operands[0] is PtgFuncVar) {
            // we need to return the col where the formula is.
            val pfunk = operands[0] as PtgFuncVar
            try {
                var loc = pfunk.parentRec!!.colNumber.toInt()
                loc += 1
                return PtgInt(loc)
            } catch (e: Exception) {
            }

        } else {
            // It's ugly, but we are going to handle the four types of references seperately, as there is no good way
            // to generically get this info
            try {
                if (operands[0] is PtgArea) {
                    val pa = operands[0] as PtgArea
                    val loc = pa.intLocation
                    return PtgInt(loc!![1] + 1)
                } else if (operands[0] is PtgRef) {
                    val pref = operands[0] as PtgRef
                    var loc = pref.intLocation!![1]
                    loc += 1
                    return PtgInt(loc)
                } else if (operands[0] is PtgName) {    // table???
                    val range = (operands[0] as PtgName).name!!.location
                    val loc = ExcelTools.getRangeCoords(range)
                    return PtgInt(loc[1] + 1)
                }
            } catch (e: Exception) {
            }

        }
        return PtgInt(-1)
    }

    /**
     * COLUMNS
     * Returns the number of columns in an array reference or array formula
     */
    // TODO: Not finished yet!
    internal fun calcColumns(operands: Array<Ptg>): Ptg {
        //
        if (operands[0] is PtgFuncVar) {
            // we need to return the col where the formula is.
            val pfunk = operands[0] as PtgFuncVar
            try {
                var loc = pfunk.parentRec!!.colNumber.toInt()
                loc += 1
                return PtgInt(loc)
            } catch (e: Exception) {
            }

        } else {
            // It's ugly, but we are going to handle the four types of references seperately, as there is no good way
            // to generically get this info
            try {
                if (operands[0] is PtgArea) {
                    val pa = operands[0] as PtgArea
                    val loc = pa.intLocation
                    val ncols = loc!![3] - loc[1] + 1
                    return PtgInt(ncols)
                } else if (operands[0] is PtgArray) {
                    val parr = operands[0] as PtgArray
                    return PtgInt(parr.numberOfColumns + 1)
                }
            } catch (e: Exception) {
            }

        }
        return PtgInt(-1)
    }

    /**
     * HLOOKUP: Looks in the top row of an array and returns the value of the
     * indicated cell.
     */
    @Throws(FunctionNotSupportedException::class)
    internal fun calcHlookup(operands: Array<Ptg>): Ptg {

        if (operands.size < 3) return PtgErr(PtgErr.ERROR_VALUE)
        var sorted = true
        var isNumber = true
        val lookup_value = operands[0]
        val table_array = operands[1]
        val row_index_num = operands[2] as PtgInt
        val rowNum = row_index_num.`val` - 1// reduce 1 for ordinal base off firstcol
        if (operands.size > 3) {
            if (operands[3].value != null) {
                val o = operands[3].value
                if (o is Boolean)
                    sorted = o.booleanValue()
                else if (o is Int)
                    sorted = o.toInt() != 0
            }
        }
        val retarea = intArrayOf(0, 0)
        val bs = lookup_value.parentRec.sheet
        val bk = table_array.parentRec.workBook
        var lookupComponents: Array<PtgRef>? = null
        var valueComponents: Array<PtgRef>? = null
        // first, get the lookup Column Vals
        if (table_array is PtgName) {
            // Handle getting vals out of name

        } else if (table_array is PtgArea || table_array is PtgArea3d) {
            try {
                val pa = table_array as PtgArea
                val range = table_array.intLocation

                //				 TODO: check rc sanity here
                val firstrow = range[0]
                lookupComponents = pa.getRowComponents(firstrow) as Array<PtgRef>
                valueComponents = pa.getRowComponents(firstrow + rowNum) as Array<PtgRef>
            } catch (/*20070209 KSC: FormulaNotFound*/e: Exception) {
                Logger.logWarn("Error in LookupReferenceCalculator: Cannot determine PtgArea location. $e")
            }

        }
        // error check
        if (lookupComponents == null || lookupComponents.size == 0) return PtgErr(PtgErr.ERROR_REF)
        // lets check if we are dealing with strings or numbers....
        try {
            val `val` = lookupComponents[0].value!!.toString()
            val d = Double(`val`)
        } catch (e: NumberFormatException) {
            isNumber = false
        }

        if (isNumber) {
            val match_num: Double
            try {
                match_num = java.lang.Double.parseDouble(
                        lookup_value.value.toString())
            } catch (e: NumberFormatException) {
                return PtgErr(PtgErr.ERROR_NA)
            }

            for (i in lookupComponents.indices) {
                val `val`: Double
                try {
                    `val` = java.lang.Double.parseDouble(
                            lookupComponents[i].value!!.toString())
                } catch (e: NumberFormatException) {
                    // Ignore entries in the table that aren't numbers.
                    continue
                }

                if (`val` == match_num) {
                    return valueComponents!![i].ptgVal
                } else if (sorted && `val` > match_num) {
                    return if (i == 0)
                        PtgErr(PtgErr.ERROR_NA)
                    else
                        valueComponents!![i - 1].ptgVal
                }
            }

            return if (sorted)
                valueComponents!![lookupComponents.size - 1].ptgVal
            else
                PtgErr(PtgErr.ERROR_NA)

        } else {
            //TODO: need to handle as string
        }

        return PtgErr(PtgErr.ERROR_NULL)
    }

    /**
     * HYPERLINK
     * Creates a shortcut or jump that opens a document
     * stored on a network server, an intranet, or the Internet
     *
     *
     * Function just returns the "friendly name" of the link,
     * Excel doesn't appear to validate the url ...
     */

    internal fun calcHyperlink(operands: Array<Ptg>): Ptg {
        try {
            return if (operands.size == 2) PtgStr(operands[1].value.toString()) else PtgStr(operands[0].value.toString())
        } catch (e: Exception) {
            return PtgErr(PtgErr.ERROR_VALUE)
        }

    }

    /**
     * INDEX
     * Returns a value or the reference to a value from within a table or range.
     * There are two forms of the INDEX function: the array form and the reference form.
     *
     *
     * Array Form:
     * Returns the value of an element in a table or an array selected by the row and column number indexes.
     * Use the array form if the first argument to INDEX is an array constant.
     * INDEX(array,row_num,column_num)
     * Array   is a range of cells or an array constant.
     * If array contains only one row or column,
     * the corresponding row_num or column_num argument is optional.
     * If array has more than one row and more than one column, and only row_num or column_num is used,
     * INDEX returns an array of the entire row or column in array.
     * Row_num   selects the row in array from which to return a value. If row_num is omitted, column_num is required.
     * Column_num   selects the column in array from which to return a value. If column_num is omitted, row_num is required.
     *
     *
     * Reference Form:
     * Returns the reference of the cell at the intersection of a particular row and column.
     * If the reference is made up of nonadjacent selections, you can pick the selection to look in.
     * INDEX(reference,row_num,column_num,area_num)
     * Reference   is a reference to one or more cell ranges.
     * If you are entering a nonadjacent range for the reference, enclose reference in parentheses.
     * If each area in reference contains only one row or column, the row_num or column_num argument, respectively, is optional. For example, for a single row reference, use INDEX(reference,,column_num).
     * Row_num   is the number of the row in reference from which to return a reference.
     * Column_num   is the number of the column in reference from which to return a reference.
     * Area_num   selects a range in reference from which to return the intersection of row_num and column_num. The first area selected or entered is numbered 1, the second is 2, and so on. If area_num is omitted, INDEX uses area 1.
     *
     *
     * Given a BiffRec Range, choose the cell within referenced by the
     * row and column operands.
     *
     *
     * example:
     * =INDEX(G3:L8,6,4)
     * returns the value of the BiffRec at row 6, col 4 in the following table
     * which is 0.11
     *
     *
     * G		H		I		J		K		L
     * 1
     * 2
     * 3		0.007	0.005	0.003	0.002	0.002	0.001
     * 4		0.025	0.017	0.012	0.008	0.006	0.005
     * 5		0.062	0.044	0.032	0.023	0.018	0.015
     * 6		0.116	0.086	0.064	0.049	0.04	0.035
     * 7		0.171	0.13	0.101	0.082	0.07	0.062
     * 8		0.211	0.165	0.132	0.11	0.096	0.088
     */
    internal fun calcIndex(operands: Array<Ptg>): Ptg {
        if (operands.size < 1) return PtgErr(PtgErr.ERROR_VALUE)
        var o = operands[0]
        //  only 1st operand is required; changed below (see doc snippet below)
        /*If each area in reference contains only one row or column,
         * the row_num or column_num argument, respectively, is optional.
         * For example, for a single row reference, use INDEX(reference,,column_num).
         */
        var rowref: Any = 1    // defaults (1-based)
        var colref: Any = 1
        var areanum = 1.0
        var retarea: IntArray? = null
        var sht: String? = null
        try {
            if (operands.size > 1) {
                val rowrefp = operands[1]
                rowref = rowrefp.value
            }
            if (operands.size > 2) {
                val colrefp = operands[2]
                colref = colrefp.value
            }
            if (operands.size > 3) {
                areanum = operands[3].doubleVal
            }
            if (o is PtgArea) {
                retarea = o.intLocation
                sht = o.sheetName
            } else if (o is PtgName) {
                //Ptg[] p=((PtgName) o).getName().getComponents();	//CellRangePtgs();
                //o= p[0];
                val r = o.name!!.location
                retarea = ExcelTools.getRangeRowCol(r!!)
                sht = if (r!!.indexOf("!") == -1) null else r!!.substring(0, r!!.indexOf("!"))
            } else if (o is PtgMemFunc) {
                val ps = o.components
                areanum--
                if (areanum >= 0 && areanum < ps.size)
                    o = ps[areanum.toInt()]
                else
                    o = ps[0]
                retarea = (o as PtgArea).intLocation
                sht = o.sheetName
            } else if (o is PtgArray) {
                val ps = o.components
                areanum--
                if (areanum >= 0 && areanum < ps.size)
                    o = ps[areanum.toInt()]
                else
                    o = ps[0]
                retarea = (o as PtgArea).intLocation
                // TODO: Sheet!
            }

            // now should have the correct area
            if (retarea != null) {
                // really we can just use the first position to get a ref
                // then the second just checks bounds
                //				 TODO: check rc sanity here
                val rowoff = retarea[0]
                val coloff = retarea[1]
                val rowck = retarea[2] + 1
                val colck = retarea[3] + 1

                val dims = IntArray(2)
                try {
                    val rr: Int
                    if (rowref is Int) {
                        rr = rowref.toInt()
                    } else {
                        var rw = rowref.toString()
                        if (rw.indexOf(".") > -1) rw = rw.substring(0, rw.indexOf("."))
                        // string non Integer Chars...
                        rr = Integer.parseInt(rw)
                    }
                    val cr: Int
                    if (colref is Int) {
                        cr = colref.toInt()
                    } else {
                        var cl = colref.toString()
                        if (cl.indexOf(".") > -1) cl = cl.substring(0, cl.indexOf("."))
                        // string non Integer Chars...
                        cr = Integer.parseInt(cl)
                    }

                    if (rr > rowck || cr > colck) {
                        return PtgErr(PtgErr.ERROR_REF)
                    }
                    dims[0] = rr + rowoff - 1
                    dims[1] = cr + coloff - 1

                    // here's a nice new ref...
                    var refp = PtgRef()
                    if (o is PtgArea3d) refp = PtgRef3d()
                    refp.parentRec = o.parentRec
                    refp.useReferenceTracker = false
                    if (sht != null) refp.sheetName = sht
                    refp.setLocation(dims)
                    if (o is PtgArea3d) {
                        refp.location = o.sheetName + "!" + ExcelTools.formatLocation(dims)
                    } else {

                    }
                    return refp

                } catch (e: NumberFormatException) {
                    //Logger.logWarn("could not calculate INDEX function: " + o.toString() + ":" + e);
                    return PtgErr(PtgErr.ERROR_NULL)    // ERR or #VALUE ??
                }

            }
        } catch (e: Exception) {
            Logger.logWarn("could not calculate INDEX function: $o:$e")
        }

        return PtgErr(PtgErr.ERROR_NULL)

    }


    /**
     * INDIRECT
     * Returns a reference indicated by a text value
     *
     *
     * INDIRECT(ref_text,a1)
     *
     *
     * Ref_text   is a reference to a cell that contains an A1-style
     * reference, an R1C1-style reference, a name defined as a reference,
     * or a reference to a cell as a text string. If ref_text is not a valid
     * cell reference, INDIRECT returns the #REF! error value.
     *
     *
     * If ref_text refers to another workbook (an external reference),
     * the other workbook must be open. If the source workbook is not open,
     * INDIRECT returns the #REF! error value.
     *
     *
     * A1   is a logical value that specifies what type of reference is contained in the cell ref_text.
     *
     *
     * If a1 is TRUE or omitted, ref_text is interpreted as an A1-style reference.
     * If a1 is FALSE, ref_text is interpreted as an R1C1-style reference.
     */
    internal fun calcIndirect(operands: Array<Ptg>): Ptg {
        var operands = operands
        try {
            if (operands[0] is PtgStr) {
                val ps = operands[0] as PtgStr

                var locx: String? = ps.toString()
                // detect range
                if (!(FormulaParser.isRef(locx) || FormulaParser.isRange(locx))) {    // see if it's a named range
                    val nmx = ps.parentRec!!.workBook!!.getName(locx)
                    if (nmx != null) { // there is a named range
                        locx = nmx.location
                    } else
                        return ps    //it's just a value
                }
                if ("" == locx)
                    return PtgInt(0)    // that's what Excel does!
                val refp = PtgArea3d(false)
                refp.parentRec = ps.parentRec
                refp.useReferenceTracker = true    // very important!!! :)
                refp.location = locx
                return refp

            } else if (operands[0] is PtgRef) {
                // check if the ptgRef value is a string representing a Named range
                val o = operands[0].value
                val ps = PtgStr(o.toString())
                ps.parentRec = operands[0].parentRec
                operands = arrayOfNulls(1)
                operands[0] = ps
                return calcIndirect(operands)
            } else if (operands[0] is PtgName) {
                return calcIndirect(operands[0].components)
            }
        } catch (e: Exception) {
            //Logger.logErr("INDIRECT: " + e.toString());
        }

        return PtgErr(PtgErr.ERROR_REF)    // 's what Excel does ...
    }

    /**
     * LOOKUP
     * The LOOKUP function returns a value either from a one-row or one-column range
     * You can also use the LOOKUP function as an alternative to the IF function for elaborate tests or
     * tests that exceed the limit for nesting of functions. See the examples in the array form.
     * For the LOOKUP function to work correctly, the data being looked up must be sorted in
     * ascending order. If this is not possible, consider using the VLOOKUP, HLOOKUP, or MATCH functions.
     *
     *
     * A vector is a range of only one row or one column.
     * The vector form of LOOKUP looks in a one-row or one-column range (known as a vector) for a value and returns a value from the same position in a second one-row or one-column range. Use this form of the LOOKUP function when you want to specify the range that contains the values that you want to match. The other form of LOOKUP automatically looks in the first column or row.
     */
    fun calcLookup(operands: Array<Ptg>): Ptg {
        val lookup = operands[0].value.toString().toUpperCase()
        if (operands.size > 2) { //normal version of lookup
            val vector = operands[1].components
            val returnvector = operands[2].components ?: // happens when operands[2] is a PtgRef
            return PtgNumber(0.0)
// this is what excel does

            //If the LOOKUP function can't find the lookup_value, the function matches the largest value in lookup_vector that is less than or equal to lookup_value.
            //If lookup_value is smaller than the smallest value in lookup_vector, LOOKUP returns the #N/A error value
            var retval: Any? = null
            for (i in vector.indices) {
                if (Calculator.compareCellValue(vector[i].value, lookup, ">"))
                    break
                if (i < returnvector.size)
                    retval = returnvector[i].value
            }
            return if (retval is Number)
                PtgNumber(retval.toDouble())
            else if (retval is Boolean)
                PtgBool(retval.booleanValue())
            else if (retval == null)
                PtgErr(PtgErr.ERROR_NA)
            else
            // assume string
                PtgStr(retval.toString())
        } else { //array form of lookup
            /*
    	 *  The array form of LOOKUP looks in the first row or column of an array for the specified value
    	 *  and returns a value from the same position in the last row or column of the array.
    	 *  Use this form of LOOKUP when the values that you want to match are in the first row
    	 *  or column of the array. Use the other form of LOOKUP when you want to specify the location
    	 *  of the column or row.
			In general, it's best to use the HLOOKUP or VLOOKUP function instead of the array form of LOOKUP. This form of LOOKUP is provided for compatibility with other spreadsheet programs.

		With the HLOOKUP and VLOOKUP functions, you can index down or across, but LOOKUP always selects the last value in the row or column.

 */
            try {
                val array = operands[1].components
                val nrs = (operands[1] as PtgArray).numberOfRows
                var ncs = (operands[1] as PtgArray).numberOfColumns
                //If array covers an area that is wider than it is tall (more columns than rows), LOOKUP searches for the value of lookup_value in the first row.
                //If an array is square or is taller than it is wide (more rows than columns), LOOKUP searches in the first column.
                var retval: Any? = null
                var found = false
                val rowbased = ncs > nrs
                ncs++ // make 1-based
                var i = 0
                var j = 0
                while (j < nrs && !found) {
                    val start = i
                    while (i < start + ncs && !found) {
                        if (Calculator.compareCellValue(array[i].value, lookup, ">")) {
                            found = true
                            break
                        }
                        // returns a value from the same position in the last row or column of the array
                        if (rowbased) {
                            retval = array[i + ncs].value
                        } else {
                            retval = array[i + 1].value
                            i++
                        }
                        i++
                    }
                    j++
                }
                return if (retval is Number)
                    PtgNumber(retval.toDouble())
                else if (retval is Boolean)
                    PtgBool(retval.booleanValue())
                else if (retval == null)
                    PtgErr(PtgErr.ERROR_NA)
                else
                // assume string
                    PtgStr(retval.toString())
            } catch (e: Exception) {
                return PtgErr(PtgErr.ERROR_NA)
            }

        }
    }


    /**
     * MATCH
     * Looks up values in a reference or array
     * Returns the relative position of an item in an array that matches a specified value in a
     * specified order. Use MATCH instead of one of the LOOKUP functions when you need the position
     * of an item in a range instead of the item itself.
     *
     *
     * MATCH(lookup_value,lookup_array,match_type)
     *
     *
     * Lookup_value is the value you want to match in lookup_array. For example, when you look up someone's number
     * in a telephone book, you are using the person's name as the lookup value, but the telephone number is
     * the value you want.
     * Lookup_value can be a value (number, text, or logical value) or a cell reference to a number,
     * text, or logical value.
     *
     *
     * Lookup_array   is a contiguous range of cells containing possible lookup values. Lookup_array must
     * be an array or an array reference.
     *
     *
     * Match_type   is the number -1, 0, or 1. Match_type specifies how Microsoft Excel matches lookup_value
     * with values in lookup_array.
     *
     *
     * If match_type is 1, MATCH finds the largest value that is less than or equal to lookup_value.
     * Lookup_array must be placed in ascending order: ...-2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE.
     *
     *
     * If match_type is 0, MATCH finds the first value that is exactly equal to lookup_value.
     * Lookup_array can be in any order.
     *
     *
     * If match_type is -1, MATCH finds the smallest value that is greater than or equal to lookup_value.
     * Lookup_array must be placed in descending order: TRUE, FALSE, Z-A, ...2, 1, 0, -1, -2, ..., and so on.
     *
     *
     * If match_type is omitted, it is assumed to be 1.
     *
     *
     * MATCH returns the position of the matched value within lookup_array, not the value itself. For example, MATCH("b",{"a","b","c"},0) returns 2, the relative position of "b" within the array {"a","b","c"}.
     * MATCH does not distinguish between uppercase and lowercase letters when matching text values.
     * If MATCH is unsuccessful in finding a match, it returns the #N/A error value.
     * If match_type is 0 and lookup_value is text, lookup_value can contain the wildcard characters asterisk (*) and question mark (?). An asterisk matches any sequence of characters; a question mark matches any single character.
     */
    fun calcMatch(operands: Array<Ptg>): Ptg {
        try {
            val lookupValue = operands[0].value    // should be one value or a reference
            var lookupArray = operands[1]    // array or array reference (PtgArea)
            var values: Array<Ptg>? = null
            var matchType = 1
            if (operands.size > 2) {
                val o = operands[2].value
                if (o is Int)
                    matchType = o.toInt()
                else
                    matchType = (o as Double).toInt()

            }
            // Step 1- get all the components of the lookupArray (Array or Array Reference
            if (lookupArray is PtgName) {
                val pa = PtgArea3d(false)
                pa.parentRec = lookupArray.parentRec
                pa.location = lookupArray.name!!.location
                lookupArray = pa
            }
            if (lookupArray is PtgArea) {
                values = lookupArray.components
            } else if (lookupArray is PtgMemFunc) {
                values = lookupArray.components
            } else if (lookupArray is PtgArray) {
                values = lookupArray.components
            } else if (lookupArray is PtgMystery) {
                // PtgMystery is return from PtgMemFunc/MemArrays
                val ptgs = ArrayList()
                val p = lookupArray.vars
                for (j in p!!.indices) {
                    if (p[j] is PtgArea) {
                        val pa = p[j].components
                        for (k in pa.indices)
                            ptgs.add(pa[k])
                    } else
                        ptgs.add(p[j])
                }
                values = arrayOfNulls(ptgs.size)
                ptgs.toTypedArray()
            } else if (lookupArray is PtgStr) {
                val pa = PtgArea3d(false)
                pa.parentRec = lookupArray.parentRec
                pa.location = lookupArray.toString()
                values = pa.components
            } else { // testing!
                Logger.logErr("match: unknown type of lookup array")
            }

            // Step # 2- traverse thru value array to find lookupValue using matchType rules
            // ALSO must ensure for matchType!=0 that array is in ascending or descending order
            var retIndex = -1
            // TODO: matchType==0 Strings can match wildcards ...
            for (i in 1..values!!.size) {
                val v0 = values[i - 1].value
                var v1: Any? = null
                if (i < values.size)
                    v1 = values[i].value
                var mType = -2 // -1 means v0<v1, 0 means v0==v1, 1 means v0>v1
                var match = -2    // test lookupValue against v0 (i-1)
                if (v0 is Int) {
                    if (v1 != null) mType = v0.compareTo((v1 as Int?)!!)
                    match = v0.compareTo(lookupValue as Int)
                } else if (v0 is Double) {
                    if (v1 != null) mType = v0.compareTo((v1 as Double?)!!)
                    match = v0.compareTo(lookupValue as Double)
                } else if (v0 is Boolean) {
                    val bv0 = v0.booleanValue()
                    // 1.6 only if (v1!=null) mType= (((Boolean) v0).compareTo((Boolean)v1));
                    if (v1 != null) {
                        val bv1 = (v1 as Boolean).booleanValue()
                        mType = if (bv0 == bv1) 0 else if (!bv0 && bv1) -1 else +1
                    }
                    // 1.6 only match= (((Boolean) v0).compareTo((Boolean)lookupValue));
                    val bv1 = (lookupValue as Boolean).booleanValue()
                    match = if (bv0 == bv1) 0 else if (!bv0 && bv1) -1 else +1
                } else if (v0 is String) {
                    if (v1 != null) mType = v0.compareTo((v1 as String?)!!)
                    match = v0.compareTo(lookupValue as String)
                }
                if (i < values.size) { // only check order
                    if (matchType == 1 && mType > 0// not in ascending order
                            || matchType == -1 && mType < 0)
                    // not in descending order
                    // DOCUMENTATION SEZ MUST BE IN DESCENDING ORDER FOR -1 BUT EXCEL ALLOWS IT IN CERTAIN CIRCUMSTANCES
                    //						)
                        return PtgErr(PtgErr.ERROR_NA)
                }
                if (matchType == 0 && match == 0) {
                    retIndex = i    // 1-based
                    break
                }
                if (matchType == 1 && match <= 0) {
                    retIndex = i    // 1-based
                } else if (matchType == -1 && match >= 0) {
                    retIndex = i    // 1-based
                }
            }
            /*
			for (int i= 0; i < values.length; i++) {
				Object val=values[i].getValue();
				int mType= -2;
				if (val instanceof Integer)
					mType= ((Integer)val).compareTo((Integer)lookupValue);
				else if (val instanceof Double)
					mType= ((Double)val).compareTo((Double)lookupValue);
				else if (val instanceof Boolean)
					mType= ((Boolean)val).compareTo((Boolean)lookupValue);
				else if (val instanceof String)
					mType= ((String)val).toLowerCase().compareTo(((String)lookupValue).toLowerCase());
				if (matchType==0) {// matches if equal
					if (mType==0) {// then got it!
						retIndex= i;
						break;
					}
				} else if (matchType==-1) {
				} else  {	// default is 1
				}
			}
*/
            if (retIndex > -1)
                return PtgInt(retIndex)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_NA)
    }

    /**
     * calcOffset
     * Returns a reference to a range that is a specified number of rows and columns from a cell or range of cells.
     * The reference that is returned can be a single cell or a range of cells. You can specify the number of rows and the number of columns to be returned.
     *
     *
     * OFFSET(reference,rows,cols,height,width)
     *
     *
     * Reference   is the reference from which you want to base the offset. Reference must refer to a cell or range of adjacent cells; otherwise, OFFSET returns the #VALUE! error value.
     * Rows   is the number of rows, up or down, that you want the upper-left cell to refer to. Using 5 as the rows argument specifies that the upper-left cell in the reference is five rows below reference. Rows can be positive (which means below the starting reference) or negative (which means above the starting reference).
     * Cols   is the number of columns, to the left or right, that you want the upper-left cell of the result to refer to. Using 5 as the cols argument specifies that the upper-left cell in the reference is five columns to the right of reference. Cols can be positive (which means to the right of the starting reference) or negative (which means to the left of the starting reference).
     * Height   is the height, in number of rows, that you want the returned reference to be. Height must be a positive number.
     * Width   is the width, in number of columns, that you want the returned reference to be. Width must be a positive number.
     */
    internal fun calcOffset(operands: Array<Ptg>): Ptg {
        val ref = operands[0] as? PtgRef
                ?: return PtgErr(PtgErr.ERROR_VALUE)    // Reference must refer to a cell or range of adjacent cells; otherwise, OFFSET returns the #VALUE! error value
        val nrows = operands[1].intVal
        val ncols = operands[2].intVal
        var height = -1    // Height   if present, how many rows to return - must be positive
        if (operands.size > 3) {
            height = operands[3].intVal
            if (height < 0)
                return PtgErr(PtgErr.ERROR_VALUE)
        }
        var width = -1    // Width   	if present, how many columns to return - must be positive
        if (operands.size > 4) {
            width = operands[4].intVal
            if (width < 0)
                return PtgErr(PtgErr.ERROR_VALUE)
        }
        var rc = ref.intLocation
        rc[0] += nrows
        rc[1] += ncols
        if (rc!!.size > 3) { // it's an area
            rc[2] += nrows
            rc[3] += ncols
        }
        // A height/width of 1, 1= a single reference
        // When we increase either the row height or column width in the offset function "=OFFSET(A1,2,0,1,1)"
        // to more than 1, the reference is converted to a range.
        // OK, this may be the wrong interpretation of the height and width, but, from research, this is what I've come up with:
        // height and width are only really applicable for an initial single reference
        // in Excel, a height and width value of more than 1 returns #VALUE! unless wrapped in SUM
        // ????
        if (height == 1 && width == 1) { // it's a single reference
            if (rc.size > 3) {    // truncate
                val temp = IntArray(2)
                System.arraycopy(rc, 0, temp, 0, 2)
                rc = temp
            }
        } else if (!(height == -1 && width == -1)) {
            if (rc.size < 3) { // make a range
                val temp = IntArray(4)
                System.arraycopy(rc, 0, temp, 0, 2)
                rc = temp
            }
            if (height > 0) {
                rc[2] = rc[0] + height - 1    // is this correct????
            }
            if (width > 0) {
                rc[3] = rc[1] + width - 1        // " "
            }
        }
        if (rc.size > 3) {
            // If rows and cols offset reference over the edge of the worksheet, OFFSET returns the #REF! error value.
            if (rc[0] < 0 || rc[1] < 0 || rc[2] < 0 || rc[3] < 0)
                return PtgErr(PtgErr.ERROR_REF)
            val pa = PtgArea(false)
            pa.parentRec = ref.parentRec
            try {
                val sh = ref.location
                val z = sh.indexOf('!')
                if (z > 0) {
                    pa.sheetName = sh.substring(0, z)
                }
            } catch (e: Exception) {
            }

            pa.setLocation(rc)
            return pa
        } else { // it's a single reference
            // If rows and cols offset reference over the edge of the worksheet, OFFSET returns the #REF! error value.
            if (rc[0] < 0 || rc[1] < 0)
                return PtgErr(PtgErr.ERROR_REF)
            val pr = PtgRef()
            pr.parentRec = ref.parentRec
            pr.setLocation(rc)
            return pr
        }
    }

    /**
     * TRANSPOSE
     * Returns the transpose of an array
     * The TRANSPOSE function returns a vertical range of cells as a horizontal range, or vice versa.
     * The TRANSPOSE function must be entered as an array formula has columns and rows.
     * Use TRANSPOSE to shift the vertical and horizontal orientation of an array or range on a
     * worksheet.
     *
     *
     * array  Required. An array or range of cells on a worksheet that you want to transpose. The transpose of an array is created by using the first row of the array as the first column of the new array, the second row of the array as the second column of the new array, and so on.
     */
    internal fun calcTranspose(operands: Array<Ptg>): Ptg {
        var retArray = ""
        val ret = PtgArray()
        if (operands[0] !is PtgArray) {
            val arr = operands[0].components
            //it's a list of values, convert to row-based
            for (i in arr.indices) {
                retArray = retArray + arr[i].value.toString() + ";"
            }
            retArray = "{" + retArray.substring(0, retArray.length - 1) + "}"
            ret.setVal(retArray)
        } else {    // transpose row/cols of an existing array
            val pa = operands[0] as PtgArray
            val arr = pa.components
            val nc = pa.numberOfColumns + 1
            val nr = pa.numberOfRows + 1
            for (i in 0 until nc) {
                var j = 0
                while (j < nc * nr) {
                    retArray = retArray + arr!![i + j].value.toString() + ","
                    j += nc
                }
                retArray = retArray.substring(0, retArray.length - 1) + ";"
            }
            retArray = "{" + retArray.substring(0, retArray.length - 1) + "}"
            ret.setVal(retArray)
        }
        return ret
    }

    /**
     * ROW
     * Returns the row number of a reference
     *
     *
     * Note this is 1 based, ie Row1 = 1.
     */
    @Throws(FunctionNotSupportedException::class)
    internal fun calcRow(operands: Array<Ptg>?): Ptg {
        if (operands == null || operands.size != 1) return PtgErr(PtgErr.ERROR_VALUE)
        if (operands[0] is PtgFuncVar) {
            // we need to return the col where the formula is.
            val pfunk = operands[0] as PtgFuncVar
            try {
                val loc = pfunk.parentRec!!.rowNumber + 1
                return PtgInt(loc)
            } catch (e: Exception) {
                Logger.logErr("Error running calcRow $e")
            }

        }
        try {        // process as an array formula ...
            val isArray = operands[0].parentRec is Array
            if (!isArray) {
                if (operands[0] is PtgRef) {
                    return PtgInt((operands[0] as PtgRef).rowCol[0] + 1)
                } else if (operands[0] is PtgName) {    // table???
                    val range = (operands[0] as PtgName).name!!.location
                    return PtgInt(ExcelTools.getRowColFromString(range!!)[0] + 1)
                }
                return PtgInt(operands[0].intLocation[0] + 1)
            } else {
                var retArry = ""
                var comps: Array<Ptg>? = null
                if (operands[0] is PtgRef) {
                    comps = operands[0].components
                } else if (operands[0] is PtgName) {    // table???
                    comps = operands[0].components
                }
                if (comps == null) {
                    return PtgInt((operands[0] as PtgRef).rowCol[0] + 1)
                }
                for (i in comps.indices) {
                    try {
                        retArry = retArry + ((comps[i] as PtgRef).intLocation!![0] + 1) + ","
                    } catch (e: Exception) {
                    }

                }
                retArry = "{" + retArry.substring(0, retArry.length - 1) + "}"
                val pa = PtgArray()
                pa.setVal(retArry)
                return pa
            }
        } catch (ex: Exception) {
            return PtgRefErr()
        }

    }


    /**
     * ROWS
     * Returns the number of rows in a reference
     *
     *
     * ROWS(array)
     *
     *
     * Array   is an array, an array formula, or a reference to a range of cells for which you want the number of rows.
     *
     *
     * =ROWS(C1:E4) Number of rows in the reference (4)
     * =ROWS({1,2,3;4,5,6}) Number of rows in the array constant (2)
     */
    @Throws(FunctionNotSupportedException::class)
    internal fun calcRows(operands: Array<Ptg>): Ptg {
        try {
            var rsz = 0
            if (operands[0] is PtgStr) {
                val rangestr = operands[0].value.toString()
                val startx = rangestr.substring(0, rangestr.indexOf(":"))
                val endx = rangestr.substring(rangestr.indexOf(":") + 1)

                val startints = ExcelTools.getRowColFromString(startx)
                val endints = ExcelTools.getRowColFromString(endx)
                rsz = endints[0] - startints[0]
                rsz++ // inclusive
            } else if (operands[0] is PtgName) {
                val rc = ExcelTools.getRangeCoords(operands[0].location)
                rsz = rc[2] - rc[0]
                rsz++ // inclusive
            } else if (operands[0] is PtgRef) {
                val rc = ExcelTools.getRangeCoords((operands[0] as PtgRef).location)
                rsz = rc[2] - rc[0]
                rsz++ // inclusive
            } else if (operands[0] is PtgMemFunc) {
                val p = operands[0].components
                if (p != null && p.size > 0) {
                    val rc0 = p[0].intLocation
                    var rc1: IntArray? = null
                    if (p.size > 1)
                        rc1 = p[p.size - 1].intLocation
                    if (rc1 == null)
                        rsz = 0
                    else
                        rsz = rc1[0] - rc0[0]
                    rsz++
                } else
                    return PtgErr(PtgErr.ERROR_VALUE)
            } else
                return PtgErr(PtgErr.ERROR_VALUE)
            return PtgInt(rsz)
        } catch (e: Exception) {
        }

        return PtgErr(PtgErr.ERROR_VALUE)
    }


    /**
     * VLOOKUP
     * Looks in the first column of an array and moves across the row to return the value of a cell
     * Searches for a value in the leftmost column of a table, and then returns a value in the same row
     * from a column you specify in the table. Use VLOOKUP instead of HLOOKUP when your comparison
     * values are located in a column to the left of the data you want to find.
     *
     *
     * Syntax
     *
     *
     * VLOOKUP(lookup_value,table_array,col_index_num,range_lookup)
     *
     *
     * Lookup_value   is the value to be found in the first column of the array. Lookup_value can be a value,
     * a reference, or a text string.
     *
     *
     * Table_array   is the table of information in which data is looked up. Use a reference to a range or
     * a range name, such as Database or List.
     *
     *
     * If range_lookup is TRUE, the values in the first column of table_array must be placed in ascending order:
     * ..., -2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE; otherwise VLOOKUP may not give the correct value.
     * If range_lookup is FALSE, table_array does not need to be sorted.
     * You can put the values in ascending order by choosing the Sort command from the Data menu and selecting Ascending.
     * The values in the first column of table_array can be text, numbers, or logical values.
     * Uppercase and lowercase text are equivalent.
     * Col_index_num   is the column number in table_array from which the matching value must be returned.
     * A col_index_num of 1 returns the value in the first column in table_array; a col_index_num of 2 returns the value
     * in the second column in table_array, and so on. If col_index_num is less than 1,
     * VLOOKUP returns the #VALUE! error value; if col_index_num is greater than the number of columns in table_array,
     * VLOOKUP returns the #REF! error value.
     *
     *
     * Range_lookup   is a logical value that specifies whether you want VLOOKUP to find an exact match or an
     * approximate match. If TRUE or omitted, an approximate match is returned. In other words,
     * if an exact match is not found, the next largest value that is less than lookup_value is returned.
     * If FALSE, VLOOKUP will find an exact match. If one is not found, the error value #N/A is returned.
     *
     *
     * Remarks
     *
     *
     * If VLOOKUP can't find lookup_value, and range_lookup is TRUE, it uses the largest value that is less than or
     * equal to lookup_value.
     * If lookup_value is smaller than the smallest value in the first column of table_array, VLOOKUP returns the
     * #N/A error value.
     * If VLOOKUP can't find lookup_value, and range_lookup is FALSE, VLOOKUP returns the #N/A value.
     *
     *
     * VLOOKUP(lookup_value,table_array,col_index_num,range_lookup)
     *
     *
     * On the preceding worksheet, where the range A4:C12 is named Range:
     *
     *
     * VLOOKUP(1,Range,1,TRUE) equals 0.946
     *
     *
     * VLOOKUP(1,Range,2) equals 2.17
     *
     *
     * VLOOKUP(1,Range,3,TRUE) equals 100
     *
     *
     * VLOOKUP(.746,Range,3,FALSE) equals 200
     *
     *
     * VLOOKUP(0.1,Range,2,TRUE) equals #N/A, because 0.1 is less than the smallest value in column A
     *
     *
     * VLOOKUP(2,Range,2,TRUE) equals 1.71
     */
    @Throws(FunctionNotSupportedException::class)
    internal fun calcVlookup(operands: Array<Ptg>): Ptg {

        var rangeLookup = true    // truth of "approximate match"; must be sorted
        var isNumber = true
        try {
            val lookup_value = operands[0]
            var table_array = operands[1]

            // can't assume that it's a PtgInt
            //PtgInt col_index_num 	= (PtgInt)operands[2].getValue();
            //int colNum = col_index_num.getVal() -1;// reduce 1 for ordinal base off firstcol
            val o = operands[2].value
            var colNum = 0
            if (o is Double)
                colNum = o.toInt() - 1 // reduce 1 for ordinal base off firstcol
            else
            // assume int?
                colNum = (o as Int).toInt() - 1 // reduce 1 for ordinal base off firstcol
            if (operands.size > 3) {
                val vx = operands[3].value
                if (vx != null) {
                    try {
                        val sort = vx as Boolean
                        rangeLookup = sort.booleanValue()
                    } catch (e: ClassCastException) {
                        val bool = vx as Int
                        if (bool == 0) rangeLookup = false
                    }

                }
            }

            var lookupComponents: Array<PtgRef>? = null
            var valueComponents: Array<PtgRef>? = null
            // first, get the lookup Column Vals
            if (table_array is PtgName) {    // 20090211 KSC:
                val pa = PtgArea3d(false)
                pa.parentRec = table_array.parentRec

                pa.location = table_array.name!!.location
                table_array = pa
                if (table_array.isExternalRef) {
                    Logger.logWarn("LookupReferenceCalculator.calcVlookup External References are disallowed")
                    return PtgErr(PtgErr.ERROR_REF)
                }
            }
            if (table_array is PtgArea) {
                try {
                    val pa = table_array
                    val range = table_array.intLocation
                    //			 TODO: check rc sanity here
                    val firstcol = range[1]
                    lookupComponents = pa.getColComponents(firstcol) as Array<PtgRef>
                    valueComponents = pa.getColComponents(firstcol + colNum) as Array<PtgRef>
                } catch (/*20070209 KSC: FormulaNotFound*/e: Exception) {
                    Logger.logWarn("LookupReferenceCalculator.calcVlookup cannot determine PtgArea location. $e")
                }

            } else if (table_array is PtgMemFunc) { //  || table_array instanceof PtgMemArea){
                try {
                    val pa = table_array
                    // int[] range = table_array.getIntLocation();
                    var firstcol = -1

                    try {
                        val rc1 = pa.firstloc!!.intLocation

                        //   				 TODO: check rc sanity here
                        firstcol = rc1[1]
                    } catch (e: Exception) {
                        Logger.logWarn("LookupReferenceCalculator.calcVlookup could not determine row col from PtgMemFunc.")
                    }

                    lookupComponents = pa.getColComponents(firstcol) as Array<PtgRef>
                    valueComponents = pa.getColComponents(firstcol + colNum) as Array<PtgRef>
                } catch (/*20070209 KSC: FormulaNotFound*/e: Exception) {
                    Logger.logWarn("LookupReferenceCalculator.calcVlookup cannot determine PtgArea location. $e")
                }

            }
            // error check
            if (lookupComponents == null || lookupComponents.size == 0)
                return PtgErr(PtgErr.ERROR_REF)
            if (lookup_value == null || lookup_value.value == null)
            // 20070221 KSC: Error trap getValue
                return PtgErr(PtgErr.ERROR_NULL)
            // lets check if we are dealing with strings or numbers....
            try {
                val `val` = lookup_value.value.toString()
                if (`val`.length == 0)
                // 20090205 KSC
                    return PtgErr(PtgErr.ERROR_NA)
                val d = Double(`val`)
            } catch (e: NumberFormatException) {
                isNumber = false
            }

            // TODO:
            //if the value you supply for the lookup_value argument is smaller than the smallest value in the first column of the table_array argument, VLOOKUP returns the #N/A error value.

            if (isNumber) {
                val match_num: Double
                try {
                    match_num = java.lang.Double.parseDouble(
                            lookup_value.value.toString())
                } catch (e: NumberFormatException) {
                    return PtgErr(PtgErr.ERROR_NA)
                }

                for (i in lookupComponents.indices) {
                    val `val`: Double

                    try {
                        `val` = java.lang.Double.parseDouble(
                                lookupComponents[i].value!!.toString())
                        if (`val` == 0.0) {    // VLOOKUP does NOT treat blanks as 0's
                            if (lookupComponents[i].refCell!![0] == null) {
                                continue
                            }
                        }
                    } catch (e: NumberFormatException) {
                        // Ignore entries in the table that aren't numbers.
                        continue
                    }

                    if (`val` == match_num) {
                        return valueComponents!![i].ptgVal
                    } else if (rangeLookup && `val` > match_num) {
                        return if (i == 0)
                            PtgErr(PtgErr.ERROR_NA)
                        else
                            valueComponents!![i - 1].ptgVal
                    }
                }

                return if (rangeLookup)
                    valueComponents!![lookupComponents.size - 1].ptgVal
                else
                    PtgErr(PtgErr.ERROR_NA)
            } else {
                if (rangeLookup) {    // approximate match
                    val match_str = lookup_value.value.toString()
                    val match_len = match_str.length
                    for (i in lookupComponents.indices) {
                        try {
                            val `val` = lookupComponents[i].value!!.toString()
                            if (`val`.equals(match_str, ignoreCase = true)) {// we found it
                                return valueComponents!![i].ptgVal
                            } else if (`val`.length >= match_len && `val`.substring(0, match_len).equals(match_str, ignoreCase = true)) { // matches up to length, but not all, return previous
                                return valueComponents!![i - 1].ptgVal
                            } else if (ExcelTools.getIntVal(`val`.substring(0, 1)) > ExcelTools.getIntVal(match_str.substring(0, 1))) {
                                return valueComponents!![i - 1].ptgVal
                            } else if (i == lookupComponents.size - 1) {// we reached the last one so use this
                                return valueComponents!![i].ptgVal
                            }
                        } catch (e: Exception) {
                        }
                        // 20070209 KSC: ignore errors in lookup cells
                    }
                } else { // unsorted
                    val match_str = lookup_value.value.toString()
                    for (i in lookupComponents.indices) {
                        try {
                            val `val` = lookupComponents[i].value!!.toString()
                            try {
                                if (`val`.equals(match_str, ignoreCase = true)) {// we found it
                                    return valueComponents!![i].ptgVal
                                } else if (i == lookupComponents.size - 1) {// we reached the last one so error out
                                    return PtgErr(PtgErr.ERROR_NA)
                                }
                            } catch (e: Exception) {
                                Logger.logErr("LookupReferenceCalculator.calcVLookup error: $e")
                                return PtgErr(PtgErr.ERROR_NA)
                            }

                        } catch (e: Exception) {
                        }
                        // 20070209 KSC: ignore errors in lookup cells
                    }
                }
            }// It's a String
        } catch (e: Exception) {    // appears that an error with operands results in a #NA error
            return PtgErr(PtgErr.ERROR_NA)

        }

        return PtgErr(PtgErr.ERROR_NULL)
    }

}
