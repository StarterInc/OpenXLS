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

import io.starter.formats.XLS.ReferenceTracker
import io.starter.toolkit.Logger

import java.util.*


/**
 * DatabaseCalculator is a collection of static methods that operate
 * as the Microsoft Excel function calls do.
 *
 *
 * All methods are called with an array of ptg's, which are then
 * acted upon.  A Ptg of the type that makes sense (ie boolean, number)
 * is returned after each method.
 *
 *
 *
 *
 * Database and List Management functions
 * Microsoft Excel includes worksheet functions that analyze data stored in lists or databases.
 * Each of these functions, referred to collectively as the Dfunctions, uses three arguments: database, field, and criteria.
 *
 *
 * These arguments refer to the worksheet ranges that are used by the function.
 *
 *
 * DAVERAGE   Returns the average of selected database entries
 *
 *
 * DCOUNT   Counts the cells that contain numbers in a database
 *
 *
 * DCOUNTA   Counts nonblank cells in a database
 *
 *
 * DGET   Extracts from a database a single record that matches the specified criteria
 *
 *
 * DMAX   Returns the maximum value from selected database entries
 *
 *
 * DMIN   Returns the minimum value from selected database entries
 *
 *
 * DPRODUCT   Multiplies the values in a particular field of records that match the criteria in a database
 *
 *
 * DSTDEV   Estimates the standard deviation based on a sample of selected database entries
 *
 *
 * DSTDEVP   Calculates the standard deviation based on the entire population of selected database entries
 *
 *
 * DSUM   Adds the numbers in the field column of records in the database that match the criteria
 *
 *
 * DVAR   Estimates variance based on a sample from selected database entries
 *
 *
 * DVARP   Calculates variance based on the entire population of selected database entries
 *
 *
 * GETPIVOTDATA   Returns data stored in a PivotTable
 *
 *
 *
 *
 * ABOUT DB
 *
 *
 * All Database Formulas take 3 arguments:
 * Database is the range of cells that makes up the list or database.
 *
 *
 * A database is a list of related data in which rows of related information
 * are records, and columns of data are fields.
 *
 *
 * The first row of the list contains labels for each column.
 * Field   indicates which column is used in the function.
 * Field can be given as text with the column label
 * enclosed between double quotation marks, such as "Age" or "Yield,"
 * or as a number that represents the position of the column within the
 * list: 1 for the first column, 2 for the second column, and so on.
 *
 *
 * Criteria   is the range of cells that contains the conditions you specify.
 *
 *
 * You can use any range for the criteria argument,  * 	as long as it
 * includes at least one column label and at least one cell below the column
 * label for specifying a condition for the column.
 *
 *
 * Make sure the criteria range does not overlap the list.
 *
 *
 *
 *
 *
 *
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
object DatabaseCalculator {
    var DEBUG = false

    /**
     * Fetch a DB from the cache or create a new one
     *
     *
     *
     *
     * Dbs store Cell refs...
     *
     * @param operands
     * @return
     */
    private fun getDb(operands: Ptg): DB? {
        val DBcache = operands.parentRec.workBook!!.refTracker

        //gonna try never caching this... painful, but if we dont'
        if (DBcache!!.listDBs[operands.toString()] != null) {
            //Logger.logInfo("getDB: " + operands.toString()+ "using cache.");
            return DBcache.listDBs[operands.toString()] as DB
        }
        //}
        // create new
        //Logger.logErr("getDB: " + operands.toString()+ "NOT cached.");
        val dbrange = PtgCalculator.getAllComponents(operands)
        val ret = DB.parseList(dbrange)
        DBcache.listDBs[operands.toString()] = ret
        return ret
    }

    private fun getCriteria(operands: Ptg): Criteria? {
        val DBcache = operands.parentRec.workBook!!.refTracker

        // test without cache
        if (DBcache!!.criteriaDBs[operands.toString()] != null) {
            //Logger.logInfo("getCriteria: " + operands.toString()+ "using cache.");
            return DBcache.criteriaDBs[operands.toString()] as Criteria
        }
        //Logger.logErr("getCriteria: " + operands.toString()+ "NOT cached.");
        val criteria = PtgCalculator.getAllComponents(operands)
        val ret = Criteria.parseCriteria(criteria)
        DBcache.criteriaDBs[operands.toString()] = ret
        return ret
    }

    /**
     * DAVERAGE   Returns the average of selected database entries
     */
    internal fun calcDAverage(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NA)
        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        var average = 0.0
        var count = 0

        // this is the colname to match
        val colname = operands[1].value.toString()
        for (i in db.rows.indices) {    // loop thru all db rows
            // check if current row passes criteria requirements
            val rwz = db.getRow(i) // slight optimization one less call to getRow -jm
            if (crit.passes(colname, rwz, db)) {
                // passes; now do action
                val vx = rwz!![fNum]
                if (vx != null) {
                    try {
                        average += java.lang.Double.parseDouble(vx.value.toString())
                        count++    // if it can be parsed into a number, increment count
                    } catch (exp: NumberFormatException) {
                    }

                }
            }
        }
        if (count > 0) average = average / count
        return PtgNumber(average)
    }

    /**
     * DCOUNT   Counts the cells that contain numbers in a database
     */
    internal fun calcDCount(operands: Array<Ptg>): Ptg {
        if (operands.size != 3)
            return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)

        var count = 0
        val nrow = db.nRows
        // this is the colname to match
        val colname = operands[1].value.toString()
        for (i in 0 until nrow) {    // loop thru all db rows
            // check if current row passes criteria requirements
            try {
                val rr = db.getRow(i)

                /* "passes" means that there is a matching
                 * cell in the row of the data db cells
                 *
                 */
                if (crit.passes(colname, rr, db)) {
                    // passes; now do action
                    val cx = db.getCell(i, fNum)
                    val vtx = cx!!.value.toString()
                    if (vtx != null) {
                        java.lang.Double.parseDouble(vtx)
                        count++    // if it can be parsed into a number, increment count
                    }
                }
            } catch (e: NumberFormatException) {
            }

        }
        return PtgNumber(count.toDouble())
    }

    /**
     * DCOUNTA   Counts nonblank cells in a database
     */
    internal fun calcDCountA(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        var count = 0

        // this is the colname to match
        val colname = operands[1].value.toString()
        for (i in 0 until db.nRows) {    // loop thru all db rows
            // check if current row passes criteria requirements
            if (crit.passes(colname, db.getRow(i), db)) {
                // passes; now do action
                val s = db.getCell(i, fNum)!!.value.toString()
                if (s != null && s.trim { it <= ' ' } != "")
                    count++    // if field is not blank, increment count
            }
        }

        return PtgNumber(count.toDouble())
    }

    /**
     * DGET   Extracts from a database a single record that matches the specified criteria
     */
    internal fun calcDGet(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NULL)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        var `val` = ""
        var count = 0

        // this is the colname to match
        val colname = operands[1].value.toString()
        for (i in 0 until db.nRows) {    // loop thru all db rows
            // check if current row passes criteria requirements
            if (crit.passes(colname, db.getRow(i), db)) {
                // passes; now do action
                `val` = db.getCell(i, fNum)!!.value.toString()
                count++
            }
        }
        if (count == 0) return PtgErr(PtgErr.ERROR_VALUE)    // no recs match
        return if (count > 1) PtgErr(PtgErr.ERROR_NUM) else PtgStr(`val`) // if more than one record matches criteria
    }

    /**
     * DMAX   Returns the maximum value from selected database entries
     */
    internal fun calcDMax(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        var max = java.lang.Double.MIN_VALUE

        // this is the colname to match
        val colname = operands[1].value.toString()
        for (i in 0 until db.nRows) {    // loop thru all db rows
            // check if current row passes criteria requirements
            if (crit.passes(colname, db.getRow(i), db)) {
                // passes; now do action
                val vtx = db.getCell(i, fNum)!!.value.toString()
                if (vtx != null) {
                    try {
                        if (vtx.length > 0) max = Math.max(max, java.lang.Double.parseDouble(vtx))
                    } catch (exp: NumberFormatException) {
                    }

                }
            }
        }

        if (max == java.lang.Double.MIN_VALUE) max = 0.0
        return PtgNumber(max)
    }

    /**
     * DMIN   Returns the minimum value from selected database entries
     */
    internal fun calcDMin(operands: Array<Ptg>): Ptg {
        if (operands.size != 3)
        // sanity checks
            return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        var min = java.lang.Double.MAX_VALUE
        // this is the colname to match
        val colnamx = operands[1].value.toString()

        for (i in 0 until db.nRows) {    // loop thru all db rows
            // check if current row passes criteria requirements
            try {
                val rwz = db.getRow(i)

                if (crit.passes(colnamx, rwz, db)) {
                    // passes; now do action
                    try {
                        val dbx = db.getCell(i, fNum)

                        if (dbx != null) {
                            val dnb = dbx.value.toString()
                            if (dnb != null) {
                                if (dnb.length > 0)
                                    min = Math.min(min, java.lang.Double.parseDouble(dnb))
                            }
                        }
                    } catch (ex: Exception) {
                        // normal blanks etc.
                    }

                }
            } catch (e: NumberFormatException) {
            }

        }
        if (min == java.lang.Double.MAX_VALUE) min = 0.0
        return PtgNumber(min)
    }

    /**
     * DPRODUCT   Multiplies the values in a particular field of records that match the criteria in a database
     */
    internal fun calcDProduct(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        var product = 1.0
        // this is the colname to match
        val colname = operands[1].value.toString()

        for (i in 0 until db.nRows) {    // loop thru all db rows
            // check if current row passes criteria requirements
            try {
                if (crit.passes(colname, db.getRow(i), db)) {
                    // passes; now do action
                    val fnx = db.getCell(i, fNum)!!.value.toString()
                    if (fnx != null)
                        product *= java.lang.Double.parseDouble(fnx)
                }
            } catch (e: NumberFormatException) {
            }

        }
        return PtgNumber(product)
    }

    /**
     * DSTDEV   Estimates the standard deviation based on a sample of selected database entries
     */
    internal fun calcDStdDev(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        val vals = java.util.ArrayList()
        var sum = 0.0
        var count = 0
        // this is the colname to match
        val colname = operands[1].value.toString()

        for (i in 0 until db.nRows) {    // loop thru all db rows
            // check if current row passes criteria requirements
            if (crit.passes(colname, db.getRow(i), db)) {
                // passes; now do action
                try {
                    val fnx = db.getCell(i, fNum)!!.value.toString()
                    if (fnx != null) {
                        val x = java.lang.Double.parseDouble(fnx)
                        sum += x
                        count++
                        vals.add(java.lang.Double.toString(x))
                    }
                } catch (e: NumberFormatException) {
                }

            }
        }
        var stdev = 0.0
        if (count > 0) {
            val average = sum / count
            // now have all values in vals
            for (i in 0 until count) {
                val x = java.lang.Double.parseDouble(vals.get(i) as String)
                stdev += Math.pow(x - average, 2.0)
            }
            if (count > 1) count--
            stdev = Math.sqrt(stdev / count)
        }
        return PtgNumber(stdev)
    }

    /**
     * DSTDEVP   Calculates the standard deviation based on the entire population of selected database entries
     */
    internal fun calcDStdDevP(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        val vals = java.util.ArrayList()
        var sum = 0.0
        var count = 0
        // this is the colname to match
        val colname = operands[1].value.toString()

        for (i in 0 until db.nRows) {    // loop thru all db rows
            // check if current row passes criteria requirements
            if (crit.passes(colname, db.getRow(i), db)) {
                // passes; now do action
                try {
                    val fnx = db.getCell(i, fNum)!!.value.toString()
                    if (fnx != null) {
                        val x = java.lang.Double.parseDouble(fnx)
                        sum += x
                        count++
                        vals.add(java.lang.Double.toString(x))
                    }
                } catch (e: NumberFormatException) {
                }

            }
        }
        var stdevp = 0.0
        if (count > 0) {
            val average = sum / count
            // now have all values in vals
            for (i in 0 until count) {
                val x = java.lang.Double.parseDouble(vals.get(i) as String)
                stdevp += Math.pow(x - average, 2.0)
            }
            stdevp = Math.sqrt(stdevp / count)
        }
        return PtgNumber(stdevp)
    }

    /**
     * DSUM   Adds the numbers in the field column of records in the database that match the criteria
     */
    internal fun calcDSum(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)

        val count = 0
        var sum = 0.0
        val nrow = db.nRows
        // this is the colname to match
        val colname = operands[1].value.toString()
        for (i in 0 until nrow) {    // loop thru all db rows
            // check if current row passes criteria requirements
            try {
                val rr = db.getRow(i)

                /* "passes" means that there is a matching
                 * cell in the row of the data db cells
                 *
                 */
                if (crit.passes(colname, rr, db)) {
                    // passes; now do action
                    try {
                        val fnx = db.getCell(i, fNum)!!.value.toString()
                        if (fnx != null)
                            sum += java.lang.Double.parseDouble(fnx)
                    } catch (e: NumberFormatException) {
                    }

                }
            } catch (e: NumberFormatException) {
            }

        }
        return PtgNumber(sum)
    }


    /**
     * DVAR   Estimates variance based on a sample from selected database entries
     */
    internal fun calcDVar(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        val vals = java.util.ArrayList()
        var sum = 0.0
        var count = 0

        // this is the colname to match
        val colname = operands[1].value.toString()
        for (i in 0 until db.nRows) {    // loop thru all db rows
            // check if current row passes criteria requirements
            if (crit.passes(colname, db.getRow(i), db)) {
                // passes; now do action
                try {
                    val fnx = db.getCell(i, fNum)!!.value.toString()
                    if (fnx != null) {
                        val x = java.lang.Double.parseDouble(db.getCell(i, fNum)!!.toString())
                        sum += x
                        count++
                        vals.add(java.lang.Double.toString(x))
                    }
                } catch (e: NumberFormatException) {
                }

            }
        }
        var variance = 0.0
        if (count > 0) {
            val average = sum / count
            // now have all values in vals
            for (i in 0 until count) {
                val x = java.lang.Double.parseDouble(vals.get(i) as String)
                variance += Math.pow(x - average, 2.0)
            }
            if (count > 1) count--
            variance = variance / count
        }
        return PtgNumber(variance)
    }

    /* DVARP   Calculates variance based on the entire population of selected database entries
     */
    internal fun calcDVarP(operands: Array<Ptg>): Ptg {
        if (operands.size != 3) return PtgErr(PtgErr.ERROR_NA)

        val db = getDb(operands[0])
        val crit = getCriteria(operands[2])
        if (db == null || crit == null)
            return PtgErr(PtgErr.ERROR_NUM)
        val fNum = db.findCol(operands[1].string.trim { it <= ' ' })
        if (fNum == -1)
            return PtgErr(PtgErr.ERROR_NUM)
        val vals = java.util.ArrayList()
        var sum = 0.0
        var count = 0

        // this is the colname to match
        val colname = operands[1].value.toString()
        for (i in 0 until db.nRows) {    // loop thru all db rows
            // check if current row passes criteria requirements
            if (crit.passes(colname, db.getRow(i), db)) {
                // passes; now do action
                try {
                    val fnx = db.getCell(i, fNum)!!.value.toString()
                    if (fnx != null) {
                        val x = java.lang.Double.parseDouble(fnx)
                        sum += x
                        count++
                        vals.add(java.lang.Double.toString(x))
                    }
                } catch (e: NumberFormatException) {
                }

            }
        }
        var varP = 0.0
        if (count > 0) {
            val average = sum / count
            // now have all values in vals
            for (i in 0 until count) {
                val x = java.lang.Double.parseDouble(vals.get(i) as String)
                varP += Math.pow(x - average, 2.0)
            }
            varP = varP / count
        }
        return PtgNumber(varP)
    }
    /* GETPIVOTDATA   Returns data stored in a PivotTable
     *
     *
     */
}


/**
 * EXPLANATION of Database Formulas
 *
 *
 * Database   is the range of cells that makes up the list or database.
 * A database is a list of related data in which rows of related
 * information are records, and columns of data are fields.
 * The first row of the list contains labels for each column.
 *
 *
 * Field   indicates which column is used in the function.
 * Field can be given as text with the column label enclosed between
 * double quotation marks, such as "Age" or "Yield," or as a number
 * that represents the position of the column within the list:
 *
 *
 * 1 for the first column,
 * 2 for the second column, and so on.
 *
 *
 * NOTE: quotes around field text is optional:
 * first row of columns and rows are field labels.
 *
 *
 * Criteria   is the range of cells that contains the conditions
 * you specify. You can use any range for the criteria argument,
 * as long as it includes at least one column label and at least one
 * cell below the column label for specifying a condition for the column.
 *
 *
 * Example
 */
internal open class DB(nCols: Int, nRows: Int) {
    var colHeaders: Array<String>
    var rows: Array<Array<Ptg>>

    val nCols: Int
        get() = colHeaders.size

    val nRows: Int
        get() = rows.size

    init {
        colHeaders = arrayOfNulls(nCols)
        // TODO: replace with PtgRef array!
        rows = Array(nRows) { arrayOfNulls(nCols) }
    }

    /**
     * return the index of the col in the DB
     *
     * @param cname
     * @return
     */
    fun getCol(cname: String): Int {
        for (t in colHeaders.indices) {
            if (colHeaders[t].equals(cname, ignoreCase = true)) {
                return t
            }
        }
        return -1
    }

    /**
     * return a row of DB ptgs
     *
     * @param i
     * @return
     */
    fun getRow(i: Int): Array<Ptg>? {
        return if (i > -1 && i < rows.size) rows[i] else null
    }

    /**
     * return a col of ??
     *
     * @param i
     * @return
     */
    fun getCol(i: Int): String? {
        return if (i > -1 && i < colHeaders.size) colHeaders[i] else null
    }

    fun getCell(row: Int, col: Int): Ptg? {
        try {
            return rows[row][col]
        } catch (e: Exception) {
            return null
        }

    }

    fun findCol(f: String?): Int {
        for (i in colHeaders.indices) {
            if (colHeaders[i].trim { it <= ' ' }.equals(f!!, ignoreCase = true)) {
                return i
            }
        }
        try {
            val j = Integer.parseInt(f!!)    // one-based index into columns
            return j - 1
        } catch (e: Exception) {
            return -1
        }

    }

    companion object {

        /**
         * Write some documentation here please... thanks! -jm
         *
         * @param dbrange
         * @return
         */
        fun parseList(dbrange: Array<Ptg>): DB? {
            var prevCol = -1
            var nCols = 0
            var nRows = 0
            var maxRows = 0

            // allocate the empty table for the dbrange
            for (i in dbrange.indices) {
                if (dbrange[i] is PtgRef) {
                    val pref = dbrange[i] as PtgRef
                    val loc = pref.intLocation
                    //				 TODO: check rc sanity here
                    if (loc!![1] != prevCol) { // count # cols
                        prevCol = loc[1]
                        nCols++
                        nRows = 0
                    } else { // count # rows
                        nRows++
                        maxRows = Math.max(nRows, maxRows)
                    }
                } else
                    return null
            }

            // now populate the table
            val dblist = DB(nCols, maxRows)
            prevCol = -1
            nCols = -1
            nRows = -1
            for (i in dbrange.indices) {
                val db1 = dbrange[i] as PtgRef
                val loc = db1.intLocation
                var vs: Any? = null    // 20081120 KSC: Must distinguish between blanks and 0's
                if (!db1.isBlank)
                    vs = db1.value

                // column headers
                if (loc!![1] != prevCol) {
                    if (vs != null)
                        dblist.colHeaders[++nCols] = vs.toString()
                    prevCol = loc[1]
                    nRows = 0
                } else { // get value Ptgs
                    try {
                        dblist.rows[nRows++][nCols] = db1
                    } catch (e: ArrayIndexOutOfBoundsException) {
                        // possible nCols==-1
                    }

                }
            }
            return dblist
        }
    }
}

internal class Criteria(nCols: Int, nRows: Int) : DB(nCols, nRows) {

    private var criteriaCache: MutableMap<*, *>? = null

    // TODO: Handle formula criteria!
    // TODO: To perform an operation on an entire column in a database, enter a blank line below the column labels in the criteria range
    // TODO: Handle various EQUALS:  currency, number ...
    fun matches(v: String?, cx: Any): Boolean {
        var bMatches = false
        var c: String? = ""

        if (cx is Ptg)
            c = cx.value.toString()
        else
            c = cx.toString()

        if (c == null || c.length == 0)
            return false
        if (v == null)
            return false    // 20070208 KSC: null means no match!

        // TODO: handle this using calc methods

        // relational
        if (c.substring(0, 1) == ">") {
            try {
                if (c.length > 1 && c.substring(0, 2) == ">=") {
                    c = c.substring(2)
                    bMatches = java.lang.Double.parseDouble(v) >= java.lang.Double.parseDouble(c)
                } else {
                    c = c.substring(1)
                    bMatches = java.lang.Double.parseDouble(v) > java.lang.Double.parseDouble(c)
                }
            } catch (e: NumberFormatException) {
            }

        } else if (c.substring(0, 1) == "<") {
            try {
                if (c.length > 1 && c.substring(0, 2) == "<=") {
                    c = c.substring(2)
                    bMatches = java.lang.Double.parseDouble(v) <= java.lang.Double.parseDouble(c)
                } else {
                    c = c.substring(1)
                    bMatches = java.lang.Double.parseDouble(v) < java.lang.Double.parseDouble(c)
                }
            } catch (e: NumberFormatException) {
            }

        } else {// Equals
            bMatches = v.equals(c, ignoreCase = true)
        }
        return bMatches
    }

    /**
     * "passes" means that there is a matching
     * and valid (criteria passing) cell in the
     * row of the data db cells
     *
     * @param field  - the field or column that is being compared
     * @param curRow - check the current row/col
     * @param db     - the db of vals
     * @return
     */
    fun passes(field: String, curRow: Array<Ptg>?, db: DB): Boolean {
        val nrows = nRows
        val ncols = nCols

        var crit_format = -1 // determine format of criteria table
        // multiple rows of criteria for one column are OR searches
        // on that column (col=x OR col=y OR col=z)
        if (nrows > 1 && ncols == 1)
            crit_format = 0

        // more than one column are AND searches (cola=x AND colb=y)
        // ONE row of criteria is applied to N rows of data
        if (nrows == 1 && ncols > 1)
            crit_format = 1

        // more than one row and more than one column:
        // OR's:  (cola=x AND colb=y) || (cola=z AND colb=w)
        // applied to each row of data
        if (nrows > 1 && ncols > 1)
            crit_format = 2

        if (nrows == 1 && ncols == 1)
            crit_format = 0

        when (crit_format) {

            0 -> return criteriaCheck1(curRow, db)

            1 -> return criteriaCheck2(field, curRow!!, db)

            2 -> return criteriaCheck3(field, curRow, db)

            else -> return criteriaCheck3(field, curRow, db)
        }
    }

    /**
     * for the current db row/column, see if all criteria for that column matches
     * (criteria may run over many criteria rows & signifies an OR search)
     * find db column that matches criteria col
     *
     *
     * multiple rows of data one row of criteria
     *
     *
     * age	height
     * <29	>99
     *
     *
     * age	height
     * 23	88
     * 21	99
     * 43	56
     * 23	56
     * 44	76
     *
     * @param curRow
     * @param db
     * @return
     */
    private fun criteriaCheck1(curRow: Array<Ptg>?, db: DB): Boolean {
        var bColOK = false
        var bRowOK = true
        val nrows = nRows
        val ncols = nCols
        val ndbrows = db.nRows

        val cl = getCol(0)
        val j = db.findCol(cl)
        if (j >= 0) {
            bColOK = false    // need one bColOK= true for it to pass
            var k = 0
            while (k < nrows && !bColOK) {    //
                try {
                    val v = curRow!![j].value.toString()
                    val r = rows[k][0]
                    val rv = r.value.toString()
                    bColOK = matches(v, r)

                    // fast succeed
                    if (bColOK)
                        return true

                } catch (ex: Exception) {
                    // Logger.logInfo("DBCalc"); // TODO: check that this is OK
                }

                k++
            }
            if (!bColOK)
                bRowOK = false    // if no criteria passes, row doesn't pass
        }
        return bRowOK
    }


    /**
     * There is only one row of criteria,
     * but may be multiple criteria per field aka:
     *
     *
     * type	age	age	 height
     * blue >21	<50	 44
     *
     *
     *
     *
     * // [v1	v2	v3]
     * // [v2	v3	v4]
     * // ...
     * // [crit1	crit2	crit3] <- val must pass this
     *
     * @param field
     * @param curRow
     * @param db
     * @return
     */
    private fun criteriaCheck2(field: String,
                               curRow: Array<Ptg>,
                               db: DB): Boolean {
        var pass = true
        for (t in curRow.indices) {
            val valcheck = curRow[t].value.toString()
            // for each value check all the criteria
            for (x in rows.indices) {
                val r = this.getCriteria(db.colHeaders[t])
                val tx = r.iterator()
                while (tx.hasNext()) {
                    val cv = tx.next() as Ptg
                    val vc = cv.value.toString()
                    if (vc != "") {
                        pass = matches(valcheck, vc)
                        // fast fail is OK
                        if (!pass)
                            return false
                    }
                }
            }
        }
        return pass
    }


    /**
     * othertimes we have a 2D criteria
     *
     *
     * type	age	age	 height
     * blue >21	>50	 44
     * blue <30 <100
     * red  >30 >40  99
     *
     *
     *
     *
     * AND criteria across criteria rows
     * OR  criteria down cols
     *
     * @param field
     * @param curRow
     * @param db
     * @return
     */
    private fun criteriaCheck3(field: String,
                               curRow: Array<Ptg>?,
                               db: DB): Boolean {
        var critRowMatch = false
        // for each value check all the criteria in a row
        // multiple rows of criteria are combined
        for (x in rows.indices) {
            critRowMatch = true // reset
            // for each row of criteria, iterate criteria cols
            for (rs in this.colHeaders.indices) {
                val critField = this.colHeaders[rs]

                val r = this.getCriteria(x, critField)
                val tx = r.iterator()
                val dv = db.getCol(critField)
                val valcheck = curRow!![dv].value.toString()
                // this criteria row may pass/fail subsequent rows may pass/fail
                // only one has to pass OK to return true all crit in this row must pass
                while (tx.hasNext() && critRowMatch) { // stop if row failure
                    val cv = tx.next() as Ptg
                    val vc = cv.value.toString()

                    if (vc != "") {
                        critRowMatch = matches(valcheck, vc)
                        /* fast fail is not OK because we
                         * may pass one row OR another
                         * AND criteria across criteria rows
                         * OR  criteria down cols
                         */
                    }
                }
            }
            if (critRowMatch)
            // fast succeed here, a row passed
                return true
        }
        return critRowMatch
    }


    /**
     * equal number of data rows/cols and criteria rows/cols
     *
     *
     * compare each criteria cell with each corresponding db cell
     *
     * @param field
     * @param curRow
     * @param db
     * @return
     */
    private fun criteriaCheck4(field: String, curRow: Array<Ptg>, db: DB): Boolean {
        var bColOK = false
        val bRowOK = false
        val nrows = nRows
        val ncols = nCols

        var k = 0
        while (k < nrows && !bRowOK) {    // for each row of criteria
            // for each col check for valid criteria
            for (i in 0 until ncols) {
                // find db column that matches criteria column
                // String coli = getCol(i); // get the col

                val coln = db.findCol(field) // which db col is this?

                if (coln >= 0) {        // matching col in dblist
                    // the current criteria matches on colname
                    var curcrit: Ptg? = null
                    if (rows[k].size > ncols) { // matched criteria cols and dbcols
                        curcrit = rows[k][coln]
                    } else {
                        Logger.logWarn("DatabaseCalculator.Criteria.criteriaCheck4: wrong criteria count for db value count")
                        return false
                    }
                    val rt = curcrit.value.toString()

                    // the current val is in the matching column
                    val rx = curRow[coln]
                    val mv = rx.value.toString()

                    // check field if this is a matching lookup
                    if (i == 0) {
                        if (mv != rt)
                            return false
                    } else { // check criteria, criteria is fast fail
                        if (rt == "") { // empty criteria is not false

                        } else {    // for the current db row/column, see if all criteria for that column matches
                            //(criteria may run over many criteria rows & signifies an OR search)
                            bColOK = matches(rt, rx)

                            if (!bColOK && nrows == 1)
                            // fast fail
                                return false


                            if (bColOK)
                            // fast succeed
                                return true
                        }
                    }
                } else { // see if it's a formula - defined to NOT match db column
                    // Logger.logWarn("DataBaseCalculator.Criteria.passes: no matching col in dblist.");
                }
            }
            if (bColOK)
                return true // does this cut the fat? bRowOK= true;	// if any row's column criteria passes, row passes
            k++
        }
        return bRowOK
    }

    /**
     * return cached array of criteria Ptgs for a given field
     *
     *
     * TODO: cache reset
     *
     * @param field
     * @return
     */
    private fun getCriteria(field: String): List<*> {
        if (true)
        //criteriaCache==null)
            criteriaCache = Hashtable() // map of vector criteria
        else if (criteriaCache!!.get(field) != null)
            return criteriaCache!!.get(field) as List<*>

        val crits = Vector()
        for (x in rows.indices) {
            for (t in colHeaders.indices) {
                if (colHeaders[t] == field) {
                    crits.add(rows[x][t])
                }
            }
        }
        criteriaCache!![field] = crits
        return crits
    }

    /**
     * return cached array of criteria Ptgs for a given field
     *
     *
     * TODO: cache reset
     *
     * @param field
     * @return
     */
    private fun getCriteria(critRow: Int, field: String): List<*> {
        val crits = Vector()
        if (critRow != -1) { // option to return criteria per row
            for (t in colHeaders.indices) {
                if (colHeaders[t] == field) {
                    crits.add(rows[critRow][t])
                }
            }
        }
        return crits
    }

    companion object {

        fun parseCriteria(criteria: Array<Ptg>): Criteria? {
            val dblist = DB.parseList(criteria) ?: return null
            val crit = Criteria(dblist.nCols, dblist.nRows)
            crit.colHeaders = dblist.colHeaders
            crit.rows = dblist.rows
            if (DatabaseCalculator.DEBUG) {
                Logger.logInfo("\nCriteria:")
                for (i in 0 until crit.nCols) {
                    Logger.logInfo("\t" + crit.getCol(i)!!)
                }
                for (j in 0 until crit.nCols) {
                    for (i in 0 until crit.nRows) {
                        Logger.logInfo("\t" + crit.getCell(i, j)!!)
                    }
                }
            }
            return crit
        }
    }
}