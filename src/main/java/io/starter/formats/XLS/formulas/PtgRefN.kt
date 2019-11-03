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


/**
 * PtgRefN is a modified PtgRef that is for shared formulas.
 * Put here by M$ to make us miserable,
 *
 *
 * it would have made much more sense to just use a PtgRef.
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rw          2       The row of the reference (so says the docs, but it is the row I think
 * 2           grbitCol    2       (see following table)
 *
 * Only the low-order 14 bits specify the Col, the other bits specify
 * relative vs absolute for both the col or the row.
 *
 * Bits        Mask        Name    Contents
 * -----------------------------------------------------
 * 15          8000h       fRwRel  =1 if row offset relative,
 * =0 if otherwise
 * 14          4000h       fColRel =1 if row offset relative,
 * =0 if otherwise
 * 13-0        3FFFh       col     Ordinal column offset or number
</pre> *
 *
 *
 *
 *
 * This token contains the relative reference to a cell in the same sheet.
 * It stores relative components as signed offsets and is used in shared formulas, conditional formatting, and data validity.
 *
 * @see WorkBook
 *
 * @see Boundsheet
 *
 * @see Dbcell
 *
 * @see Row
 *
 * @see Cell
 *
 * @see XLSRecord
 */
class PtgRefN(useReference: Boolean) : PtgRef() {
    /* private int formulaRow;
    private int formulaCol;*/
    private var realRow: Int = 0
    private var realCol: Int = 0
    internal var column: Short = 0
    private var parea: PtgArea? = null

    override val isReference: Boolean
        get() = true

    /**
     * Returns the location of the Ptg as a string (ie c4)
     *
     *
     * TODO: look into this possible bug
     *
     *
     * There is a problem here as the location will always be relative and incorrect.
     * this is deprecated and should be calling convertToPtgRef
     */
    /* Set the location of this PtgRef.  This takes a location
       such as "a14"

       TODO: check why this is overridden / reversed 12/02 -jm
    */
    override//if (!populated){throw new FormulaNotFoundException("Cannot set location, no Formula Present");}
    // the row is a relative location
    // the column is a relative location
    // - colNew;
    // 20080215 KSC: replace address stripping
    //stripped of sheet name, if any ...
    // 20060301 KSC: Keep relativity
    // handle row shifting issues
    // 20090325 KSC: trap OOXML external reference link, if any
    var location: String?
        get() {
            realRow = rw
            realCol = col
            if (fRwRel) {
                realRow += formulaRow.toShort().toInt()
            }
            if (fColRel) {
                realCol = formulaCol.toShort().toInt()
                if (realCol >= XLSConstants.MAXCOLS) realCol -= XLSConstants.MAXCOLS
            }
            val s = ExcelTools.getAlphaVal(realCol)
            val y = (realRow + 1).toString()

            return (if (fColRel) "" else "$") + s + (if (fRwRel) "" else "$") + y
        }
        set(address) {
            var address = address
            if (record != null) {
                val s = ExcelTools.stripSheetNameFromRange(address)
                address = s[1]

                val res = ExcelTools.getRowColFromString(address)
                if (fRwRel) {
                    rw += formulaRow - res[0]
                    if (rw < 0)
                        rw = 0
                    formulaRow = res[0]
                } else {
                    rw = res[0]
                }
                if (fColRel) {
                    col += formulaCol - res[1]
                    formulaCol = res[1]
                } else {
                    col = res[1]
                }

                updateRecord()
                init(record)
                if (s[3] != null)
                    externalLink1 = Integer.valueOf(s[3].substring(1, s[3].length - 1)).toInt()
                if (s[4] != null)
                    externalLink2 = Integer.valueOf(s[4].substring(1, s[4].length - 1)).toInt()

            } else {
                Logger.logWarn("PtgRefN.setLocation() failed: NO record data: $address")
            }
        }

    /**
     * returns the row/col ints for the ref
     * adjusted for the host cell
     *
     * @return
     */
    override// the row is a relative location
    // the column is a relative location
    // - colNew;
    val rowCol: IntArray
        get() {
            realRow = rw
            realCol = col
            if (fRwRel) {
                realRow += formulaRow.toShort().toInt()
            }
            if (fColRel) {
                realCol = formulaCol.toShort().toInt()
                if (realCol >= XLSConstants.MAXCOLS) realCol -= XLSConstants.MAXCOLS
            }
            return intArrayOf(realRow, realCol)
        }

    /**
     * this
     *
     * @return
     */
    val realRowCol: IntArray
        get() = intArrayOf(rw, col)

/**
 * /*
 * (try to) return int[] array containing the row/column
 * referenced by this PtgRefN.
 *
 * @returns int[] row/col absolute (non-offset) location
 * @see io.starter.formats.XLS.formulas.PtgRef.getIntLocation
*/
public override// the row is a relative location
// the column is a relative location
// 20070205 KSC: Added	20080102 KSC: added =
val intLocation:IntArray?
get() {

var rowNew = rw
var colNew = col
if (fRwRel)
{
rowNew += formulaRow
}
if (fColRel)
{
colNew += formulaCol
}
if (colNew >= XLSConstants.MAXCOLS) colNew -= XLSConstants.MAXCOLS

val returning = IntArray(2)
returning[0] = rowNew
returning[1] = colNew
return returning
}


/**
 * custom RefTracker usage:  uses entire range covered by all shared formulas
*/
// TODO: determine if this is an OK maxcol (Excel 2007)
// TODO: determine if this is an OK maxcol (Excel 2007)
val area:PtgArea
get() {
val sh = this.parentRec as Shrfmla?
val i = IntArray(4)
if (fRwRel)
{
i[0] = sh!!.firstRow + rw
}
else
{
i[0] = rw
}
if (fColRel)
{
i[1] = sh!!.firstCol + col
}
else
{
i[1] = col
}
if (fRwRel)
{
i[2] = sh!!.lastRow + rw
}
else
{
i[2] = rw
}
if (fColRel)
{
i[3] = sh!!.lastCol + col
}
else
{
i[3] = col
}

if ((i[1] >= XLSConstants.MAXCOLS_BIFF8 && !this.parentRec!!.workBook!!.isExcel2007))
i[1] -= XLSConstants.MAXCOLS_BIFF8
if ((i[3] >= XLSConstants.MAXCOLS_BIFF8 && !this.parentRec!!.workBook!!.isExcel2007))
i[3] -= XLSConstants.MAXCOLS_BIFF8


val parea = PtgArea(i, sh, true)
return parea
}

public override fun init(b:ByteArray) {
opcode = b[0]
record = b
populateVals()
hashcode = hashCode        // different from PtgRef calc
}

init{
this.useReferenceTracker = useReference
}

/* Set the location of this PtgRef.  This takes a location
such as {1,2}
*/
public override fun setLocation(rowcol:IntArray) {
if (useReferenceTracker) this.removeFromRefTracker()
if (record != null)
{    // 20090217 KSC: had some errors here, redid
if (fRwRel)
formulaRow = rowcol[0]
else
rw = rowcol[0]
if (fColRel)
{
formulaCol = rowcol[1]
}
else
{
col = rowcol[1]
}
this.updateRecord()
init(record)
}
else
{
Logger.logWarn("PtgRefN.setLocation() failed: NO record data: " + rowcol.toString())
}
hashcode = hashCode
if (useReferenceTracker)
this.addToRefTracker()

}

/**
 * Convert this PtgRefN to a PtgRef based on the offsets included in the PtgExp &
 * if this uses relative or absolute offsets
 *
 * @param pxp
 * @return
*/
fun convertToPtgRef(r:XLSRecord/*PtgExp pxp*/):PtgRef {
//XLSRecord r = (XLSRecord)pxp.getParentRec();
val i = IntArray(2)
if (fRwRel)
{
i[0] = r.rowNumber + rw
}
else
{
i[0] = rw
}
if (fColRel)
{
i[1] = r.colNumber + col
}
else
{
i[1] = col
}

if ((i[1] >= XLSConstants.MAXCOLS_BIFF8 && !r.workBook!!.isExcel2007))
// TODO: determine if this is an OK maxcol (Excel 2007)
i[1] -= XLSConstants.MAXCOLS_BIFF8

val prf = PtgRef(i, r, false)
//	  	String s = ExcelTools.formatLocation(i, fRwRel, fColRel);
//	  	PtgRef prf = new PtgRef(s, r /*pxp.getParentRec()*/, false); //false);

return prf
}

/**
 * set formula row
 *
 * @param r int new row
*/
// 20060301 KSC: access to formula row/col
fun setFormulaRow(r:Int) {
formulaRow = r
}

/**
 * set formula col
 *
 * @param c int new col
*/
// 20060301 KSC: access to formula row/col
fun setFormulaCol(c:Int) {
formulaCol = c
}

/**
 * add "true" area to reference tracker i.e. entire range referenced by all shared formula members
*/
public override fun addToRefTracker() {
val iParent = this.parentRec!!.opcode.toInt()
if (iParent == XLSConstants.SHRFMLA.toInt())
{
// KSC: TESTING - local ptgarea gets finalized and messes up ref. tracker on multiple usages without close
//				getArea();
//				parea.addToRefTracker();
val parea = area // otherwise is finalized if local var --- but take out ptgarea finalize for now
parea.addToRefTracker()
}
}

/**
 * remove "true" area from reference tracker i.e. entire range referenced by all shared formula members
*/
public override fun removeFromRefTracker() {
val iParent = this.parentRec!!.opcode.toInt()
if (iParent == XLSConstants.SHRFMLA.toInt())
{
val parea = area // otherwise is finalized if local var --- but take out ptgarea finalize for now
if (parea != null)
{
parea.removeFromRefTracker()
//			  		parea.close();
}
//	parea= null;
}
}

public override fun close() {
//removeFromRefTracker();
if (parea != null)
parea!!.close()
parea = null
}

companion object {

/**
 * serialVersionUID
*/
private val serialVersionUID = 2652944516984815274L
}
}