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

import io.starter.OpenXLS.*
import io.starter.formats.OOXML.*
import io.starter.formats.XLS.charts.Chart
import io.starter.formats.XLS.charts.Fontx
import io.starter.formats.XLS.charts.GenericChartObject
import io.starter.formats.XLS.formulas.FormulaParser
import io.starter.formats.XLS.formulas.Ptg
import io.starter.formats.XLS.formulas.PtgRef
import io.starter.formats.cellformat.CellFormatFactory
import io.starter.toolkit.*
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserException
import org.xmlpull.v1.XmlPullParserFactory

import java.io.*
import java.util.*
import java.util.zip.ZipEntry
import java.util.zip.ZipFile

/**
 * **Boundsheet: WorkSheet Information 0x85**<br></br>
 *
 *
 * This record stores the sheet name, type and stream position.
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       lbPlyPos    4       Stream position of the BOF for the sheet
 * 8       grbit       2       Option flags
 * 10      cch         1       Length of sheet name
 * 11      grbitChr    1       Compressed/Uncompressed Unicode
 * 12      rgch        var     Sheet name
</pre> *
 *
 *
 * ----  File Layout ----
 *
 *
 * BOUNDSHEET
 * Bof
 * Index
 * Row1 is first index in DBCELL
 * Row2 is the offset of any DBCELLS
 * Row...
 * CELLREC
 * CELLREC
 * CELLREC
 *
 *
 * DBCELL rgdb rg rg rg rg
 * ROWS
 * CELLS
 * DBCELL rg rg rg
 * ...
 * EOF
 *
 *
 * ----------------------
 *
 *
 * lbplypos used to be the most important thing on earth.  now it is not an issue.
 *
 * @see WorkBook
 *
 * @see Index
 *
 * @see Dbcell
 *
 * @see Row
 *
 * @see Cell
 *
 * @see XLSRecord
 */
 class Boundsheet:XLSRecord(), Sheet {
override var myBof:Bof? = null
private set
override var myEof:Eof? = null
private set
private var sheetname = ""
private var sheetHash = ""
private val rows = LinkedHashMap<Int, Row>()

private var cellsByRow:SortedMap<CellAddressible, BiffRec> = TreeMap(
CellAddressible.RowMajorComparator())

private var cellsByCol:SortedMap<CellAddressible, BiffRec> = TreeMap(
CellAddressible.ColumnMajorComparator())

private val arrFormulaLocs = HashMap()                        // use
 // for
    // trapping
    // array
    // formula
    // refs
    // to
    // original
    // cell
    // reference
    // [OOXML
    // Array
    // Formulas]
    protected var arrayformulas:AbstractList<*> = ArrayList()                        // trap
 // array
    // formulas
    // that
    // span
    // one
    // or
    // more
    // cells
    private val colinfos = TreeMap<ColumnRange, Colinfo>(
ColumnRange.Comparator())
private var SheetRecs:AbstractList<*> = ArrayList()
private var localrecs:AbstractList<*>? = null

/**
 * Records containting various bits of print setup.
 */
    private var printRecs:MutableList<*>? = null

 // These records are for boundsheet transferral to a new
    // book.
    private var transferXfs:MutableList<*> = ArrayList()
private var transferFonts:MutableList<*>? = ArrayList()
 var imageMap:HashMap<*, *>? = HashMap()
private val charts = ArrayList()                        // chart
 // specific
    // for
    // this
    // sheet
    private var lbPlyPos:Long = 0
/**
 * get the type of sheet as a short
 */
    override var sheetType:Short = 0
private set
private var cch:Byte = 0
/**
 * @return Returns the grbitChr.
 */
    /**
 * @param grbitChr The grbitChr to set.
 */
    override var grbitChr:Byte = 0
 // private int sheetnum;
    private var myidx:Index? = null
/**
 * get the last BiffRec added to this sheet
 */
    override var lastCell:BiffRec? = null
private set
private var lastRow:Row? = null
override var window2:Window2? = null
private var scl:Scl? = null
private var pane:Pane? = null
 var dvalRec:Dval? = null

 // KSC:
    // track
    // last-used
    // Object
    // id
    // for
    // this
    // sheet

    override var header:Headerrec? = null
protected set
override var footer:Footerrec? = null
protected set
override var wsBool:WsBool? = null
override var guts:Guts? = null

 var lastObjId = 0                                    // 20100210

/**
 * return whether to shift (formula cells, named ranges) "inclusive" or not
 *
 * @return
 */
     var isShiftInclusive = false
private set
private var cond_formats:AbstractList<*>? = Vector()
private val autoFilters = Vector()                // 20100111
 // KSC

    // OOXML use: stores external sheet-level OOXML objects
    private var ooxmlObjects:AbstractList<*> = ArrayList()

 // OOXML-specific sheet attributes TODO: translate to Excel
    // 2003 version IF POSSIBLE
    private var thickBottom = false
private var thickTop = false
private var zeroHeight = false
private var customHeight = false
/**
 * return the default row height in points (Excel 2007-Specific)
 */
    /**
 * set the default row height in points (Excel 2007-Specific)
 */
     var defaultRowHeight = 12.75                    // measured
 // in
    // point
    // size
    private var defaultColWidth = (-1.0).toFloat()
 var defColWidth:DefColWidth? = null
/**
 * return map of Excel 2007 shapes in this workbook
 *
 * @return
 */
     var ooxmlShapes:HashMap<*, *>? = null
private set                        // stores
 // OOXML
    // shapes
    /**
 * returns the Excel 2007 sheetView element for this sheet (controls topLeftCell, pane attributes ...
 *
 * @return
 */
    /**
 * set the Excel 2007 sheetView element for this sheet (controls topLeftCell, pane attributes ...
 *
 * @param s
 */
     var sheetView:SheetView? = null
/**
 * returns the Excel 2007 sheetPr sheet Properties element for this sheet (controls codename, tabColor ...)
 *
 * @return
 */
    /**
 * set the Excel 2007 sheetView element for this sheet (controls topLeftCell, pane attributes ...
 *
 * @param s
 */
     var sheetPr:SheetPr? = null
/**
 * returns the Excel 2007 autoFilter element for this sheet (temporarily hides rows based upon filter criteria)
 * TODO: Merge with 2003 AutoFilter
 *
 * @return
 */
    /**
 * set the Excel 2007 autoFilter element for this sheet (temporarily hides rows based upon filter criteria)
 * TODO: Merge with 2003 AutoFilter
 *
 * @param strref
 */
     var ooAutoFilter:io.starter.formats.OOXML.AutoFilter? = null
private var protector:SheetProtectionManager? = null

@Transient private var sheetNameRecs:HashMap<*, *>? = HashMap()            // sheet
 // scoped
    // names
    internal var mc:MutableList<*> = CompatibleVector()
/**
 * @return Returns the lastselection.
 */
    /**
 * @param lastselection The lastselection to set.
 */
     var lastselection:Selection? = null

/**
 * Gets this sheet's SheetProtectionManager.
 */
     val protectionManager:SheetProtectionManager
get() {
if (protector == null)
protector = SheetProtectionManager(this)
return protector
}

 /*
     * TODO: find calls to this method which really need to be
     * calling 'assembleSheetRecs() -jm 8/05
     */
    override val sheetRecs:List<*>
get() =SheetRecs

/**
 * returns the images list
 */
     val imageVect:List<*>
get() {
val im = ArrayList()
val ir = imageMap!!.keys.iterator()
while (ir.hasNext())
{
im.add(ir.next())
}
return im

}

/**
 * Get a collection of all names in the worksheet
 */
     val sheetScopedNames:Array<Name>
get() {
if (this.sheetNameRecs == null)
this.sheetNameRecs = HashMap()
val a = ArrayList(this.sheetNameRecs!!.values)
val n = arrayOfNulls<Name>(a.size)
a.toTypedArray()
return n
}

/**
 * for whatever reason, we return a Handle from an internal class
 *
 * @return
 */
    /*
         * 20071026 KSC: since there may be multiple copies of the
         * same
         * image in the sheet, must build imageHandle array by hand
         */ val images:Array<ImageHandle>?
get() {
if (imageMap == null)
return null
val im = arrayOfNulls<ImageHandle>(imageMap!!.size)
val ir = imageMap!!.keys.iterator()
var i = 0
while (ir.hasNext())
{
im[i++] = ir.next() as ImageHandle
}
return im
}

 val indexOfMsodrawingselection:Int
get() {
var rec:BiffRec? = null

val size = SheetRecs.size
var foundIndex = -1
for (i in 0 until size)
{
rec = SheetRecs[i] as BiffRec
if (rec is MSODrawingSelection)
{
foundIndex = i
break
}
}
return foundIndex
}

 val indexOfWindow2:Int
get() {
var rec:BiffRec? = null
val size = SheetRecs.size
var foundIndex = -1
for (i in 0 until size)
{
rec = SheetRecs[i] as BiffRec
if (rec is Window2)
{
foundIndex = i
break
}
}
return foundIndex
}

 val indexOfDimensions:Int
get() {
var rec:BiffRec? = null
val size = SheetRecs.size
var foundIndex = -1
for (i in 0 until size)
{
rec = SheetRecs[i] as BiffRec
if (rec is Dimensions)
{
foundIndex = i + 1
break
}
}
return foundIndex
}

/**
 * Return an array of all the dvRecs within
 * this boundsheet (Dval parent rec)
 *
 * @return
 */
     val dvRecs:List<*>?
get() =if (this.dvalRec != null) this.dvalRec!!.dvs else null

/**
 * @return conditional formats for this sheet
 */
     val conditionalFormats:List<*>?
get() =cond_formats

override var workBook:WorkBook?
get() =wkbook
set

/**
 * Determine if the boundsheet is a chart only boundsheet
 *
 * @return
 */
    override val isChartOnlySheet:Boolean
get() =if (myBof != null) myBof!!.isChartBof else false

/**
 * get the min/max dimensions
 * for this sheet.
 */
    private var dimensions:Dimensions? = null

override val minRow:Int
get() =dimensions!!.getRowFirst()

/**
 * return true maximum/last row on the sheet
 */
    override val maxRow:Int
get() =dimensions!!.getRowLast()

override val minCol:Int
get() =dimensions!!.getColFirst()

override val maxCol:Int
get() =dimensions!!.getColLast()

/**
 * set the numeric sheet number
 */
    override val sheetNum:Int
get() =this.wkbook!!.sheetVect.indexOf(this)

/**
 * get whether this sheet is hidden upon opening (either regular or "very hidden"
 */
    override val hidden:Boolean
get() =sheetType != VISIBLE.toShort()

 val veryHidden:Boolean
get() =sheetType == VERY_HIDDEN.toShort()

internal var selected = false

/**
 * get the number of defined rows on this sheet
 */
    override val numRows:Int
get() =rows.size

/**
 * get the number of defined cells on this sheet
 */
    override val numCells:Int
get() {
var counter = 0
val cellset = rows.keys
val rws = cellset.toTypedArray()
if (rws.size == 0)
return 0
for (i in rws.indices)
{
val r = rows[rws[i]]
counter += r.numberOfCells
}
return counter

}

/**
 * get the FastAddVector of columns defined on this sheet
 */
    override val colNames:List<*>
get() {
val retvec = FastAddVector()
for (x in 0 until this.realMaxCol)
{
val c = ExcelTools.getAlphaVal(x)
retvec.add(c)
}
return retvec
}

/**
 * get the Number of columns defined on this sheet
 */
    override val numCols:Int
get() =realMaxCol

/**
 * get the FastAddVector of rows defined on this sheet
 */
    override val rowNums:List<*>
get() {
val e = rows.keys
val iter = e.iterator()
val rownames = FastAddVector()
while (iter.hasNext())
{
rownames.add(rownames.size, iter.next())
}
return rownames
}

/**
 * return the map of row in this sheet sorted by row #
 * (will be unsorted if insertions and deletions)
 *
 * @return
 */
     val sortedRows:SortedMap<*, *>
get() =TreeMap(rows)

/**
 * return a Map of the Rows
 */
     val rowMap:Map<*, *>
get() =rows

 var fastCellAdds = false // performance setting which skips

private var copypriorformats = true

/**
 * Returns the *real* last col num.  Unfortunately the dimensions record
 * cannot be counted on to give a correct value.
 */
     var realMaxCol = -1
private set
private var maximumCellRow = -1

/**
 * get an array of all cells for this worksheet
 */
    override val cells:Array<BiffRec>
get() {
val cells = cellsByRow.values
return cells.toTypedArray<BiffRec>()
}

/**
 * Get the built in names referring to this boundsheet
 *
 * @return
 */
    private val builtInNames:ArrayList<*>
get() {
val retlist = ArrayList()
val ns = this.workBook!!.names
for (i in ns.indices)
{
if (ns[i].isBuiltIn && (ns[i].getIxals().toInt() == this.sheetNum + 1 || ns[i].getItab().toInt() == this.sheetNum + 1))
{
retlist.add(ns[i])
}
}
return retlist
}

/**
 * Get the print area name rec for this
 * boundsheet, return null if not exists
 *
 * @return
 */
    protected val printAreaNameRec:Name?
get() =getPrintAreaNameRec(Name.PRINT_AREA)

/**
 * Get the print area set for this WorkSheetHandle.
 *
 *
 * If no print area is set return null;
 */
    /**
 * Set the print area for this worksheet.
 */
    /*
                 * if (p instanceof PtgArea3d) {// can be other than ptgarea
                 * ...
                 * ((PtgRef)p).clearLocationCache();// why??
                 * return p.toString();
                 * }
                 */ var printArea:String?
get() {
val n = this.printAreaNameRec
if (n != null)
{
var ret = ""
val s = n.expression
for (x in s!!.indices)
{
val p = s[x] as Ptg
ret += p.toString()
}
return ret
}
return null
}
set(range) =setPrintArea(range, Name.PRINT_AREA)

/**
 * Get the Print Titles range set for this WorkSheetHandle.
 *
 *
 * If no Print Titles are set, this returns null;
 */
    /**
 * Set the print titles for this worksheet= row(s) or col(s) to repeat at the top of each page
 */
     var printTitles:String?
get() {
val n = getPrintAreaNameRec(Name.PRINT_TITLES)
if (n != null)
{
val s = n.expression
for (x in s!!.indices)
{
val p = s[x] as Ptg
return p.toString()
}
}

return null
}
set(range) =setPrintArea(range, Name.PRINT_TITLES)

/**
 * returns an arrayList of notes in the worksheet
 *
 * @return
 */
     val notes:ArrayList<*>
get() {
val notes = ArrayList()
var idx = this.getIndexOf(XLSConstants.NOTE)
while (idx > -1)
{
notes.add(SheetRecs[idx++])
if ((SheetRecs[idx] as BiffRec).opcode != XLSConstants.NOTE)
break
}
return notes
}

/***
 */
    override// 20081031 KSC- don't automatically add new!
 val mergedCellsRec:Mergedcells?
get() =if (mc.size == 0) null else this.mergedCellsRecs[this.mergedCellsRecs.size - 1] as Mergedcells

override/*
         * 20081031 don't add a merged cell rec automatically
         * if (mc.size()>0) {
         * return mc;
         * }
         * Mergedcells mec =
         * (Mergedcells)Mergedcells.getPrototype();
         * mec.setSheet(this);
         * this.getStreamer().addRecordAt(mec,
         * this.getSheetRecs().size()-1);
         * this.addMergedCellsRec(mec);
         */ val mergedCellsRecs:List<*>
get() =mc

/**
 * return existing merged cell records without adding new blank
 *
 * @return
 */
     val mergedCells:List<*>
get() =mc

/**
 * get the name of the sheet
 */
    /**
 * change the displayed name of the sheet
 *
 *
 * Affects the following byte values:
 * 10      cch         1       Length of sheet name
 * 11      grbitChr    1       Compressed/Uncompressed Unicode
 * 12      rgch        var     Sheet name
 */
    override// if (!ByteTools.isUnicode(namebytes)){
 var sheetName:String
get() =sheetname
set(newname) {

cch = newname.length.toByte()
var namebytes = newname.toByteArray()
if (!ByteTools.isUnicode(newname))
{
grbitChr = 0x0
}
else
{
grbitChr = 0x1
}
try
{
if (grbitChr.toInt() == 0x1)
{
namebytes = newname.toByteArray(charset(WorkBookFactory.UNICODEENCODING))
}
else
{
namebytes = newname.toByteArray(charset(WorkBookFactory.DEFAULTENCODING))
}
}
catch (e:UnsupportedEncodingException) {
namebytes = newname.toByteArray()
Logger.logWarn("UnsupportedEncodingException in setting sheet name: "
+ e + " falling back to system default.")
}

val newdata = ByteArray(namebytes.size + 8)
if (data == null)
this.data = newdata
else
System.arraycopy(this.getData()!!, 0, newdata, 0, 8)

System.arraycopy(namebytes, 0, newdata, 8, namebytes.size)
newdata[6] = cch
newdata[7] = grbitChr
this.setData(newdata)
this.init()
}

/**
 * Returns a serialized copy of this Boundsheet
 *
 * @throws IOException
 */
    override val sheetBytes:ByteArray?
@Throws(IOException::class)
get() {
this.setLocalRecs()
var obs:ObjectOutputStream? = null
var b:ByteArray? = null

val baos = ByteArrayOutputStream()
obs = ObjectOutputStream(baos)
obs.writeObject(this)
b = baos.toByteArray()

return b
}

/**
 * get the type of sheet as a string
 */
    override val sheetTypeString:String?
get() {
when (sheetType) {
SHEET_DIALOG -> return "Sheet or Dialog"
XL4_MACRO -> return "XL4 Macro"
CHART -> return "Chart"
VBMODULE -> return "VB Module"
else -> return null
}
}

/**
 * @return Returns the localrecs.
 */
    override val localRecs:List<*>?
get() =localrecs

/**
 * Gets the printer setup handle for this sheet.
 */
     val printerSetupHandle:PrinterSettingsHandle
get() =PrinterSettingsHandle(this)

/**
 * return the default column width in # characters of the maximum digit width of the normal style's font
 *
 *
 * This is currently a floating point value, something I question.  I don't understand the need for this,
 * and possibly it should be an int?
 */
    /**
 * set the default column width in # characters of the maximum digit width of the normal style's font
 */
    // biff8 setting
 // ooxml setting
 // biff8 setting
 var defaultColumnWidth:Float
get() =if (defColWidth != null) {
            defColWidth!!.defaultWidth.toFloat()
        } else defaultColWidth
set(w) {
defaultColWidth = w
if (defColWidth != null)
{
defColWidth!!.setDefaultColWidth(w.toInt())
}
}

/**
 * Get all the names for this boundsheet
 *
 * @return
 */
     val allNames:Array<Name>
get() {
if (this.sheetNameRecs == null)
this.sheetNameRecs = HashMap()
val retnames = ArrayList(sheetNameRecs!!.values)
val names = arrayOfNulls<Name>(retnames.size)
return retnames.toTypedArray() as Array<Name>
}

/**
 * given sheet.xml input stream, parse OOXML into the current sheet
 *
 * @param bk
 * @param sheet
 * @param ii
 * @param sst               The sst.
 * @param formulas          Arraylist stores all formulas/info - must be added after all sheets and cells
 * @param hyperlinks
 * @param inlineStrs        Hashmap stores inline strings and cell addresses; must be added after all sheets and cells
 * @throws IOException
 * @throws XmlPullParserException
 * @throws CellNotFoundException
 */
    internal var shExternalLinkInfo:HashMap<String, String>? = null

override fun setHeader(h:BiffRec) {
this.header = h as Headerrec
}

override fun setFooter(ftr:BiffRec) {
this.footer = ftr as Footerrec
}

/**
 * Please add comments for this method
 *
 * @param bAddUnconditionally
 */
    @JvmOverloads  fun insertImage(im:ImageHandle, bAddUnconditionally:Boolean = false) {
var msodg = this.wkbook!!.msoDrawingGroup
val msoDrawing = MSODrawing.prototype as MSODrawing?
msoDrawing!!.setSheet(this)
msoDrawing.coords = im.coords

im.setMsgdrawing(msoDrawing) // 20070924 KSC: link 2 actual msodrawing
 // that describes this image for setting
        // bounds, etc.
        var insertIndex = -1
val obj = Obj.prototype as Obj?
 // now add to proper place in stream
        if (msodg != null)
{ // already have drawing records; just add to
 // records + update msodg
            insertIndex = this.getIndexOf(XLSConstants.MSODRAWINGSELECTION)
if (insertIndex < 0)
insertIndex = this.getIndexOf(XLSConstants.WINDOW2)
if (msodg.getMsoHeaderRec(this) == null)
 // handle case of multiple
                // sheets- each needs
                // it's own mso header
                // ...
                msoDrawing.setIsHeader()
}
else
{ // No images present in workbook, must add appropriate
 // records
            // Create new msodg rec
            this.wkbook!!.msoDrawingGroup = MSODrawingGroup
.prototype as MSODrawingGroup?
msodg = this.wkbook!!.msoDrawingGroup
msodg!!.initNewMSODrawingGroup() // generate and add required records
 // for drawing records
            // also add 1st portion for drawing rec
            msoDrawing.setIsHeader()
 // insertion point for new msodrawing rec
            insertIndex = getIndexOf(XLSConstants.DIMENSIONS) + 1
}
if (insertIndex > 0)
{ // should! then have a drawing record to insert
 // 20071120 KSC: retrieve idx in order to reuse/link to
            // existing image bytes if duplicating images
            val idx = msodg.addImage(im.imageBytes, im
.imageType, bAddUnconditionally)
imageMap!![im] = Integer.valueOf(im.imageIndex - 1) // add
 // new
            // image
            // to
            // map
            // and
            // link
            // to
            // actual
            // imageIndex
            // -
            // moved
            // from
            // above
            msoDrawing.createRecord(++this.wkbook!!.lastSPID, im
.imageName, im.shapeName, idx) // generate
 // msoDrawing
            // using correct
            // values moved
            // from above
            this.SheetRecs.add(insertIndex++, msoDrawing)
this.SheetRecs.add(insertIndex++, obj)
msodg.addMsodrawingrec(msoDrawing) // add the new drawing rec to
 // the msodrawinggroup set of
            // recs
            wkbook!!.updateMsodrawingHeaderRec(this) // find the msodrawing
 // header record and update
            // it (using info from other
            // msodrawing recs)
            // 20080908 KSC: moved from above
            msodg.spidMax = this.wkbook!!.lastSPID + 1 // was ++lastSPID
msodg.updateRecord() // given all information, generate appropriate
 // bytes
            msodg.dirtyflag = true
}
else
{
Logger.logErr("Boundsheet.insertImage:  Drawing Group not created.")
}
}

/**
 * Rationalizes the itab (sheet reference) for name records,
 * this has to occur after sheet insert/delete operations to keep the
 * references intact.  Unfortunately these references do not use the Externsheet,
 * so are not ilbl listeners.
 */
    internal fun updateLocalNameReferences() {
if (sheetNameRecs == null)
return
val i = this.sheetNameRecs!!.values.iterator()
while (i.hasNext())
{
val n = i.next() as Name
n.setItab((this.sheetNum + 1).toShort())
}
}

/**
 * column formatting records
 *
 *
 * Note that it checks if exists.  This is due to externally copied boundsheets already having
 * the record in the array when addrecord occurs.
 */
    override fun addColinfo(c:Colinfo) {
if (!this.colinfos.containsValue(c))
{
this.colinfos[c] = c
}
}

/**
 * For workbooks that do not contain a dval record,
 * insert a default dval rec
 *
 * @return
 */
     fun insertDvalRec():Dval {
if (this.dvalRec != null)
return this.dvalRec
val d = Dval.prototype as Dval?
d!!.setSheet(this)
var insertIdx = window2!!.recordIndex + 1
 // correct position for DV block is before sheet protection
        // records (if any)
        // or before EOF
        var opc = (SheetRecs[insertIdx] as BiffRec).opcode.toInt()
while (opc != XLSConstants.EOF.toInt())
{
if (opc == XLSConstants.SHEETPROTECTION.toInt() || opc == XLSConstants.RANGEPROTECTION.toInt()
|| opc == XLSConstants.SHEETLAYOUT.toInt())
break
insertIdx++
opc = (SheetRecs[insertIdx] as BiffRec).opcode.toInt()
}
this.SheetRecs.add(insertIdx, d)
this.dvalRec = d
return d
}

/**
 * Create a dv (validation record)
 * record gets inserted into the byte stream from
 * within Dval
 *
 * @param location
 * @return
 */
     fun createDv(location:String):Dv {
if (this.dvalRec == null)
this.insertDvalRec()
val dv = this.dvalRec!!.createDvRec(location)
var insertIdx = this.SheetRecs.size - 2 // start at 1 before EOF
var opc = (SheetRecs[insertIdx] as BiffRec).opcode.toInt()
while (opc != XLSConstants.DV.toInt() && opc != XLSConstants.DVAL.toInt())
{
insertIdx-- // insert after last DV
opc = (SheetRecs[insertIdx] as BiffRec).opcode.toInt()
}
this.SheetRecs.add(insertIdx + 1, dv) // insert after DVAL or last DV
return dv
}

/**
 * Create a Condfmt (Conditional format) record and
 * add it to sheet recs
 *
 * @param location
 * @return
 */
     fun createCondfmt(location:String, wbh:WorkBookHandle):Condfmt {
val cfx = Condfmt.prototype as Condfmt?
var insertIdx = window2!!.recordIndex + 1
var rec = this.SheetRecs[insertIdx] as BiffRec
while (rec.opcode != XLSConstants.HLINK && rec.offset != XLSConstants.DVAL.toInt()
&& rec.opcode.toInt() != 0x0862 /* SHEETLAYOUT */
&& rec.opcode.toInt() != 0x0867 /* SHEETPROTECTION */
&& rec.opcode.toInt() != 0x0868 /* RANGEPROTECTION */
&& rec.opcode != XLSConstants.EOF)
rec = this.SheetRecs[++insertIdx] as BiffRec

this.SheetRecs.add(insertIdx, cfx)
cfx!!.streamer = streamer
cfx.workBook = this.workBook
cfx.resetRange(location)
this.addConditionalFormat(cfx)
cfx.setSheet(this)
return cfx
}

/**
 * Create a Cf (Conditional format rule) record and
 * add it to sheet recs
 *
 * @param Conditional format
 * @param range
 * @return
 */
     fun createCf(cfx:Condfmt):Cf {
val cf = Cf.prototype as Cf?
 // we add this rec to vec right after its Condfmt
        val insertIdx = cfx.recordIndex + 1
this.SheetRecs.add(insertIdx, cf)
cf!!.streamer = streamer
cf.workBook = this.workBook
cf.setSheet(this)
cf.condfmt = cfx
cfx.addRule(cf)
return cf
}

/**
 * obtain the desired image handle via the MsoDrawing Image Index
 * used for mapping images from copied worksheets
 *
 * @param index
 * @return
 */
     fun getImageByMsoIndex(index:Int):ImageHandle? {
if (imageMap == null)
return null
val ir = imageMap!!.keys.iterator()
var ret:ImageHandle? = null
while (ir.hasNext() && ret == null)
{
val im = ir.next() as ImageHandle
if (im.msodrawing!!.getImageIndex() == index)
ret = im
}
return ret
}

 // Generic getIndexOf - replace specific hardocoded cases
    // ...
     fun getIndexOf(opc:Short):Int {
var rec:BiffRec? = null

val size = SheetRecs.size
var foundIndex = -1
var i = 0
while (i < size && foundIndex == -1)
{
rec = SheetRecs[i] as BiffRec
if (rec.opcode == opc)
foundIndex = i
i++
}
return foundIndex
}

/**
 * return the desired record from the sheetrecs, or null if doesn't exist
 *
 * @param opc
 * @return
 */
     fun getSheetRec(opc:Short):BiffRec? {
var rec:BiffRec? = null

val size = SheetRecs.size
val foundIndex = -1
var i = 0
while (i < size && foundIndex == -1)
{
rec = SheetRecs[i] as BiffRec
if (rec.opcode == opc)
return rec
i++
}
return null
}

 // 20070916 KSC: access for inserting records into sheetrecs
    // collection
     fun insertSheetRecordAt(r:BiffRec, index:Int) {
r.setSheet(this)
if (index > -1 && index < SheetRecs.size)
SheetRecs.add(index, r)
else
SheetRecs.add(r)
}

/**
 * Sheet hash is a cross-workbook identifier for OpenXLS.  The first time it is called it creates the sheet hash.
 *
 * @return
 */
     fun getSheetHash():String {
if (sheetHash == "")
{
sheetHash = (this.sheetName + this.sheetNum
+ this.realRecordIndex)
}
return sheetHash
}

/**
 * assembleSheetRecs assembles the array of records, then ouputs
 * the ordered list to the bytestreamer, which should be the only
 * thing calling this.
 */
     fun assembleSheetRecs():List<*> {
return WorkBookAssembler.assembleSheetRecs(this)
}

/**
 * write this sheet as tabbed text output: <br></br>
 * All rows and all characters in each cell are saved. Columns of data are
 * separated by tab characters, and each row of data ends in a carriage
 * return. If a cell contains a comma, the cell contents are enclosed in
 * double quotation marks. All formatting, graphics, objects, and other
 * worksheet contents are lost. The euro symbol will be converted to a
 * question mark. If cells display formulas instead of formula values, the
 * formulas are saved as text.
 */
    /*
     * From Excel: To preserve the formulas if you reopen the
     * file in Microsoft
     * Excel, select the Delimited option in the Text Import
     * Wizard, and select
     * tab characters as the delimiters. Note If your workbook
     * contains special
     * font characters, such as a copyright symbol (©), and you
     * will be using
     * the converted text file on a computer with a different
     * operating system,
     * save the workbook in the text file format appropriate for
     * that system.
     * For example, if you are using Windows and want to use the
     * text file on a
     * Macintosh computer, save the file in the Text (Macintosh)
     * format. If you
     * are using a Macintosh computer and want to use the text
     * file on a system
     * running Windows or Windows NT, save the file in the Text
     * (Windows)
     * format.
     */
    @Throws(IOException::class)
 fun writeAsTabbedText(dest:OutputStream) {
val lastrow = this.maxRow
val lastcol = this.maxCol
var isInteger = false
val tab = byteArrayOf(9)
val crlf = byteArrayOf(13, 10)
for (i in 0 until lastrow)
{
val r = this.getRowByNumber(i)
if (r != null)
{
for (j in 0 until lastcol)
{
var c:BiffRec? = null
try
{
 // Look for the cell and output
                        c = r.getCell(j.toShort())
val type = (c as XLSRecord).cellType
val o:Any?
if (type != Cell.TYPE_FORMULA)
{
isInteger = type == Cell.TYPE_INT
o = c.stringVal
}
else
{
o = (c as Formula).calculateFormula()
isInteger = o is Int || o is Double && o.toInt().toDouble() == o
.toDouble()
}
try
{
var output = CellFormatFactory
.fromPatternString(c.formatPattern)
.format(o!!.toString())
if (output.indexOf(",") != -1)
output = "\"" + output + "\""
dest.write(output.toByteArray())
}
catch (e:Exception) {
Logger.logWarn("Boundsheet.writeAsTabbedText: error writing "
+ c.cellAddress + ":" + e.toString())
}

}
catch (e1:CellNotFoundException) {
 // No cell exists at this location, continue
                    }

dest.write(tab)
}
 // } catch (RowNotFoundException e) { }
            }
dest.write(crlf)
}
dest.flush()
dest.close()
}

/**
 * do all of the expensive updating here
 * only right before streaming record.
 */
    override fun preStream() {}

/**
 * add a new Condtional Format rec for this sheet
 *
 * @param cf
 */
     fun addConditionalFormat(cf:Condfmt) {
if (cond_formats == null)
cond_formats = ArrayList()
if (cond_formats!!.indexOf(cf) == -1)
cond_formats!!.add(cf)
}

/**
 * set/save the Pane rec for this sheet
 * also links the Window2 rec to the pane rec
 *
 * @param p
 */
     fun setPane(p:Pane?) {
var p = p
if (p == null)
{ // adds new
p = Pane.prototype as Pane?
val insertIdx = window2!!.recordIndex + 1
this.SheetRecs.add(insertIdx, p)
}
pane = p
pane!!.setWindow2(window2)
}

/**
 * retrieve the Pane rec for this sheet
 *
 * @return
 */
     fun getPane():Pane? {
return pane
}

/**
 * remove pane rec, effectively unfreezing
 */
     fun removePane() {
SheetRecs.remove(pane)
pane = null
}

/**
 * Remove a BiffRec from this WorkSheet.
 *
 */
    @Deprecated("Use {@link #removeCell(int, int)} instead.")
override fun removeCell(celladdr:String) {
val c = this.getCell(celladdr)
if (c != null)
{
this.removeCell(c)
}
}

/**
 * Remove a BiffRec from this WorkSheet.
 */
     fun removeCell(row:Int, col:Int) {
val c:BiffRec
try
{
c = this.getCell(row, col)
this.removeCell(c)
}
catch (e:CellNotFoundException) {
 // cell does not exist, this is fine
        }

}

/**
 * remove a BiffRec from the worksheet.
 *
 *
 * Unfortunately this also has to manage mulrecs
 */
    override fun removeCell(cell:BiffRec) {
if (cell.opcode == XLSConstants.MULBLANK)
{
(cell as Mulblank).removeCell(cell.colNumber)
}
if (cell.opcode == XLSConstants.FORMULA)
{
val f = cell as Formula
this.wkbook!!.removeFormula(f)
}
cellsByRow.remove(cell)
cellsByCol.remove(cell)
}

/**
 * removes an image from the imagehandle cache (should be in WSH)
 *
 *
 * Jan 22, 2010
 *
 * @param img
 * @return
 */
     fun removeImage(img:ImageHandle):Boolean {
return imageMap!!.remove(img) != null
}

/**
 * remove a record from the vector via it's index
 * into the SheetRecs aray, includes firing a change event
 *
 * @param idx
 */
     fun removeRecFromVec(idx:Int) {
try
{
val rec = this.SheetRecs[idx] as BiffRec
removeRecFromVec(rec)
}
catch (e:Exception) {
Logger.logErr("Boundsheet.removeRecFromVec: $e")
}

}

/**
 * Removes some rows and all associated cells from this sheet.
 * References are not handled; for those see [ReferenceTracker].
 *
 * @param first the zero-based index of the first row to be removed
 * @param count the number of rows to be removed
 * @param shift whether to shift subsequent rows up to fill the empty space
 */
     fun removeRows(first:Int, count:Int, shift:Boolean) {

for (rowIdx in first until first + count)
{
 // this.removeRowContents(rowIdx);

            val row = rows[rowIdx] ?: continue

val iter = row.cells.iterator()
while (iter.hasNext())
{
val cell = iter.next()

 // This removes the cell from the Row's map without
                // perturbing
                // the iterator. When removeCell tries to remove it later
                // the
                // map will silently do nothing instead of throwing a CME.
                iter.remove()

this.removeCell(cell)
}

rows.remove(rowIdx)
this.removeRecFromVec(row)
}

 // shift all following rows up to fill the gap left by the
        // removed rows
        if (shift && !rows.isEmpty())
{
val shiftBy = -1 * count
val lastrow = lastRow!!.rowNumber
for (rowIdx in first + 1..lastrow)
{
val row = rows[rowIdx] ?: continue

this.shiftRow(row, shiftBy)
}
}

 // update sheet dimensions
        this.dimensions!!
.setRowLast(if (null != lastRow) lastRow!!.rowNumber else 0)
}

/**
 * remove rec from the vector, includes firing
 * a changeevent.
 */
    override fun removeRecFromVec(rec:BiffRec) {
var removerec = true
 // is it an RK, maybe part of a Mulrk??
        if (rec.opcode == XLSConstants.RK)
{
val thisrk = rec as Rk
this.removeMulrk(thisrk)
}
else if (rec.opcode == XLSConstants.FORMULA)
{
val f = rec as Formula
this.wkbook!!.removeFormula(f)
}
else if (rec.opcode == XLSConstants.LABELSST)
{
val lst = rec as Labelsst
val strtable = wkbook!!.sharedStringTable
lst.initUnsharedString()
strtable!!.removeUnicodestring(lst.unsharedString!!)
}
else if (rec is Mulblank)
{ // KSC: Added
removerec = rec.removeCell(rec.colNumber)
}
if (removerec)
{
if (streamer!!.removeRecord(rec))
{
if (DEBUGLEVEL > 5)
Logger.logInfo("Boundsheet RemoveRec Removed: $rec")
}
else
{
if (rec is Mul)
{
if (!(rec as Mul).removed())
if (DEBUGLEVEL > 1)
Logger.logWarn("RemoveRec failed: "
+ rec.javaClass.getName()
+ " not found in Streamer Vec")
}
else
{
if (DEBUGLEVEL > 1)
Logger.logWarn("RemoveRec failed: "
+ rec.javaClass.getName()
+ " not found in Streamer Vec")
}
}
}
}

/**
 * Called from removeCell(), removeMulrk() handles the fact that you
 * are trying to delete a rk that is really just a part of a Mulrk.  This
 * is handled by truncating the mulrk at the cell, then creating individual numbers
 * after the deleted cell.
 */
    override fun removeMulrk(thisrk:Rk) {
val mymul = thisrk.myMul as Mulrk
if (mymul != null)
{ // Part of a mulrk. JOY!
val vect = mymul.removeRk(thisrk)
var deletemulrk = false
if (mymul.getColFirst() == thisrk.colNumber.toInt())
{
deletemulrk = true
}
if (vect != null)
{ // the mulrk contiued past the cell deleted
 // Create new records for each of the Rks,
                val itv = vect.iterator()
while (itv.hasNext())
{
val temprk = itv.next() as Rk
temprk.setNoMul()
val loc = temprk.cellAddress

val d = temprk.dblVal
val g = this.getCell(loc)
val fmt = g!!.ixfe
g.row.removeCell(g)
this.removeCell(loc)

this.addValue(d, loc)
this.getCell(loc)!!.ixfe = fmt
streamer!!.removeRecord(temprk)
}
}
if (deletemulrk)
{
mymul.removed = true
this.removeRecFromVec(mymul)
}
}

}

/**
 * Remove a row, do not shift any other rows
 *
 * @throws RowNotFoundException
 */
    @Throws(RowNotFoundException::class)
 fun removeRowContents(rownum:Int) {
var r = this.getRowByNumber(rownum)
 // First delete the desired row
        if (r != null)
{
val cells = r.cellArray
for (x in cells.indices)
{ // adjust cell's in row
this.removeCell(cells[x] as BiffRec)
}
rows.remove(Integer.valueOf(rownum))
this.removeRecFromVec(r)
r = null
}
else
{
throw RowNotFoundException(this.sheetName + ":" + rownum)
}

}

/**
 * Removes a set of columns and their associated cells from this sheet.
 * Optionally shifts the subsequent columns left to fill the empty space.
 * This method only updates the sheet and cell records. It doesn't adjust
 * references; that's handled by [WorkSheetHandle.removeCols].
 *
 * @param first this zero-based index of the first column to be removed
 * @param count the number of columns to remove
 * @param shift whether to shift subsequent columns left
 */
     fun removeCols(first:Int, count:Int, shift:Boolean) {

if (shift)
{
ReferenceTracker.updateReferences(first, count * -1, this, false) // shift
 // or
            // expand/contract
            // ALL
            // affected
            // references
            // including
            // named
            // ranges
        }

for (colIdx in first until first + count)
{

 // update or remove the ColInfo record as appropriate
            val info = this.getColInfo(colIdx)
if (null != info)
{
if (info.colLast < first + count)
{
if (info.colFirst >= first)
{
this.removeColInfo(info)
}
else
{
info.colLast = first - 1
}
}
else if (info.colFirst >= first)
{
info.colFirst = first + count
}
}

 // remove the cells in the column
            try
{
val cells = this.getCellsByCol(colIdx)
val cellCount = cells.size
for (idx in cellCount - 1 downTo 0)
{
val cell = cells[idx] ?: continue

this.removeCell(cell)
}
}
catch (e:CellNotFoundException) {
 // This is fine, no cells in this column
            }

}

if (shift)
{
val shiftBy = -1 * count
val maxcol = this.realMaxCol
for (colIdx in first + 1..maxcol)
{
this.shiftCol(colIdx, shiftBy)
}
}

 // make sure dimensions record is correctly updated upon
        // output
        this.dimensions!!.setColLast(this.realMaxCol)
}

/**
 * Access an arrayList of cells by column
 *
 * @param colNum
 * @return
 */
    @Throws(CellNotFoundException::class)
 fun getCellsByCol(colNum:Int):ArrayList<BiffRec> {
val theCells = cellsByCol
.subMap(CellAddressible.RangeBoundary(0, colNum,
true), CellAddressible.RangeBoundary(0, colNum + 1,
false))
if (theCells.size == 0)
throw CellNotFoundException(this.sheetname, 0, colNumber.toInt())

val cells = theCells.values
val i = cells.iterator()
while (i.hasNext())
{
val biffrec = i.next()
if (biffrec.opcode == XLSConstants.MULBLANK)
{
(biffrec as Mulblank).setCurrentCell(colNum.toShort())
}
}
return ArrayList(cells)
}

/**
 * Access an arrayList of cells by column
 *
 * @param colNum
 * @return
 */
    @Throws(CellNotFoundException::class)
 fun getCellsByRow(rowNum:Int):ArrayList<BiffRec> {
val theCells = cellsByRow
.subMap(CellAddressible.Reference(rowNum,
0), CellAddressible.Reference(rowNum + 1, 0))
if (theCells.size == 0)
throw CellNotFoundException(this.sheetname, 0, colNumber.toInt())

val cells = theCells.values
return ArrayList(cells)
}

/**
 * get a handle to a specific column of cells in this sheet
 */
    override fun getColInfo(col:Int):Colinfo? {
val info = this.colinfos[ColumnRange.Reference(col, col)] ?: return null

return if (info.inrange(col)) info else null
}

/**
 * remove all Sheet records from Sheet.
 */
    override fun removeAllRecords() {
 // this.setSheetRecs();
        val rx = arrayOfNulls<XLSRecord>(SheetRecs.size)
SheetRecs.toTypedArray()

for (t in rx.indices)
{
val opcode = (rx[t] as BiffRec).opcode.toInt() // Handle continues
 // masking mso's
            if (opcode != XLSConstants.MSODRAWING.toInt() && !(opcode == XLSConstants.CONTINUE.toInt() && (rx[t] as Continue).maskedMso != null))
this.removeRecFromVec(rx[t])
else // must update MSODrawingGroup record as well ...
                if (opcode == XLSConstants.MSODRAWING.toInt())
this.wkbook!!.msodg!!
.removeMsodrawingrec(rx[t] as MSODrawing, this, false) // don't
else
 // a Continue record masking an MSoo
                    this.wkbook!!.msodg!!
.removeMsodrawingrec((rx[t] as Continue).maskedMso!!, this, false)// remove
 // assoc
 // object
 // record
 // don't
 // remove
            // assoc
            // object
            // record
            rx[t] = null
}
SheetRecs.clear()
 // System.gc();

    }

/**
 * Shifts a single column.
 * This adjusts any mention of the column number in the associated records.
 * References are not handled; for those see [ReferenceTracker].
 *
 * @param col   the column to be shifted
 * @param shift the number of columns by which to shift
 */
    private fun shiftCol(colNum:Int, shift:Int) {
val info = this.getColInfo(colNum)
val newCol = colNum + shift

val cells:List<BiffRec>
try
{
cells = this.getCellsByCol(colNum)
for (cell in cells)
{
cell.setCol(newCol.toShort())
this.updateDimensions(cell.rowNumber, cell.colNumber.toInt())
}
}
catch (e:CellNotFoundException) {
 // No cells exist in this column
        }

if (null != info)
{
val first = info.colFirst
if (first == colNum || first > newCol)
{
info.colFirst = newCol
}

val last = info.colLast
if (last == colNum || last < newCol)
info.colLast = newCol
}
}

private fun removeColInfo(ci:Colinfo) {
this.removeRecFromVec(ci)
this.colinfos.remove(ci)
}

 /*
     *
     */

    /**
 * set the Bof record for this Boundsheet
 */
    override fun setBOF(b:Bof) {
myBof = b
b.setSheet(this)
}

override fun setEOF(f:Eof) {
myEof = f
}

/**
 * return the pos of the Bof for this Sheet
 */
    override fun getLbPlyPos():Long {//
return if (myBof != null) myBof!!.lbPlyPos else this.lbPlyPos

}

/**
 * set the pos of the Bof for this Sheet
 */
    override fun setLbPlyPos(newpos:Long) {
val newposbytes = ByteTools.cLongToLEBytes(newpos.toInt())
System.arraycopy(newposbytes, 0, this.getData()!!, 0, 4)
this.lbPlyPos = newpos
}

/**
 * the beginning of the Dimensions record
 * is the index of the RowBlocks
 */
    override fun getDimensions():Dimensions? {
return dimensions
}

override fun setDimensions(d:Dimensions) {
 // only set the first dimensions. Other dimensions records
        // may exist within
        // the boundsheet stream from charts & msodrawing objects,
        // but going to run with the
        // assumption that the first one is the identifier for
        // valrec start
        if (dimensions == null)
{
dimensions = d
if (myidx != null)
this.myidx!!.setDimensions(d)
}

}

/**
 * Shifts a single row.
 * This adjusts any mention of the row number in the row records. Formula
 * references are not handled; for those see [ReferenceTracker].
 *
 * @param row   the row to be shifted
 * @param shift the number of rows by which to shift
 */
    private fun shiftRow(row:Row, shift:Int) {
val cells = row.cells.iterator()
var skipMulBlank:Mulblank? = null
while (cells.hasNext())
{
val cell = cells.next() as BiffRec

if (cell === skipMulBlank)
continue
else if (cell.opcode == XLSConstants.MULBLANK)
skipMulBlank = cell as Mulblank

this.shiftCellRow(cell, shift)
}

val oldRow = row.rowNumber
val newRow = oldRow + shift
row.rowNumber = newRow

rows.remove(oldRow)
rows[newRow] = row

if (this.dimensions!!.getRowLast() < newRow)
{
this.dimensions!!.setRowLast(newRow)
this.lastRow = row
}
}

/**
 * update the INDEX record with the new max Row #
 * why we need so many redundant references to the Min/Max Row/Cols
 * is a question for the Redmond sages.
 */
    override fun updateDimensions(row:Int, c:Int) {
if (DEBUGLEVEL > 10)
Logger.logInfo("Boundsheet Updating Dimensions: " + row + ":"
+ colNumber)
val col = c.toShort()
realMaxCol = Math.max(realMaxCol, col.toInt())
maximumCellRow = Math.max(maximumCellRow, row)
if (dimensions != null)
dimensions!!.updateDimensions(row - 1, col)
if (this.myidx != null)
this.myidx!!.updateRowDimensions(this.minRow, this.maxRow) // TODO:
 // investigate
        // why
        // no
        // Index
        // is
        // possible
    }

/**
 * set the associated sheet index
 */
    override fun setSheetIDX(idx:Index) {
idx.setSheet(this)
myidx = idx
}

/**
 * set the associated sheet index
 */
     fun getSheetIDX():Index? {
return myidx
}

/**
 * Adjusts a cell to reflect its parent row being shifted.
 * This adjusts any mention of the row number in the cell record. Formula
 * references are not handled; for those see [ReferenceTracker].
 *
 * @param cell  the cell record to be shifted
 * @param shift the number of rows by which to shift the cell
 */
    private fun shiftCellRow(cell:BiffRec, shift:Int) {
val newrow = cell.rowNumber + shift
cell.rowNumber = newrow

 // handle per-record special cases
        when (cell.opcode) {
XLSConstants.RK -> (cell as Rk).setMulrkRow(newrow)

XLSConstants.FORMULA -> {
val formula = cell as Formula

 // must also shift shared formulas if necessary
                if (formula.isSharedFormula)
{
if (formula.internalRecords.size > 0)
{// is it the
 // parent?
                        val o = formula.internalRecords[0]
if (o is Shrfmla)
{ // should!
o.firstRow = o.firstRow + shift
o.lastRow = o.lastRow + shift
}
}
}
}
}
}

/**
 * add a row to the worksheet as well
 * as to the RowBlock which will handle
 * the updating of Dbcell index behavior
 *
 * @param BiffRec the cell being added (can't add a row without one...)
 */
    private fun addNewRow(cell:BiffRec):Row? {
val rn = cell.rowNumber
if (this.getRowByNumber(rn) != null)
return this.getRowByNumber(rn) // already exists!
val r = Row(rn, wkbook)
try
{ // Out-of-spec wb's may not have dimensions record -- will
 // be handled upon validation
            if (rn >= this.maxRow)
dimensions!!.setRowLast(rn)
}
catch (e:NullPointerException) {}

r.setSheet(this)
this.addRowRec(r)
return r

}

/**
 * Inserts a row and shifts subsequent rows down by one.
 *
 * @param rownum the zero-based index of the row to be created
 * @return the row that was just inserted
 */
    // TODO: reduce this functionality to simply inserting a row
    // and shifting the row number of subsequent rows and cells
     fun insertRow(rownum:Int, firstcol:Int, flag:Int, shiftrows:Boolean):Row? {
var roe:Row? = null
if (shiftrows && !this.fastCellAdds)
{
try
{
 // shift all rows after this one down...
                // moves refs, formats, merges, etc.
                if (lastRow != null)
{
var startrow = lastRow!!.rowNumber
if (startrow == XLSConstants.MAXROWS)
startrow-- // 20080925 KSC: can't add more than maxrows

for (t in startrow downTo rownum)
{ // traverse from
 // last row to
                        // current
                        val rowtoshift = rows[Integer.valueOf(t)]
if (rowtoshift != null)
{
try
{
this.shiftRow(rowtoshift, 1)// pass original
 // row # for
                                // formula
                                // shifting +
                                // flag
                            }
catch (e:Exception) {
Logger.logWarn("Boundsheet.insertRow() failed shifting row: "
+ t + " - " + e.toString())
}

}
}
}

 // we add a blank because a row cannot be empty
                roe = this.getRowByNumber(rownum)
if (roe == null)
{
val rc = intArrayOf(rownum, firstcol) // added firstcol to not
 // add bad cells at a1
                    this.addRecord(Blank.prototype!!, rc)
roe = this.getRowByNumber(rownum)
}

}
catch (a:Exception) {
Logger.logInfo("Boundsheet.insertRow:  Shifting row during Insert failed: $a")
}

}
else
{
roe = this.getRowByNumber(rownum)
if (roe == null)
{
val r = Row(rownum, wkbook)
 // must also update maxrow on sheet
                if (rownum >= this.maxRow)
dimensions!!.setRowLast(rownum)
r.setSheet(this)
this.addRowRec(r)
roe = this.getRowByNumber(rownum)
roe!!.resetCacheBytes()
}
}
return roe
}

/**
 * shifts Merged cells. 10-15-04 -jm
 */ // used???
    override fun updateMergedCells() {
if (this.mc.size < 1)
return
val mcs = this.mc.iterator()
while (mcs.hasNext())
(mcs.next() as Mergedcells).update()
}

/**
 * associate an existing Row with this Boundsheet
 * if the row already exists... ignore?
 */
     fun addRowRec(r:Row) {
val rwn = r.rowNumber
if (rows.containsKey(Integer.valueOf(rwn)))
{
if (DEBUGLEVEL > 2)
Logger.logWarn("Sheet.addRow() attempting to add existing row")
}
else
{
rows[Integer.valueOf(rwn)] = r
if (lastRow == null)
{
this.lastRow = r
}
else if (rwn > lastRow!!.rowNumber)
{
lastRow = r
}
}
}

/**
 * set whether this sheet is hidden upon opening
 */
    override fun setHidden(gr:Int) {
sheetType = gr.toShort()
val bt = ByteTools.shortToLEBytes(sheetType)
System.arraycopy(bt, 0, getData()!!, 4, 2)
}

/**
 * returns the selected sheet status
 */
    override fun selected():Boolean {
return selected
}

/**
 * set whether this sheet is selected upon opening
 */
    override fun setSelected(b:Boolean) {
if (this.window2 != null)
this.window2!!.setSelected(b)
if (b)
{
this.workBook!!.setSelectedSheet(this)
}
selected = b
}

/**
 * associate an Array formula with this Boundsheet
 */
     fun addArrayFormula(a:Array) {
arrayformulas.add(a)
}

/**
 * Returns an array formula for the set address
 */
     fun getArrayFormula(addr:String):Array? {
var form:Array? = null
for (i in arrayformulas.indices)
{
form = arrayformulas[i] as Array
if (form.isInRange(addr))
return form
}
return null
}

/**
 * map array formula range reference to the parent array formula address
 * <br></br>for Array Formula Parent Records only
 * boundsheet
 */
     fun addParentArrayRef(addr:String, ref:String) {
if (arrFormulaLocs.containsKey(addr))
Logger.logWarn("PARENT ARRAY ALREADY FOUND")
arrFormulaLocs.put(addr, ref)
}

/**
 * see if an array formula is part of an existing array formula
 * <br></br>by checking to see if the address in quesion is
 * referenced by any array formula references on this sheet
 *
 * @param rc row col of cell in question
 * @return
 * @see addArrayFormula
 */
     fun getArrayFormulaParent(rc:IntArray):Any? {
val i = arrFormulaLocs.keys.iterator()
while (i.hasNext())
{
val addr = i.next() as String
val arrayRC = ExcelTools
.getRangeRowCol(arrFormulaLocs.get(addr) as String)
if (rc[1] >= arrayRC[1] && rc[1] <= arrayRC[3]
&& rc[0] >= arrayRC[0] && rc[0] <= arrayRC[2])
{
return arrayRC
}
}
return null // no parent?
}

/**
 * return true if address refers to an Array Formula Parent
 * <br></br>i.e. the parent array formula refers to one or multiple cell addreses
 *
 * @param addr
 * @return
 */
     fun isArrayFormulaParent(addr:String):Boolean {
return arrFormulaLocs.get(addr) != null
}

/**
 * given an parent array formula at address formAddress,
 * look up in saved arrFormulaLocs for the cell range it references
 *
 * @param formAddress
 * @return
 */
     fun getArrayRef(formAddress:String):String {
return arrFormulaLocs.get(formAddress)
}

/**
 * inserts a col and shifts all of the other rows over one
 *
 * @param first zero-based int for the column (0='A')
 */
     fun insertCols(first:Int, count:Int) {

ReferenceTracker.updateReferences(first + 1, count, this, false) // shift
 // or
        // expand/contract
        // ALL
        // affected
        // references
        // including
        // named
        // ranges

        // shift the existing columns to the right to make room
        for (colIdx in this.realMaxCol downTo first)
{
this.shiftCol(colIdx, count)
}

 // update the new colinfos to include the formatting and the
        // width of the inserted col
        val movedCol = this.getColInfo(first + count)
if (movedCol != null)
{
for (i in 0 until count)
{
val newcol = this.getColInfo(first + i)
if (newcol == null)
{
this.addColinfo(first + i, first + i, movedCol
.colWidth, movedCol
.ixfe, movedCol.grbit)
}
else
{
newcol.grbit = movedCol.grbit
newcol.colWidth = movedCol.colWidth
newcol.ixfe = movedCol.ixfe
}
}
}

 // ensure the sheet bounds are accurate
        this.dimensions!!.setColLast(this.realMaxCol)
}

/**
 * Add a new colinfo
 *
 * @param first The beginning column number (0 based)
 * @param last  The end column number
 * @param width Initial width of the column
 * @param ixfe  formatting
 * @param grbit ??
 * @return Colinfo
 */
     fun createColinfo(first:Int, last:Int, width:Int, ixfe:Int, grbit:Int):Colinfo {
val ci = Colinfo.getPrototype(first, last, width, ixfe)
ci.grbit = grbit
ci.workBook = workBook
ci.setSheet(this)
this.addColinfo(ci)
var recpos = this.getDimensions()!!.recordIndex
recpos--
val sr = this.sheetRecs
 // get to last Colinfo record
        var rec = sr[recpos] as BiffRec
 // TODO: is it ABSOLUTELY true that if no Colinfos there
        // must be a DefColWidth record????
        while (rec !is Colinfo && rec !is DefColWidth
&& recpos > 0)
{ // loop until we find either a colinfo or
 // DEFCOLWIDTH
            rec = sr[--recpos] as BiffRec
}
 // now position this Colinfo in the proper position within
        // the Colinfo set
        val cf = ci.colFirst
while (rec is Colinfo && rec.colFirst > cf)
{
rec = sr[--recpos] as BiffRec
}
recpos++
this.streamer!!.addRecordAt(ci, recpos)
return ci
}

/**
 * Create a colinfo using the values from an existing colinfo
 *
 * @param first    first col in the colinfo
 * @param last     last col in the colinfo
 * @param template template column
 * @return
 */
     fun createColinfo(first:Int, last:Int, template:Colinfo):Colinfo {
return this.createColinfo(first, last, template.colWidth, template
.ixfe, template.grbit)
}

 fun createColinfo(first:Int, last:Int):Colinfo {
return this.createColinfo(first, last, Colinfo.DEFAULT_COLWIDTH, 0, 0)
}

/**
 * get a handle to the Row at the specified
 * row index
 *
 *
 * Zero-based Index.
 *
 *
 * ie: row 0 contains cell A1
 */
    override fun getRowByNumber(r:Int):Row? {
return rows[Integer.valueOf(r)]
}

/**
 * return an Array of the Rows
 */
    override fun getRows():Array<Row> {
val rxs = TreeMap(rows) // treemap does ordering... LHM does not
val rarr = arrayOfNulls<Row>(rxs.size)
return rxs.values.toTypedArray() as Array<Row>
}
 // safety checks

    override fun addValue(obj:Any, address:String):BiffRec {
return addValue(obj, address, false)
}

/**
 * Add a Value record to a WorkSheet.
 * This method's purpose is to handle default formatting
 * of the cell that is being added, and to do any manipulations
 * neccessary to handle mulrks, mulblanks, etc.  It is also the
 * main entry point of adding values to the boundsheet.   These values
 * are then passed into createValrec() which
 *
 * @param obj     the value of the new Cell
 * @param address the address of the new Cell
 */
     fun addValue(obj:Any, address:String, fixNumberAsString:Boolean):BiffRec {

 // first see if there's an existing item
        val rc = ExcelTools.getRowColFromString(address)

return addValue(obj, rc, fixNumberAsString)

}

/**
 * adds a value to the sheet
 *
 * @param obj
 * @param rc
 * @return
 */
     fun addValue(obj:Any, rc:IntArray, fixNumberAsString:Boolean):BiffRec {
return addValue(obj, rc, this.workBook!!
.defaultIxfe, fixNumberAsString)
}

/**
 * adds a cell to the Sheet
 *
 * @param obj
 * @param rc
 * @param FORMAT_ID
 * @param fixNumberAsString - whether to attempt to convert to a number if it is a NSaS situation
 * @return
 */
    @JvmOverloads  fun addValue(obj:Any?, rc:IntArray, FORMAT_ID:Int, fixNumberAsString:Boolean = false):BiffRec {
var obj = obj
var FORMAT_ID = FORMAT_ID

if (rc[1] > WorkBook.MAXCOLS)
throw InvalidRecordException("Cell Column number: " + rc[1]
+ " is greater than maximum allowable Columns: "
+ WorkBook.MAXCOLS)

 // sanity checks
        if (rc[0] > WorkBook.MAXROWS)
throw InvalidRecordException("Cell Row number: " + rc[0]
+ " is greater than maximum allowable row: "
+ WorkBook.MAXROWS)

var r = this.getRowByNumber(rc[0])

 /*
         * from Doc: The default cell format is always present in an
         * Excel file,
         * described by the XF record with the fixed index 15
         * (0-based).
         *
         * By default, it uses the worksheet/workbook default cell
         * style,
         * described by the very first XF record (index 0).
         */
        if (FORMAT_ID <= 0)
FORMAT_ID = this.workBook!!.defaultIxfe
if (FORMAT_ID == this.workBook!!.defaultIxfe)
{
if (this.getColInfo(rc[1]) != null)
{ /*
             * get default colinfo if
             * possible
             */
val co = this.getColInfo(rc[1])
if (co != null && co.ixfe != 0)
FORMAT_ID = co.ixfe
}
if (r != null && r.explicitFormatSet)
{
FORMAT_ID = r.ixfe
}
}

var merge_range:CellRange? = null

if (!fastCellAdds)
{
try
{
val mycell = this.getCell(rc[0], rc[1])
merge_range = mycell.mergeRange
 // specific cell format overrides any other formats: if
                // (FORMAT_ID == defaultFormatId)
                if (mycell.ixfe != this.workBook!!.defaultIxfe && mycell.ixfe != 0)
FORMAT_ID = mycell.ixfe
this.removeCell(mycell)
}
catch (cnfe:CellNotFoundException) {
 // good!
            }

}

 // Handle detection of Number stored as Strings
        var fixed:Array<Any>? = null
if (!fastCellAdds && fixNumberAsString && obj != null)
{
try
{
fixed = this.fixNumberStoredAsString(obj)
obj = fixed!![0]
}
catch (e:Exception) {
 // not a number!
            }

}

val rec = this.createValrec(obj, rc, FORMAT_ID)

if (!fastCellAdds)
{ // reapply conditional format and merges
if (merge_range != null)
rec.mergeRange = merge_range
}

 // check this does not touch affectedcells
        this.addRecord(rec, rc)

if (fixed != null)
{
val f = FormatHandle(this.wkbook,
this.workBook!!.defaultIxfe)
f.formatPattern = fixed!![1].toString()
rec.setXFRecord(f.formatId)
}

rec.resetCacheBytes()

if (r == null)
{ // if no row initially, check default row height; if
 // not Excel's default, set row height
            val rh = this.defaultRowHeight
if (rh != 12.75)
{// the default
r = this.getRowByNumber(rc[0])
r!!.rowHeight = (rh * 20).toInt()
}
}
return rec
}

/**
 * for numbers stored as strings, try to guess the
 * format pattern used, and strip the value to a number
 *
 *
 * TODO: increase sophistication of pattern matching to better guess pattern used
 *
 * @param s
 * @return Object[Double value, String formatPattern]
 */
    @Throws(NumberFormatException::class)
internal fun fixNumberStoredAsString(s:Any):Array<Any> {
var input = s.toString()
if (input.indexOf(" ") > -1)
{
input = StringTool.allTrim(input)
}
var p = "" // the format pattern
var matched = false

for (fmts in NumberAsStringFormat.values())
{
if (input.indexOf(fmts.identifier) > -1)
{
input = StringTool.strip(input, fmts.identifier)
p = fmts.pattern
matched = true
var d:Double? = Double(input)
d = fmts.adjustValue(d!!)
val ret = arrayOfNulls<Any>(2)
ret[0] = d // value
ret[1] = p // format pattern
return ret
}
}

throw NumberFormatException()
}

/**
 * A simple enum to store matching strings, format patterns, and value
 * switches where necessary
 */
    private enum class NumberAsStringFormat private constructor(private val identifier:String, private val pattern:String) {
PERCENT("%", "0%"),
EURO("€", "€#,##0;(€#,##0)"),
YEN("¥", "¥#,##0;(¥#,##0)"),
POUND("£", "£#,##0;(£#,##0)"),
DOLLAR("$", "$#,##0;(€#,##0)"),
ALT_POUND("₤", "₤#,##0;(₤#,##0)");

 fun identifier():String {
return this.identifier
}

 fun pattern():String {
return this.pattern
}

/**
 * adjust the value where necessary
 */
         fun adjustValue(inputVal:Double):Double {
return if (this.identifier === "%") inputVal * .01 else inputVal
}
}

/**
 * Creates a valrec (Value containing XLSRecord).   This method observes
 * the object passed in, then creates a XLS record of the correct type depending
 * on the object type. A default FormatID is handled as well.
 *
 *
 * The valrec at this point is not fully formed, it needs the row/col set
 * along with some other default actions that occur in addRecord(). This is
 * due to addRecord being the merge point for adding cells to a boundsheet between
 * the parse-level additions and the user-level additions!
 *
 * @param obj        the value of the new Cell
 * @param row        & col address of the new Cell
 * @param FORMAT_ID, index to the XF record for this valrec
 * @return partially formed XLS Record.
 */
    private fun createValrec(obj:Any?, rc:IntArray, FORMAT_ID:Int):XLSRecord {
 /*
         * try{
         * BiffRec cx = this.getCell(rc[0],rc[1]);
         * this.removeCell(cx);
         * }catch(CellNotFoundException e){}
         */
        var rec:XLSRecord? = null
if (obj == null)
{
rec = Blank()
}
else if (obj is Formula)
{
rec = obj
}
else if (obj is Double)
{
rec = NumberRec(obj.toDouble())
}
else if (obj is String)
{
if (obj.startsWith("="))
{
try
{
 // Logger.logInfo("adding formula");
                    rec = FormulaParser
.getFormulaFromString(obj as String?, this, rc)

 // this is a problem because workbook adds the rec to
                    // lastbounds
                    // in other words it will show up on the last sheet. Needed?
                    workBook!!.addRecord(rec!!, false)

 // getWorkBook().addFormula((Formula)rec); // next best
                    // thing methinks...
                }
catch (e:Exception) {
 // 20070212 KSC: add sheet name + address
                    // Logger.logWarn("adding new Formula at row:" + rc[0] +"
                    // col: " + rc[1] + ":"+obj.toString()+" failed, adding
                    // value to worksheet as String.");
                    throw FunctionNotSupportedException(
"Adding new Formula at " + this.sheetName + "!"
+ ExcelTools.formatLocation(rc)
+ " failed: " + e.toString() + ".")
 // rec = Labelsst.getPrototype((String) obj,
                    // this.getWorkBook().getSharedStringTable());
                }

}
else if (obj.startsWith("{="))
{ // interpret array
 // formulas as well
                // 20090526 KSC:
                // changed from "{"
                // to "{=" tracy vo
                // complex string
                // addition
                try
{
rec = FormulaParser
.getFormulaFromString(obj as String?, this, rc)
rec!!.isFormula = true
}
catch (e:Exception) {
throw FunctionNotSupportedException(
"Adding new Formula at " + this.sheetName + "!"
+ ExcelTools.formatLocation(rc)
+ " failed: " + e.toString() + ".")
}

}
else if (obj.toString().equals("", ignoreCase = true))
{
rec = Blank()
}
else
{
rec = Labelsst.getPrototype(obj as String?, this.workBook)
}
}
else if (obj is Int)
{
val l = obj.toInt()
rec = NumberRec(l)

}
else if (obj is Long)
{
val l = obj.toLong()
rec = NumberRec(l)
}
else if (obj is Boolean)
{
 // Logger.logErr("Adding Boolean Not Implemented");
            rec = Boolerr.prototype
rec!!.booleanVal = obj.booleanValue()
}
else
{
val d = Double(obj.toString()) // 20080211
 // KSC:
            // Double.valueOf(String.valueOf(obj)).doubleValue();
            rec = NumberRec(d)
}
rec!!.workBook = workBook
rec.setXFRecord(FORMAT_ID)
 // 20100607 KSC: update maxrow/maxcol if necessary
        if (rc[0] > maxRow || rc[1] > maxCol)
this.updateDimensions(rc[0], rc[1])
return rec
}

/**
 * Add an XLSRecord to a WorkSheet.
 *
 *
 * Creates the container cell for a record, sets the default
 * information on the valrec (ie row/col/bs), checks to see if
 * there is a container row for the cell, if not, then it creates the
 * row.  Finally, the cell is passed on to addCellToRowCol where it performs
 * final initialization and is added to it's row
 */
    override fun addRecord(rec:BiffRec, rc:IntArray) {
 // check to see if there is a BiffRec already at the address
        // add the rec to the Cell,
        // set as value if it's a val type rec

        rec.setSheet(this)// create a new BiffRec if none exists
rec.setRowCol(rc)
rec.isValueForCell = true
rec.streamer = streamer
rec.workBook = this.workBook

if (!this.fastCellAdds)
{

var ro:Row? = null
ro = rows[Integer.valueOf(rc[0])]
if (ro == null)
ro = this.addNewRow(rec)

}
if (copypriorformats && !this.fastCellAdds)
this.copyPriorCellFormatForNewCells(rec)

try
{
this.addCell(rec as CellRec)
}
catch (ax:ArrayIndexOutOfBoundsException) {
Logger.logErr("Boundsheet.addRecord() failed. Column " + rc[1]
+ " is greater than Maximum column count")
throw InvalidRecordException("Adding cell failed. Column "
+ rc[1] + " is greater than the maximum column limit.")
}

}

/**
 * Add a cell to this boundsheet record and populate the cells array
 *
 * @param cell
 */
    override fun addCell(cell:CellRec?) {
cellsByRow[cell] = cell
cellsByCol[cell] = cell
var row:Row? = rows[Integer.valueOf(cell!!.rowNumber)]
if (null == row)
row = this.addNewRow(cell)
row!!.addCell(cell)
cell?.setSheet(this)
this.updateDimensions(cell.rowNumber, cell.colNumber.toInt())
}

override fun setCopyPriorCellFormats(f:Boolean) {
this.copypriorformats = f
}

private fun copyPriorCellFormatForNewCells(c:BiffRec):Boolean {
val row = c.rowNumber + 1 // get the prior cell addy
val cnm = ExcelTools.getAlphaVal(c.colNumber.toInt())
val ch = this.getCell(cnm + row) ?: return false // try it...
c.ixfe = ch.ixfe
return true
}

/**
 * Add a new colinfo
 *
 * @param begCol The beginning column number (0 based)
 * @param endCol The end column number
 * @param width  Initial width of the column
 * @param ixfe   formatting
 * @param grbit  ??
 * @return Colinfo
 */
     fun addColinfo(begCol:Int, endCol:Int, width:Int, ixfe:Int, grbit:Int):Colinfo {
val ci = Colinfo.getPrototype(begCol, endCol, width, ixfe)
ci.grbit = grbit
ci.workBook = workBook
ci.setSheet(this)
this.addColinfo(ci)
var recpos = this.getDimensions()!!.recordIndex
recpos--
val sr = this.sheetRecs
 // get to last Colinfo record
        var rec = sr[recpos]
 // TODO: is it ABSOLUTELY true that if no Colinfos there
        // must be a DefColWidth record????
        while (rec !is Colinfo && rec !is DefColWidth
&& recpos > 0)
{ // loop until we find either a colinfo or
 // DEFCOLWIDTH
            rec = sr[--recpos]
}
 // now position this Colinfo in the proper position within
        // the Colinfo set
        val cf = ci.colFirst
while (rec is Colinfo && rec.colFirst > cf)
{
rec = sr[--recpos]
}
recpos++
this.streamer!!.addRecordAt(ci, recpos)
return ci
}

/**
 * get  a colinfo by name
 */
    override fun getColinfo(c:String):Colinfo? {
return this.getColInfo(ExcelTools.getIntVal(c))
}

/**
 * get the Collection of Colinfos
 */
    override fun getColinfos():Collection<Colinfo> {
return Collections.unmodifiableCollection(colinfos.values)
}

/**
 * Moves a cell location from one address to another
 */
     fun moveCell(startaddr:String, endaddr:String) {
val c = getCell(startaddr)
if (c!!.opcode == WorkBookFactory.RK)
{
try
{
val d = c.dblVal
this.removeCell(c)
this.addValue(d, endaddr)
}
catch (e:Exception) {
Logger.logInfo("Boundsheet.moveCell() error :$e")
}

}
else
{
val s = ExcelTools.getRowColFromString(endaddr)
c.setCol(s[1].toShort())
c.rowNumber = s[0]
removeCell(startaddr)
this.addCell(c as CellRec?)
}
}

/**
 * Moves a cell location from one address to another,
 * without any clearing of previous locations.  This is used in sorting
 * and other cell movements where we do not want to delete from starting address
 */
     fun updateCellReferences(c:BiffRec, endaddr:String) {
if (c.opcode == WorkBookFactory.RK)
{
try
{
val d = c.dblVal
this.removeCell(c)
this.addValue(d, endaddr)
}
catch (e:Exception) {
Logger.logInfo("Boundsheet.moveCell() error :$e")
}

}
else
{
val s = ExcelTools.getRowColFromString(endaddr)
c.setCol(s[1].toShort())
c.rowNumber = s[0]
this.addCell(c as CellRec)
}
}

/**
 * Gets a cell on this sheet by its Excel A1-style address.
 *
 * @param address the A1-style address of the cell to retrieve
 * @return the cell record
 * or `null` if no cell exists at the given address
 */
    @Deprecated("Use {@link #getCell(int, int)} instead.")
override fun getCell(address:String):BiffRec? {
val rc = ExcelTools.getRowColFromString(address)
try
{
return this.getCell(rc[0], rc[1])
}
catch (ex:CellNotFoundException) {
return null
}

}

/**
 * Gets a cell on this sheet by its row and column indexes.
 *
 * @param row the zero-based index of the cell's parent row
 * @param col the zero-based index of the cell's parent column
 * @return the cell record at the given address
 * @throws CellNotFoundException if no cell exists at the given address
 */
    @Throws(CellNotFoundException::class)
override fun getCell(row:Int, col:Int):BiffRec {
 // get the nearest entry from the cell map
        val theCell = cellsByRow[CellAddressible.Reference(row, col)] ?: throw CellNotFoundException(this.sheetname, row, col)

if (theCell != null && theCell.opcode == XLSConstants.MULBLANK)
{
(theCell as Mulblank).setCurrentCell(col.toShort())
}
return theCell
}

override fun addMergedCellsRec(r:Mergedcells) {
mc.add(r)
}

/**
 * Get the print area or titles name rec for this
 * boundsheet, return null if not exists
 *
 * @return
 */
    protected fun getPrintAreaNameRec(type:Byte):Name? {
val names = this.builtInNames
for (i in names.indices)
{
val n = names[i] as Name
if (n.builtInType == type)
{
return n
}
}
return null
}

/**
 * adds the _FILTERDATABASE name necessary for AutoFilter
 * if not already presetn
 */
    private fun addFilterDatabase() {
val names = this.builtInNames
var n:Name? = null
var i = 0
while (i < names.size && n == null)
{
if ((names[i] as Name)
.builtInType == Name._FILTER_DATABASE)
{
n = names[i] as Name
}
i++
}
if (n == null)
{ // not present
try
{
n = Name(this.workBook, "Built-in: _FILTER_DATABASE")
n.setBuiltIn(Name._FILTER_DATABASE)
val xref = this.workBook!!.getExternSheet(true)!!
.insertLocation(this.sheetNum, this.sheetNum)
n.setExternsheetRef(xref)
n.updateSheetReferences(this)
n.setSheet(this)
n.setIxals(this.sheetNum/* +1 */.toShort())
n.setItab((this.sheetNum + 1).toShort())
val loc = ExcelTools
.formatLocation(intArrayOf(this.minRow, this.minCol, this.maxRow - 1, this.maxCol - 1), false, false)
val s = Stack()
s.push(PtgRef.createPtgRefFromString(this.sheetName + "!"
+ loc, n))
n.expression = s
}
catch (e:Exception) {

}

}
}

/**
 * remove the _FILTER_DATABASE name (necessary for AutoFilters) for this sheet
 */
    private fun removeFilterDatabase() {
val names = this.builtInNames
var n:Name? = null
try
{
var i = 0
while (i < names.size && n == null)
{
if ((names[i] as Name)
.builtInType == Name._FILTER_DATABASE)
{
n = names[i] as Name
this.workBook!!.removeName(n)
break
}
i++
}
}
catch (e:Exception) {}

}

/**
 * Set the print area or titles for this worksheet.
 */
     fun setPrintArea(printarea:String, type:Byte) {
if (type != Name.PRINT_TITLES)
{// can have multiple print title refs
val n = this.getPrintAreaNameRec(type)
if (n != null)
{ // TODO: should check if same SHEET -- look at!
val s = n.expression
for (x in s!!.indices)
{
val p = s[x] as Ptg
if (p is PtgRef)
{
val ptg = PtgRef.createPtgRefFromString(printarea, n)
s.removeAt(x)
s.add(x, ptg)
}
}
return
}
}
 // create the name
        try
{
val t:String
if (type == Name.PRINT_AREA)
{
t = "PRINT_AREA"
}
else
{
t = "PRINT_TITLES"
}
val n = Name(this.workBook, "Built-in: $t")
n.setBuiltIn(type)
val xref = this.workBook!!.getExternSheet(true)!!
.insertLocation(this.sheetNum, this.sheetNum)
n.setExternsheetRef(xref)
n.updateSheetReferences(this)
n.setSheet(this)
n.setIxals(this.sheetNum.toShort())
n.setItab((this.sheetNum + 1).toShort())
val s = Stack()
val p = PtgRef.createPtgRefFromString(printarea, n)
s.push(p)
n.expression = s
}
catch (e:Exception) {
Logger.logErr("Error setting print area in boundsheet: $e")
}

}

/**
 * adds a merged cell to this sheet
 *
 * @return
 */
     fun addMergedCellRec():Mergedcells {
val mec = Mergedcells.prototype as Mergedcells?
mec!!.setSheet(this)
this.streamer!!.addRecordAt(mec, this.sheetRecs.size - 1)
this.addMergedCellsRec(mec)
return mec
}

/**
 * return truth of "has merged cells"
 *
 * @return
 */
     fun hasMergedCells():Boolean {
return mc.size > 0
}

/**
 * get the name of the sheet
 */
    override fun toString():String {
return sheetName
}

/**
 * initialize the SheetImpl with data from
 * the byte array.
 */
    override fun init() {
super.init()
val lt = ByteTools.readInt(this.getByteAt(0), this.getByteAt(1), this
.getByteAt(2), this.getByteAt(3))

 // this is the index used by the BOF's Sheet to associate
        // the record
        lbPlyPos = lt.toLong()
sheetType = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
if (DEBUGLEVEL > 9)
{
Logger.logInfo("Sheet grbit: $sheetType")
Logger.logInfo(" lbplypos: $lbPlyPos")
}
cch = this.getByteAt(6)
grbitChr = this.getByteAt(7)
val namebytes = this.getBytesAt(8, this.length - 12)
try
{
if (grbitChr.toInt() == 0x1)
{
sheetname = String(namebytes!!,
WorkBookFactory.UNICODEENCODING)
}
else
{
sheetname = String(namebytes!!,
WorkBookFactory.DEFAULTENCODING)
}
}
catch (e:UnsupportedEncodingException) {
Logger.logInfo("Boundsheet.init() Unsupported Encoding error: $e")
}

if (DEBUGLEVEL > 9)
Logger.logInfo("Sheet name: $sheetname")
ooxmlObjects = ArrayList() // possible that boundsheet is created
 // by readObject and therefore
        // ooxmlObjects will not be set
    }

/**
 * prior to serializing the worksheet,
 * we need to initialize the records which belong to this sheet
 * instance.
 */
    override fun setLocalRecs() {
localrecs = CompatibleVector()

val newSheetRecs = this.assembleSheetRecs()

val shtr = newSheetRecs.iterator()
while (shtr.hasNext())
{
try
{
val x = shtr.next() as XLSRecord
x.getData()
if (x is Labelsst)
{ // put the String in the label
x.initUnsharedString()
}
localrecs!!.add(x)
}
catch (e:Exception) {
Logger.logWarn("Setting Boundsheet records problem: $e")
}

}
 // add the charts to the boundsheet, as they are stored in
        // the workbook normally. (why?)
        charts.clear()
val chts = this.workBook!!.charts
for (i in chts.indices)
{
if (chts[i].sheet == this)
{
charts.add(chts[i])
}
}
}

/**
 * This seems incorrect, should we not be just returning
 * charts for the boundsheet in question?
 *
 * @return
 */
     fun getCharts():List<*> {
return charts
}

/**
 * add chart to sheet-specific list of charts
 *
 * @param c
 * @see WorkBook.addRecord
 */
     fun addChart(c:Chart) {
charts.add(c)
}

/**
 * Adds a chart from a bytestream, keeps the title intact.
 *
 * @param inbytes
 * @return
 */
     fun addChart(inbytes:ByteArray, coords:ShortArray):Chart? {
return this.addChart(inbytes, "useDefault", coords)
}

 /*
     * Inserts a serialized boundsheet into the workbook, and
     * changes the name.
     */
    override fun addChart(inbytes:ByteArray, NewChartName:String, coords:ShortArray):Chart? {
var destChart:Chart? = null
 // Deserialize bytes
        try
{
val bais = ByteArrayInputStream(inbytes)
val bufstr = BufferedInputStream(bais)
val o = ObjectInputStream(bufstr)
destChart = o.readObject() as Chart
}
catch (e:Exception) {
Logger.logInfo("Boundsheet.addChart() failed:$e")
}

if (destChart != null)
{ // got chart
if (NewChartName != "useDefault")
destChart.title = NewChartName // set new name
 // why do we need this??? shouldn't it be already set??
            // destChart.getChartFormat().setParentChart(destChart); //
            // make same as WorkBook.addChart
            destChart.setSheet(this)
 // BUGTRACKER 2372: chart bounds are dependent upon row+col
            // sizes so use coordinates (which are row/col independent)
            // does it makes sense to only set h + w and NOT x and y
            val origCoords = destChart.coords
coords[0] = origCoords[0] // don't set X and Y (keep to original
 // row and column
            coords[1] = origCoords[1]
destChart.coords = coords // but set w + h
destChart.id = this.lastObjId + 1 // 20100210 KSC: track last obj
 // id per sheet ...
            var localFonts:HashMap<*, *>? = null // fonts currently in workbook
if (this.getTransferFonts() != null && this.getTransferFonts()!!.size > 0)
{ // then must
 // translate old
                // font indexes
                // to new font
                // indexes
                localFonts = this.workBook!!.fontRecsAsXML as HashMap<*, *> // fonts
 // in
                // this
                // workbook
            }
val recs = destChart.xlSrecs
for (i in recs.indices)
{
val rec = recs[i] as XLSRecord
rec.workBook = wkbook
rec.setSheet(this)
if (rec.opcode == XLSConstants.MSODRAWING)
{
wkbook!!.addChartUpdateMsodg(rec as MSODrawing, this)
continue
}
if (rec !is Bof)
 // TODO: error/problem with the BOF
                    // record!!!
                    rec.init()
if (rec is Dimensions)
destChart.setDimensions(rec)
if (rec is FontBasis)
{ // 20090506 KSC: fontbasis font
 // indexes link to subsequent
                    // text displays [added for
                    // BUGTRACKER 2372]
                    var fid = rec.fontIndex
 // see if must translate old font indexes to new font
                    // indexes
                    fid = translateFontIndex(fid, localFonts)
rec.fontIndex = fid
}
if (rec is Fontx)
{ // 20080911 KSC: must handle out of
 // bounds font references upon chart
                    // copies [JPM BugTracker 1434]
                    var fid = rec.ifnt
if (fid > 0)
fid = translateFontIndex(fid, localFonts)
rec.ifnt = fid
}
try
{
(rec as GenericChartObject).parentChart = destChart
}
catch (e:ClassCastException) { // Scl, Obj and others are not
 // chart objects
                }

 // try{Logger.logInfo("Boundsheet Added new Chart rec:" +
                // rec);}catch(Exception
                // e){Logger.logWarn("Boundsheet.addChart() could not get
                // String for rec: "+ rec.getCellAddress());}
            }
this.wkbook!!.chartVect.add(destChart)
}
charts.add(destChart)
return destChart
}

/**
 * @param fid
 * @return
 */
    internal fun translateFontIndex(fid:Int, localFonts:HashMap<*, *>?):Int {
var fid = fid
if (transferFonts != null && fid - 1 < transferFonts!!.size)
{
 // must translate fid to corrent font index for current
            // fonts
            // translate font style and see if already present
            val thisFont = transferFonts!![fid - 1] as Font
val xmlFont = "<FONT><" + thisFont.xml + "/></FONT>"
val fontNum = localFonts!!.get(xmlFont)
if (fontNum != null)
{ // then get the fontnum in this book
fid = (fontNum as Int).toInt()
}
else
{ // it's a new font for this workbook, add it in
fid = this.workBook!!.insertFont(thisFont) + 1
localFonts[xmlFont] = Integer.valueOf(fid)
}
}
if (fid > this.workBook!!.numFonts)
{ // if fid is still
 // incorrect, set to 0
            fid = 0
}
return fid
}

/**
 * populateForTransfer is a method that takes all of the shared resources (SST, XF, Font, etc) records and
 * verifies that they are populated for use in a destination workbook
 */
     fun populateForTransfer() {
this.getSheetHash()
val recs = this.cells
for (i in recs.indices)
{
if (recs[i].opcode == XLSConstants.LABELSST)
{
val mylabel = recs[i] as Labelsst
mylabel.initUnsharedString()
}
}
transferXfs = this.workBook!!.xfrecs
for (i in transferXfs.indices)
{
val x = transferXfs[i] as Xf
x.populateForTransfer()
}

transferFonts = this.workBook!!.fontRecs
for (i in transferFonts!!.indices)
{
val x = transferFonts!![i] as Font
x.getData()
}
}

/**
 * return local XF records, used for boundsheet transferral.
 */
     fun getTransferXfs():List<*> {
return transferXfs
}

/**
 * return local Font records, used for boundsheet transferral.
 */
     fun getTransferFonts():List<*>? {
return transferFonts
}

/**
 * @param localrecs The localrecs to set.
 */
     fun setLocalRecs(l:FastAddVector) {
this.localrecs = l
}

 fun setSheetRecs(shtRecs:AbstractList<*>) {
this.SheetRecs = shtRecs
}

/**
 * Set to true to turn off checking for existing cells, conditional formats and merged ranges in order to
 * accelerate adding new cells
 *
 * @param fastCellAdds The fastCellAdds to set.
 */
     fun setFastCellAdds(fastCellAdds:Boolean) {
this.fastCellAdds = fastCellAdds
}

/**
 * scl is for zoom
 *
 * @return
 */
     fun getScl():Scl {
if (scl == null)
{ // we needs one!
scl = Scl()
this.SheetRecs.add(this.indexOfWindow2, scl)
scl!!.setSheet(this)
}
return scl
}

/**
 * scl is for zoom
 *
 * @param scl
 */
     fun setScl(s:Scl) {
this.scl = s
}

/**
 * Set whether to shift formula cells inclusively
 * i.e. if inserting row 5, shift formula D5:D6 down or D6:D7 OR shift inclusive by D5:D7
 */
     fun setShiftRule(bShiftInclusive:Boolean) {
isShiftInclusive = bShiftInclusive
}

/**
 * Gets a list of the print related records for this sheet.
 *
 * @return an unmodifiable list of all printing-related records
 */
     fun getPrintRecs():List<*> {
return Collections.unmodifiableList<Any>(printRecs!!)
}

/**
 * Adds a print-related record to the list of said.
 */
     fun addPrintRec(record:BiffRec) {
if (printRecs == null)
printRecs = ArrayList()
printRecs!!.add(record)
}

/**
 * return the AutoFilter record for this Boundsheet, if any
 * TODO: Merge with OOXML Autofilter
 *
 * @return
 */
     fun getAutoFilters():List<*> {
return autoFilters
}

/**
 * Adds a new Note or Comment to the sheet at the desired address
 *
 * @param address - String cell address
 * @param txt     - Text of Note
 * @param author  - String Author of Note
 * @return NoteHandle - a handle to the Note object which allows manipulation
 */
     fun createNote(address:String, txt:String?, author:String):Note {
var address = address
 // first check if a note is already attached to this addrss
        val notes = this.notes
if (address.indexOf('!') == -1)
address = this.sheetName + "!" + address
for (i in notes.indices)
{
val n = notes[i] as Note
if (n.cellAddressWithSheet == address)
{
n.text = txt
n.author = author
return n
}
}

 // add required Mso/object records
        val coords = ExcelTools.getRowColFromString(address)
var insertIndex = this.insertMSOObjectsForNote(coords)

 // after mso/obj/mso, now add txo/continue/continue and note
        // record
        val t = Txo.prototype as Txo?
t!!.setSheet(this)
this.SheetRecs.add(insertIndex++, t)
this.SheetRecs.add(insertIndex++, t.text) // add the associated
 // Continues that defines
        // the text
        t.text!!.predecessor = t // link this continues to it's predecessor
val c = Continue.basicFormattingRunContinues
c.predecessor = t // TODO: is this correct????
c.setSheet(this)
this.SheetRecs.add(insertIndex++, c) // and add associated formatting
 // runs
        try
{
t.stringVal = txt // must do *after* adding continues
}
catch (e:IllegalArgumentException) {
Logger.logErr(e.toString())
}

 // after (mso/obj/mso/txo/continue/continue) * n, note
        // records are listed in order
        insertIndex = this.getIndexOf(XLSConstants.WINDOW2)
val n = Note.getPrototype(author) as Note
n.id = this.lastObjId // same as Obj record above
n.setSheet(this)
n.setRowCol(coords[0], coords[1])
this.SheetRecs.add(insertIndex, n)
return n
}

/**
 * Handles the MSO manipulations necessary for creating a note record
 *
 *
 * // For each note:
 * // [msodrawing
 * //  obj - ftNts note
 * //  msodrawing - attached shape
 * //  Txo (text object), continue, continue] x n
 * // [note 1] x n
 * // window 2
 * // ************************************************************************************
 * // NOTE:
 * // SOME TEMPLATES HAVE [obj= ftNts, Continue, Txo, continue, continue, continue]
 * // MORE INFO (get this):
 * // Obj, Continue= 2nd MSO!!!, Txo, Continue, Continue, Continue= 1st MSO!!!
 * // ************************************************************************************
 *
 * @param coords rowcol of the note record
 * @return insertion index for note
 */
    private fun insertMSOObjectsForNote(coords:IntArray):Int {
var insertIndex:Int
var msodg = this.wkbook!!.msoDrawingGroup
if (msodg == null)
{
msodg = this.wkbook!!.createMSODrawingGroup()
msodg!!.initNewMSODrawingGroup()
}

 // insert either above first NOTE record or before WINDOW2
        // and certain other XLSRECORDS
        insertIndex = this.getIndexOf(XLSConstants.NOTE)
if (insertIndex == -1)
 // no existing notes - find proper insert index
            insertIndex = this.getIndexOf(XLSConstants.WINDOW2)
while (insertIndex - 1 > 0)
{
val opc = (SheetRecs[insertIndex - 1] as BiffRec).opcode
if (opc == XLSConstants.MSODRAWING || opc == XLSConstants.CONTINUE)
{
val rec:MSODrawing?
if (opc == XLSConstants.MSODRAWING)
rec = SheetRecs[insertIndex - 1] as MSODrawing
else
{
rec = (SheetRecs[insertIndex - 1] as Continue).maskedMso
if (rec == null)
break
}
if (rec.solverContainerLength == 0)
break // solver containers must be last, apparently ...
 // sigh ...
                // else
                // Logger.logInfo("Boundsheet.InsertMSOObjectsForNote.
                // SOLVER CONTAINER ENCOUNTED");
            }
else if (opc == XLSConstants.OBJ || opc == XLSConstants.CONTINUE || opc == XLSConstants.DIMENSIONS
|| opc.toInt() == 0x866 || opc.toInt() == 0x1C2)
break
insertIndex--
}

var msoheader = msodg.getMsoHeaderRec(this)
var msoDrawing:MSODrawing = MSODrawing.prototype as MSODrawing?
msoDrawing.setSheet(this)
msoDrawing.workBook = this.workBook
if (msoheader == null)
{
msoDrawing.setIsHeader()
msoheader = msoDrawing
}

 // mso record which creates a text box
        msoDrawing.createCommentBox(coords[0], coords[1])
this.SheetRecs.add(insertIndex++, msoDrawing)
msoheader.numShapes++
msodg.addMsodrawingrec(msoDrawing) // add the new drawing rec to the
 // msodrawinggroup set of recs

        // object record which defines a basic note
        val obj = Obj.getBasicObjRecord(Obj.otNote.toInt(), ++this.lastObjId) // create
 // a
        // note
        // object
        this.SheetRecs.add(insertIndex++, obj)

 // now add attached text-type mso, specifying the shape has
        // attached text
        msoDrawing = MSODrawing.textBoxPrototype as MSODrawing
msoDrawing.setSheet(this)
this.SheetRecs.add(insertIndex++, msoDrawing)
msodg.addMsodrawingrec(msoDrawing) // add the new drawing rec to the
 // msodrawinggroup set of recs

        // now update msodg + msoheader rec
        wkbook!!.updateMsodrawingHeaderRec(this) // find the msodrawing header
 // record and update it (using
        // info from other msodrawing
        // recs)
        msodg.spidMax = this.wkbook!!.lastSPID + 1
msodg.updateRecord() // given all information, generate appropriate
 // bytes for the Mso rec
        msodg.dirtyflag = true
return insertIndex
}

/**
 * Handles the MSO records necessary for defining a DropDown list object
 *
 * @return
 */
     fun insertDropDownBox(colNum:Int):Int {
var insertIndex:Int
var msodg = this.wkbook!!.msoDrawingGroup
if (msodg == null)
{
msodg = this.wkbook!!.createMSODrawingGroup()
msodg!!.initNewMSODrawingGroup()
}

 // insert either above first NOTE record or before WINDOW2
        // and certain other XLSRECORDS
        insertIndex = this.getIndexOf(XLSConstants.NOTE)
if (insertIndex == -1)
 // no existing notes - find proper insert index
            insertIndex = this.getIndexOf(XLSConstants.WINDOW2)
while (insertIndex - 1 > 0)
{
val opc = (SheetRecs[insertIndex - 1] as BiffRec).opcode
if (opc == XLSConstants.MSODRAWING || opc == XLSConstants.CONTINUE)
{
val rec:MSODrawing?
if (opc == XLSConstants.MSODRAWING)
{
rec = SheetRecs[insertIndex - 1] as MSODrawing
}
else
{
rec = (SheetRecs[insertIndex - 1] as Continue).maskedMso
if (rec == null)
break
}
if (rec.solverContainerLength == 0)
break // solver containers must be last, apparently ...
 // sigh ...
                // else
                // Logger.logInfo("Boundsheet.InsertMSOObjectsForNote.
                // SOLVER CONTAINER ENCOUNTED");
            }
else if (opc == XLSConstants.OBJ)
{
val rec = SheetRecs[insertIndex - 1] as Obj
if (rec.objType == Obj.otDropdownlist.toInt())
 // TODO: verify that
                    // drop downs are
                    // reused/shared in
                    // all cases!!!!
                    return rec.objId // already have one return object id
break
}
else if (opc == XLSConstants.OBJ || opc == XLSConstants.CONTINUE || opc == XLSConstants.DIMENSIONS
|| opc.toInt() == 0x866 || opc.toInt() == 0x1C2)
{
break
}
insertIndex--
}

var msoheader = msodg.getMsoHeaderRec(this)
val msoDrawing = MSODrawing.prototype as MSODrawing?
msoDrawing!!.setSheet(this)
msoDrawing.workBook = this.workBook
if (msoheader == null)
{
msoDrawing.setIsHeader()
msoheader = msoDrawing
}

msoDrawing.createDropDownListStyle(colNum) // create the records
 // necessary to define the
        // dropdown box symbol at
        // the desired column

        // object record which defines a basic dropdown list
        val obj = Obj.getBasicObjRecord(Obj.otDropdownlist.toInt(), ++this.lastObjId) // create
 // a
        // drop-down
        // object
        // record
        // for
        // each
        val objID = obj.objId

 // insert new mso + obj records into sheet
        this.SheetRecs.add(insertIndex++, msoDrawing)
this.SheetRecs.add(insertIndex++, obj)

 // now update msodg + msoheader rec
        msoheader.numShapes++
msodg.addMsodrawingrec(msoDrawing) // add the new drawing rec to the
 // msodrawinggroup set of recs
        wkbook!!.updateMsodrawingHeaderRec(this) // find the msodrawing header
 // record and update it (using
        // info from other msodrawing
        // recs)
        msodg.spidMax = this.wkbook!!.lastSPID + 1
msodg.updateRecord() // given all information, generate appropriate
 // bytes
        msodg.dirtyflag = true

return objID
}

/**
 * Adds a new Note or Comment to the sheet at the desired address
 * with Formatting (Font) information
 *
 * @param address - String cell address
 * @param txt     - Unicode string reprentation of the note, including formatting
 * @param author  - String Author of Note
 * @return NoteHandle - a handle to the Note object which allows manipulation
 */
     fun createNote(address:String, txt:Unicodestring, author:String):Note {
val nh = this.createNote(address, txt.stringVal, author)
 // TODO: deal with formats - incorporate into Txo/Continues
        // -- for now they are just stored as-is, no modification
        // allowed
        nh.formattingRuns = txt.formattingRuns
return nh
}

/**
 * removes the desired note from the sheet
 *
 * @param n
 */
     fun removeNote(n:Note) {
val id = n.id
var idx = this.getIndexOf(XLSConstants.OBJ)
if (idx == -1)
return  // should't!
while (idx < this.SheetRecs.size)
{
if ((this.SheetRecs[idx] as BiffRec).opcode == XLSConstants.OBJ)
{
val o = this.SheetRecs[idx] as Obj
 // if it's of type Note + has the same id, this is it
                if (o.objType == 0x19 && o.objId == id)
{ // got it!
 // apparently sometimes you don't find the mso/obj/mso
                    // combo, so check
                    if ((this.SheetRecs[idx - 1] as BiffRec)
.opcode == XLSConstants.MSODRAWING)
{
idx--
break
}
else if ((this.SheetRecs[idx + 1] as BiffRec)
.opcode == XLSConstants.CONTINUE && (this.SheetRecs[idx + 1] as Continue).maskedMso != null)
{
 // idx++;
                        break
}
}
}
idx++
}
 // usual format= mso/obj/mso/txo/continue/continue but can
        // also be:
        // obj/continue (mso)/txo/continue/continue/continue (mso)
        var objidx = 0
var msoidx = 0
var maskedMso = true // handle continues masking mso's
while (idx < this.SheetRecs.size)
{
var mso:MSODrawing? = null
val rec = this.SheetRecs[idx] as BiffRec
if (rec.opcode == XLSConstants.OBJ)
objidx++
else if (rec.opcode == XLSConstants.MSODRAWING)
{
mso = rec as MSODrawing
maskedMso = false
if (mso.getShapeType() == MSODrawingConstants.msosptTextBox)
msoidx++
else if (rec.isShape)
 // it's not a text box or
                    // the associated text
                    // "oddball" mso, so
                    // break (Another test:
                    // SPID==0??)
                    break
}
else if (rec.opcode == XLSConstants.CONTINUE && maskedMso)
{
mso = (rec as Continue).maskedMso
if (mso!!.getShapeType() == MSODrawingConstants.msosptTextBox)
msoidx++
else if ((rec as MSODrawing).isShape)
 // it's not a text box or
                    // the associated text
                    // "oddball" mso, so
                    // break (Another test:
                    // SPID==0??)
                    break
}
else if (rec.opcode == XLSConstants.NOTE)
break
if (objidx > 1 || msoidx > 1)
 // reached the next set of
                // note-associated recs, so get out
                break
this.SheetRecs.removeAt(idx) // otherwise, ok to delete
if (mso != null && mso.isShape)
{// if removed an mso, must update
 // msodg
                val msodg = this.wkbook!!.msoDrawingGroup
msodg!!.removeMsodrawingrec(mso, this, true)
}
}
 // now remove the actual note record
        idx = this.getIndexOf(XLSConstants.NOTE)
while (idx < this.SheetRecs.size && (this.SheetRecs[idx] as BiffRec).opcode == XLSConstants.NOTE)
{
if (this.SheetRecs[idx] == n)
{
this.SheetRecs.removeAt(idx)
break // we're done
}
idx++
}
}

/**
 * Adds a new AutoFilter to the specified column
 *
 * @param int column - 0-based column number
 * @return AutoFilterHandle Handle to the new AutoFilter
 */
     fun addAutoFilter(column:Int):AutoFilter {
 // if there are no existing AutoFilters on the sheet,
        // then must add a mso/obj pair for each column on the sheet
        // to define dropdown box next to each column
        // also must add built-in name _FILTERDATABASE +
        // must add a mystery XlSRecord with opcode== 0x9D -- cannot
        // find any information about this opcode
        if (this.autoFilters == null || this.autoFilters.size == 0)
{
 // add _FILTERDATABASE Name
            this.addFilterDatabase()

 /*
             * 20100216 KSC: WHAT ARE THESE RECORD??? They are necessary
             * for new AutoFilter's
             */
            var zz = getIndexOf(XLSConstants.COLINFO)
if (zz == -1)
zz = getIndexOf(XLSConstants.DEFCOLWIDTH) + 1
else
{
while ((this.SheetRecs[zz] as BiffRec)
.opcode == XLSConstants.COLINFO)
zz++
}
 // insert after COLINFOs or DefColWidth
            var rec = XLSRecord()
rec.opcode = 155.toShort() // no data for this record
rec.setData(byteArrayOf())
this.SheetRecs.add(zz++, rec)
rec = XLSRecord()
rec.opcode = 157.toShort() // this has SOMETHING to do with #
 // columns ...
            rec.setData(byteArrayOf(this.maxCol.toByte(), 0))
this.SheetRecs.add(zz, rec)

 // add required Mso/object records
            var insertIndex:Int
var msodg = this.wkbook!!.msoDrawingGroup
if (msodg != null)
{ // already have drawing records; just add to
 // records + update msodg
                insertIndex = this.getIndexOf(XLSConstants.MSODRAWINGSELECTION)
if (insertIndex < 0)
insertIndex = this.getIndexOf(XLSConstants.WINDOW2)
}
else
{ // No images present in workbook, must add appropriate
 // records
                msodg = this.wkbook!!.createMSODrawingGroup()
msodg!!.initNewMSODrawingGroup() // generate and add required
 // records for drawing records
                // insertion point for new msodrawing rec
                insertIndex = getIndexOf(XLSConstants.DIMENSIONS) + 1
}

var msoheader = msodg.getMsoHeaderRec(this)

 // Must add for each column
            for (i in 0 until this.realMaxCol)
{
try
{
if (this.getCellsByCol(i).size == 0)
break
}
catch (e:CellNotFoundException) {
break
}

 // Colinfo ci= (Colinfo) this.colinfos.get(i);
                // short j= (short) ci.getColFirst(); // column number'
                val j = i.toShort()
val msoDrawing = MSODrawing.prototype as MSODrawing?
msoDrawing!!.workBook = this.wkbook
msoDrawing.setSheet(this)
if (msoheader == null)
{
msoDrawing.setIsHeader()
msoheader = msoDrawing
}

msoDrawing.createDropDownListStyle(j.toInt()) // create the records
 // necessary to define
                // the dropdown box
                // symbol

                // object record which defines a basic dropdown list
                val obj = Obj
.getBasicObjRecord(Obj.otDropdownlist.toInt(), ++this.lastObjId) // create
 // a
                // drop-down
                // object
                // record
                // for
                // each

                // insert new mso + obj records into sheet
                this.SheetRecs.add(insertIndex++, msoDrawing)
this.SheetRecs.add(insertIndex++, obj)

 // now update msodg + msoheader rec
                msoheader.numShapes++
msodg.addMsodrawingrec(msoDrawing) // add the new drawing rec
 // to the msodrawinggroup
                // set of recs
                wkbook!!.updateMsodrawingHeaderRec(this) // find the msodrawing
 // header record and
                // update it (using info
                // from other msodrawing
                // recs)
                msodg.spidMax = this.wkbook!!.lastSPID + 1
msodg.updateRecord() // given all information, generate
 // appropriate bytes
                msodg.dirtyflag = true
}
}

val af = AutoFilter.prototype as AutoFilter?
af!!.setSheet(this)
af.col = column
val i = getIndexOf(XLSConstants.DIMENSIONS) // insert just before DIMENSIONS record
this.SheetRecs.add(i, af)

this.autoFilters.add(af)
return af
}

/**
 * removes all autofilters from this sheet
 */
     fun removeAutoFilter() {
this.removeFilterDatabase() // remove the _FILTER_DATABASE name
 // necessary for AutoFilters
        var zz = getIndexOf(XLSConstants.AUTOFILTER) // remove all AutoFitler records
while (zz != -1)
{
this.SheetRecs.removeAt(zz)
zz = getIndexOf(XLSConstants.AUTOFILTER)
}
 // remove the two unknown records
        zz = getIndexOf(155.toShort())
if (zz > -1)
this.SheetRecs.removeAt(zz)
zz = getIndexOf(157.toShort())
if (zz > -1)
this.SheetRecs.removeAt(zz)
 // and hows about the Mso/Obj records, huh? huh?
        this.autoFilters.clear()
 // finally, must set all rows to NOT hidden - I believe
        // Excel does this when AutoFilters are turned off
        for (i in 0 until rows.size)
rows[Integer.valueOf(i)].isHidden = false
}

/**
 * adds a sxview - pivot table lead record - and required associated records to the worksheet
 * <br></br>other methods that add data, row, col and page fields will fill in the pivot table fields and formatting info
 *
 * @param ref Cell Range which identifies pivot table data range
 * @param wbh WorkBookHandle
 * @param sId Stream or cachid Id -- links back to SxStream set of records
 * @return
 */
     fun addPivotTable(ref:String, wbh:WorkBookHandle, sId:Int, tablename:String):Sxview {
this.wkbook!!.addPivotCache(ref, wbh, sId) // create the
 // directory/storage for a
        // pivot cache, if not
        // already created
        // ensure the proper directory/storage and pivot cache
        // record is created
        var zz = window2!!.recordIndex - 1
while (zz > 0)
{
if ((this.SheetRecs[zz] as BiffRec).opcode == XLSConstants.NOTE)
break
if ((this.SheetRecs[zz] as BiffRec).opcode == XLSConstants.OBJ)
break
if ((this.SheetRecs[zz] as BiffRec).opcode == XLSConstants.DIMENSIONS)
break
zz--
}
zz++
 // minimal configuration
        val sx = Sxview.prototype as Sxview?
this.SheetRecs.add(zz++, sx)
this.SheetRecs.addAll(zz, sx!!.addInitialRecords(this))
sx.tableName = tablename
this.wkbook!!.addPivotTable(sx) // add to lookup
return sx
}

/**
 * update row filter (hidden status) by evaluating AutoFilter conditions on the sheet
 * <br></br>Must do after autofilter updates or additions
 */
     fun evaluateAutoFilters() {
 // first must set all rows to NOT hidden
        for (i in 0 until rows.size)
try
{
rows[Integer.valueOf(i)].isHidden = false
}
catch (e:NullPointerException) {
 // blank rows ...
            }

 // now evaluate all autofilters
        for (i in autoFilters.indices)
{
(this.autoFilters.get(i) as AutoFilter).evaluate()
}
}

/**
 * returns the list of Excel 2007 objects which are external or auxillary to this sheet
 * e.g printerSeettings, vmlDrawings
 *
 * @return
 */
     fun getOOXMLObjects():List<*> {
return ooxmlObjects
}

/**
 * adds the object-specific signature of the external or auxillary Excel 2007 object
 * e.g. oleObjects, vmlDrawings
 *
 * @param o
 */
     fun addOOXMLObject(o:Any) {
ooxmlObjects.add(o)
}

 // TODO: Handle below options in Excel 2003 i.e. create
    // appropriate records
    // *************************************************************

    /**
 * set if row has thick bottom by default (Excel 2007-Specific)
 */
     fun hasThickBottom():Boolean {
return thickBottom
}

/**
 * return true if row has thick top by default (Excel 2007-Specific)
 */
     fun hasThickTop():Boolean {
return thickTop
}

/**
 * return true if rows are hidden by default (Excel 2007-Specific)
 */
     fun hasZeroHeight():Boolean {
return zeroHeight
}

/**
 * return true if defaultrowheight is manually set (Excel 2007-Specific)
 */
     fun hasCustomHeight():Boolean {
return customHeight
}

/**
 * set if row has thick bottom by default (Excel 2007-Specific)
 */
     fun setThickBottom(b:Boolean) {
thickBottom = b
}

/**
 * set if row has thick top by default (Excel 2007-Specific)
 */
     fun setThickTop(b:Boolean) {
thickTop = b
}

/**
 * set if rows are hidden by default (Excel 2007-Specific)
 */
     fun setZeroHeight(b:Boolean) {
zeroHeight = b
}

/**
 * set if defaultrowheight is manually set (Excel 2007-Specific)
 */
     fun setHasCustomHeight(b:Boolean) {
customHeight = b
}

/**
 * store Excel 2007 shape via Shape Name
 *
 * @param tca
 */
     fun addOOXMLShape(tca:io.starter.formats.OOXML.TwoCellAnchor) {
if (ooxmlShapes == null)
ooxmlShapes = HashMap()
ooxmlShapes!![tca.name] = tca
}

/**
 * store Excel 2007 shape via Shape Name
 *
 * @param tca
 */
     fun addOOXMLShape(oca:io.starter.formats.OOXML.OneCellAnchor) {
if (ooxmlShapes == null)
ooxmlShapes = HashMap()
ooxmlShapes!![oca.name] = oca
}

/**
 * Store Excel 2007 legacy drawing shapes
 *
 * @param vml
 */
     fun addOOXMLShape(vml:Any) {
if (ooxmlShapes == null)
ooxmlShapes = HashMap()
ooxmlShapes!!["vml"] = vml // only 1 vml (=legacy drawing info) per
 // sheet so just refer to it as "vml"
    }

/**
 * returns a scoped named range by name string
 *
 * @param t
 * @return
 */
     fun getScopedName(nameRef:String):Name? {
val o = this.sheetNameRecs!![nameRef.toUpperCase()] ?: return null // case
 // insensitive
        return o as Name
}

/**
 * Add a sheet-scoped name record to the boundsheet
 *
 *
 * Note this is not that primary repository for names, it just contains the name records
 * that are bound to this sheet, adding them here will not add them to the workbook;
 *
 * @param sheetNameRecs
 */
     fun addLocalName(name:Name) {
if (sheetNameRecs == null)
sheetNameRecs = HashMap()
sheetNameRecs!![name.nameA] = name
}

/**
 * Remove a sheet-scoped name record from the boundsheet.
 *
 *
 * Note this is not that primary repository for names, it just contains the name records
 * that are bound to this sheet, removing them here will not remove them completely from the workbook.
 *
 *
 * In order to do that you will need to call book.removeName
 *
 * @param sheetNameRecs
 */
     fun removeLocalName(name:Name) {
sheetNameRecs!!.remove(name.nameA)
}

/**
 * Get a sheet scoped name record from the boundsheet
 *
 * @return
 */
     fun getName(name:String):Name? {
if (sheetNameRecs == null)
return null
val o = sheetNameRecs!![name.toUpperCase()] // case insensitive
return if (o != null) o as Name? else null
}

/**
 * add pritner setting record to worksheet recs
 *
 * @param r printer setting record (Margins, PLS)
 */
     fun addMarginRecord(r:BiffRec) {
r.setSheet(this)
var i = this.getIndexOf(XLSConstants.SETUP)
val thisOpCode = r.opcode.toInt()
 // iterate up from SETUP record
        // desired order:
        // WsBool, HeaderRec, FooterRec, HCenter, VCenter,
        // LeftMargin, RightMargin, TopMargin, BottomMargin, Pls
        while (i > 0)
{
val prevOpCode = (this.sheetRecs[--i] as BiffRec)
.opcode.toInt()
if (prevOpCode == XLSConstants.VCENTER.toInt()
|| /* assume AT LEAST a VCENTER or FOOTERREC */
prevOpCode == XLSConstants.FOOTERREC.toInt() || prevOpCode == XLSConstants.LEFT_MARGIN.toInt())
{
break
}
if ((prevOpCode == XLSConstants.BOTTOM_MARGIN.toInt() || prevOpCode == XLSConstants.TOP_MARGIN.toInt()
|| prevOpCode == XLSConstants.RIGHT_MARGIN.toInt()) && thisOpCode == XLSConstants.PLS.toInt())
break
if ((prevOpCode == XLSConstants.TOP_MARGIN.toInt() || prevOpCode == XLSConstants.RIGHT_MARGIN.toInt()) && thisOpCode == XLSConstants.BOTTOM_MARGIN.toInt())
break
if (prevOpCode == XLSConstants.RIGHT_MARGIN.toInt() && thisOpCode == XLSConstants.TOP_MARGIN.toInt())
break
}
this.SheetRecs.add(++i, r)
}

/**
 * inserts a row and shifts all of the other rows down one
 *
 *
 * the rownum is zero based.  calling insertrow(9,true) will
 * create a row containing A10, and subsequently shift rows > 9 by 1.
 *
 * @return the row that was just inserted
 */
    private fun insertRow(rownum:Int, shiftrows:Boolean):Row? {
return insertRow(rownum, 0, WorkSheetHandle.ROW_INSERT_MULTI, shiftrows)
}

@Throws(XmlPullParserException::class, IOException::class)
internal fun parseOOXML(bk:WorkBookHandle, sheet:WorkSheetHandle, ii:InputStream, sst:ArrayList<*>, formulas:ArrayList<*>, hyperlinks:ArrayList<*>, inlineStrs:HashMap<*, *>?) {
var inlineStrs = inlineStrs
val sfindex = formulas.size

 // try {
        var r:Row? = null
var cellAddr:String? = null
var formatId = 0
var type = ""
shExternalLinkInfo = HashMap()

val factory = XmlPullParserFactory.newInstance()
factory.isNamespaceAware = true
val xpp = factory.newPullParser()

xpp.setInput(ii, null) // using XML 1.0 specification
var eventType = xpp.eventType
while (eventType != XmlPullParser.END_DOCUMENT)
{
if (eventType == XmlPullParser.START_TAG)
{
val tnm = xpp.name
if (tnm == "sheetFormatPr")
{ // baseColWidth, customHeight
 // (true if defaultRowHeight
                    // has been manually set),
                    // defaultColWidth,
                    // defaultRowHeight -
                    // optimiztion so that we
                    // don't have to write out
                    // values on each
                    // thickBottom - true if
                    // rows have a thick bottom
                    // by default
                    // thickTop - true if rows
                    // have a thick top by
                    // default
                    // zeroHeight - true if rows
                    // are hidden by default (an
                    // optimization)
                    for (i in 0 until xpp.attributeCount)
{
val n = xpp.getAttributeName(i)
val v = xpp.getAttributeValue(i)
if (n == "thickBottom")
this.setThickBottom(true)
else if (n == "thickTop")
this.setThickTop(true)
else if (n == "zeroHeight")
this.setZeroHeight(v == "1")
else if (n == "customHeight")
this.setHasCustomHeight(v == "1")
else if (n == "defaultColWidth")
this.defaultColumnWidth = Float(v)
else if (n == "defaultRowHeight")
this.defaultRowHeight = Double(v)
}
}
else if (tnm == "sheetView")
{ // TODO: finish handling
 // options
                    val s = SheetView.parseOOXML(xpp)
.cloneElement() as SheetView
this.sheetView = s
this.window2!!
.showGridlines = s.getAttrS("showGridlines") != "0"
if (s.getAttr("showRowColHeaders") != null)
this.window2!!
.showSheetHeaders = s
.getAttrS("showRowColHeaders") == "1"
if (s.getAttr("showZeros") != null)
this.window2!!
.showZeroValues = s.getAttrS("showZeros") == "1"
if (s.getAttr("showOutlineSymbols") != null)
this.window2!!
.showOutlineSymbols = s
.getAttrS("showOutlineSymbols") == "1"
if (s.getAttr("tabSelected") != null)
this.setSelected(s.getAttrS("tabSelected") == "1")
if (s.getAttr("zoomScale") != null)
this.getScl()
.zoom = Double(s.getAttrS("zoomScale"))
.toFloat() / 100
}
else if (tnm == "sheetPr")
{ // sheet properties element
val sp = SheetPr.parseOOXML(xpp)
.cloneElement() as SheetPr
this.sheetPr = sp
}
else if (tnm == "dimension")
{ // ref attribute
 /*
                     * this may not reflect actual rows/cols in sheet
                     * just let our normal machinery set the sheet dimensions
                     * String ref= xpp.getAttributeValue(0);
                     * int[] rc= ExcelTools.getRangeCoords(ref);
                     * this.updateDimensions(rc[2]-1, rc[3]);
                     */
                }
else if (tnm == "sheetProtection")
{ // ref
 // attribute
                    for (i in 0 until xpp.attributeCount)
{
val nm = xpp.getAttributeName(i)
val v = xpp.getAttributeValue(i)
if (nm == "password")
this.protectionManager.setPasswordHashed(v)
else if (nm == "sheet")
this.protectionManager
.protected = OOXMLReader.parseBoolean(v)
}
}
else if (tnm == "col")
{ // min, max, width
var min = 0
var max = 0
var style = 0
var width = 0.0
var hidden = false
for (i in 0 until xpp.attributeCount)
{
val nm = xpp.getAttributeName(i)
val v = xpp.getAttributeValue(i)
if (nm == "min")
min = Integer.valueOf(v).toInt()
else if (nm == "max")
max = Integer.valueOf(v).toInt()
else if (nm == "width")
width = Double(v)
else if (nm == "hidden")
hidden = true
else if (nm == "style")
 // customFormat?
                            style = Integer.valueOf(v).toInt()
}
if (max > WorkBook.MAXCOLS)
max = WorkBook.MAXCOLS - 1
val col = sheet.addCol(min - 1, max - 1)
col.width = (width * OOXMLReader.colWFactor).toInt()
if (style > 0)
col.formatId = style
 // col.setColLast(max-1);
                    if (hidden)
col.isHidden = true
}
else if (tnm == "row")
{
var ht = -1
var ixfe = 0
var customHeight = false
for (i in 0 until xpp.attributeCount)
{ // r, v=
 // row
                        // #+1,
                        // ht,
                        // ...
                        val nm = xpp.getAttributeName(i)
val v = xpp.getAttributeValue(i)
if (nm == "r")
{
val rownum = Integer.valueOf(v).toInt() - 1
r = this.insertRow(rownum, false) // now insertRow
 // with no shift
                            // rows does NOT
                            // add a blank
                            // cell so no
                            // need to
                            // delete extra
                            // cell anymore
                            r!!.ixfe = this.workBook!!.defaultIxfe
}
else if (nm == "ht")
{
ht = (Double(v) * OOXMLReader.rowHtFactor).toInt()
}
else if (nm == "s")
{ // customFormat?
ixfe = Integer.valueOf(v).toInt()
}
else if (nm == "customFormat")
{
r!!.ixfe = ixfe
}
else if (nm == "hidden")
{
r!!.isHidden = true
}
else if (nm == "collapsed")
{ // 20090513 KSC:
 // Added
                            // collapsed,
                            // outlineLevel
                            // [BUGTRACKER
                            // 2371]
                            val h = r!!.isHidden // setCollapsed
 // unconditionally sets
                            // hidden
                            r.isCollapsed = true
if (!h)
r.isHidden = false
}
else if (nm == "outlineLevel")
{
r!!.outlineLevel = Integer.valueOf(v).toInt()
}
else if (nm == "customHeight")
{
customHeight = true
}
else if (nm == "thickBot")
{
r!!.setHasAnyThickBottomBorder(true)
}
else if (nm == "thickTop")
{
r!!.hasAnyThickTopBorder = true
}
if (ht != -1 && customHeight)
 // if customHeight is NOT
                            // set do not set row
                            // height (encountered
                            // in Baxter XLSM
                            // templates)
                            r!!.rowHeight = ht
}
 // if customheight is NOT specified do not set row height
                }
else if (tnm == "c")
{// element c child v= value
if (cellAddr != null)
{
if (r!!.explicitFormatSet || formatId != this.workBook!!
.defaultIxfe && formatId != 0)
{ // default
 // or
                            // not
                            // specified
                            // NOTE:
                            // default
                            // for
                            // OOXML
                            // is
                            // 0
                            // not
                            // 15
                            val rc = ExcelTools.getRowColFromString(cellAddr)
OOXMLReader
.sheetAdd(sheet, null, rc[0], rc[1], formatId)
}
cellAddr = null
}
formatId = 0
type = "n" // reset for those cells that don't specify a
 // type, default = number
                    for (i in 0 until xpp.attributeCount)
{
val nm = xpp.getAttributeName(i) // r, s=style, t=
 // type
                        val v = xpp.getAttributeValue(i)
if (nm == "r")
{ // cell address
cellAddr = v // save for setting later
}
else if (nm == "s")
{
formatId = Integer.valueOf(v).toInt() // save
 // for
                            // setting
                            // later
                        }
else if (nm == "t")
{
type = v
}
}
 // would be great if could peek at next tag to determine
                    // whether to add a blank cell here rather than catch it at
                    // end tag below
                }
else if (tnm == "is")
{ // inline string child of <c cell
 // element
                    if (inlineStrs == null)
inlineStrs = HashMap()
val s = OOXMLReader.getInlineString(xpp)
inlineStrs!![this.sheetName + "!" + cellAddr] = s
val rc = ExcelTools.getRowColFromString(cellAddr!!)
OOXMLReader.sheetAdd(sheet, "", rc[0], rc[1], formatId) // add
 // placeholder
                    // here
                    cellAddr = null
}
else if (tnm == "f")
{ // formula
if (cellAddr != null)
{
 // do not process now since formulas may be dependent
                        // upon other sheet data; save and process after all
                        // sheets have been added
                        var ftype = type
var ref = ""
var si = ""
var ca:String? = null
for (i in 0 until xpp.attributeCount)
{
val nm = xpp.getAttributeName(i)
val v = xpp.getAttributeValue(i)
if (nm == "t")
ftype += "/$v" // add to data type the
else if (nm == "ref")
 // only valid for master
                                // shared formula or
                                // array record
                                ref = v
else if (nm == "si")
 // shared index only valid
                                // for shared formulas
                                si = (Integer.parseInt(v) + sfindex).toString()
else if (nm == "ca")
 // calculate cell, always
                                // set for volatile
                                // functions
                                ca = "1"// formula type: shared,
 // array, datatable, normal
}
val v = OOXMLReader.getNextText(xpp)
formulas.add(arrayOf<String>(this.sheetName, cellAddr, "=" + v!!, si, ref, ftype, ca, Integer.valueOf(formatId)!!.toString(), ""))
type = "f" // cell will not be added below; rather,
 // formula cells are processed en mass in
                        // parse
                    }
}
else if (tnm == "v")
{
/**
 * Cell Value
 * handle based upon cell data type
 */
                    if (cellAddr != null)
{ // shouldn't be
val v = OOXMLAdapter.getNextText(xpp)
 // use fast add method - uses int[] location
                        val rc = ExcelTools.getRowColFromString(cellAddr)
if (type == "s")
{ // shared string
 // the SST has already been populated, now we just
                            // need to add
                            // Labelsst recs and hook up with the isst.
                            val labl = Labelsst
.getPrototype(null, bk.workBook)
labl.setIsst(Integer.valueOf(v!!).toInt())
labl.setIxfe(formatId)
this.addRecord(labl, rc)
}
else if (type == "n")
{
try
{
if (v != "null")
OOXMLReader.sheetAdd(sheet, Integer
.valueOf(v!!), rc[0], rc[1], formatId)
else
 // Should nepver get here
                                    Logger.logWarn("OOXMLAdapter.parse: Unexpected null encountered at: $cellAddr")
}
catch (n:NumberFormatException) { // could be a
 // double or
                                // float instead
                                // of an int
                                try
{
OOXMLReader.sheetAdd(sheet, Double(
v!!), rc[0], rc[1], formatId)
}
catch (nn:NumberFormatException) {
OOXMLReader.sheetAdd(sheet, Float(
v!!), rc[0], rc[1], formatId)
}

}

}
else if (type == "b")
{
val trx = v == "1" || v!!.equals("true", ignoreCase = true)
OOXMLReader.sheetAdd(sheet, java.lang.Boolean
.valueOf(trx), rc[0], rc[1], formatId)
}
else if (type == "f")
{ // grab cached value
val s = formulas[formulas.size - 1] as Array<String>
s[8] = v
formulas[formulas.size - 1] = s
}
else if (type != "e")
{ // added handling for
 // 'e' type which is a
                            // formula as well
                            // (containing an ERR
                            // cachedval)
                            OOXMLReader
.sheetAdd(sheet, v, rc[0], rc[1], formatId)
}
cellAddr = null // denote we processed this cell
}
}
else if (tnm == "mergeCell")
{
val ref = xpp.getAttributeValue(0)
try
{
val cr = CellRange(
this.sheetName + "!" + ref, bk)
cr.mergeCells(false)
}
catch (e:CellNotFoundException) { /*
                     * necessary to report
                     * error??
                     */}

}
else if (tnm == "conditionalFormatting")
{
Condfmt.parseOOXML(xpp, bk, this)
}
else if (tnm == "dataValidations")
{
Dval.parseOOXML(xpp, this)
}
else if (tnm == "autoFilter")
{ // Appears to sometimes
 // work in tandem with
                    // dataValidtions (see
                    // Modeling Workbook -
                    // WKSHT.xlsm)
                    this.ooAutoFilter = io.starter.formats.OOXML.AutoFilter
.parseOOXML(xpp) as io.starter.formats.OOXML.AutoFilter
}
else if (tnm == "hyperlink")
{
var ref = ""
var rid = ""
var desc = ""
for (i in 0 until xpp.attributeCount)
{
if (xpp.getAttributeName(i) == "ref")
ref = xpp.getAttributeValue(i)
else if (xpp.getAttributeName(i) == "id")
 // external
                            // ref
                            rid = xpp.getAttributeValue(i)
else if (xpp.getAttributeName(i) == "display")
 // display
                            // or
                            // description
                            // text
                            desc = xpp.getAttributeValue(i)
 // TODO: Also handle location, tooltip ...
                    }
hyperlinks.add(arrayOf(rid, ref, desc)) // must
 // save
                    // hyperlink
                    // refernce
                    // cell
                    // and
                    // id
                    // and
                    // link
                    // to
                    // target
                    // info
                    // in
                    // .rels
                    // file
                    // External OOXML Objects controls=embedded controls,
                    // oleObject= embedded objects
                    // These external objects contain link information which
                    // links to id's in vmlDrawingX.vml, activeX.xml ... must
                    // save and reset for later use
                }
else if (tnm == "pageSetup")
{ // scale orientation r:id
 // ...
                    addExternalInfo(shExternalLinkInfo, xpp)
}
else if (tnm == "oleObject")
{ // progId shapeId r:id ...
addExternalInfo(shExternalLinkInfo, xpp)
}
else if (tnm == "control")
{ // progId shapeId r:id ...
addExternalInfo(shExternalLinkInfo, xpp)
 // TODO: handle AlternateContent Machinery!
                    // for now, we are ignoring choice and fallback and ONLY
                    // extracting control element
                }
else if (tnm == "AlternateContent")
{ // defines a
 // mechanism for
                    // the storage
                    // of content
                    // which is not
                    // defined by
                    // this Office
                    // Open XML
                    // Standard, for
                    // example
                    // extensions
                    // developed by
                    // future
                    // software
                    // applications
                    // which
                    // leverage the
                    // Open XML
                    // formats
                    // skip, for now - may have elements
                    // Choice->
                    // control->controlPr
                    // Fallback
                    // control
                    // i.e. 1st choice is a control with control settings
                    // if not possible, fallback is
                }
else if (tnm == "Fallback")
{
OOXMLReader.getCurrentElement(xpp) // skip as can replicate
 // Choice
                }
else if (tnm == "controlPr")
{
OOXMLReader.getCurrentElement(xpp) // skip for now!!
}
else if (tnm == "extLst")
{ // skip for now!!
OOXMLReader.getCurrentElement(xpp) // skip for now!!
} /*
                 * else {
                 * if (true)
                 * Logger.logWarn("unprocessed XLSX sheet element: " + tnm);
                 * }
                 */
}
else if (eventType == XmlPullParser.END_TAG)
{
val endTag = xpp.name
if (endTag == "row" && cellAddr != null)
{

val rc = ExcelTools.getRowColFromString(cellAddr)
 // if masking an explicit row format or if it's a unique
                    // format, set to new blank cell
                    if (r!!.explicitFormatSet || formatId != this.workBook!!.defaultIxfe && formatId != 0)
{ // default or not
 // specified NOTE:
                        // default for OOXML
                        // is 0 not 15
                        // (unless converted
                        // from XLS ((:
                        // if (r.myRow.getExplicitFormatSet() || (/*formatId!=15
                        // && */formatId!=0 && uniqueFormat)) { //default or not
                        // specified NOTE: default for OOXML==0 NOT 15
                        OOXMLReader
.sheetAdd(sheet, null, rc[0], rc[1], formatId)
 // } else{
                        // sheetAdd(sheet,null,rc[0],rc[1],formatId);
                    }
cellAddr = null
}
else /**/if (endTag == "worksheet")
 // we're done!
                    break
}
eventType = xpp.next()
}
}

/**
 * parses OOXML content files given a content list cl from zip file zip
 * recurses if content file has it's own content
 * *************************************
 * NOTE: certain elements we do not as of yet process; we "pass-through" or store such elements along with any embedded objects associated with them
 * for example, activeX objects, vbaProject.bin, etc.
 * *************************************
 *
 * @param bk        WorkBookHandle
 * @param sheet     WorkSheetHandle (set if recursing)
 * @param zip       currently open ZipOutputStream
 * @param cl        ArrayList of Contents (type, filename, rId) to parse
 * @param parentDir Parent Directory for relative paths in content lists
 * @param formulas, hyperlinks, inlineStrs -- ArrayLists/Hashmaps stores sheet-specific info for later entry
 * @throws CellNotFoundException
 * @throws XmlPullParserException
 */
    @Throws(XmlPullParserException::class, CellNotFoundException::class)
 fun parseSheetElements(bk:WorkBookHandle, zip:ZipFile, cl:ArrayList<*>, parentDir:String, externalDir:String, formulas:ArrayList<*>, hyperlinks:ArrayList<*>, inlineStrs:HashMap<*, *>, pivotTables:HashMap<String, WorkSheetHandle>) {
var p:String
var target:ZipEntry?

try
{
for (i in cl.indices)
{
var c = cl[i] as Array<String>
val ooxmlElement = c[0]

 // if(DEBUG)
                // Logger.logInfo("OOXMLReader.parse: " + ooxmlElement + ":"
                // + c[1] + ":" + c[2]);

                p = StringTool.getPath(c[1])
p = OOXMLReader.parsePathForZip(p, parentDir)
if (ooxmlElement != "hyperlink")
 // if it's a hyperlink
                    // reference, don't
                    // strip path info :)
                    c[1] = StringTool.stripPath(c[1])
val f = c[1]
val rId = c[2]

if (ooxmlElement == "drawing")
{ // images, charts
 // parse drawing rels to obtain image file names and chart
                    // xml files
                    target = OOXMLReader.getEntry(zip, p + "_rels/"
+ f.substring(f.lastIndexOf("/") + 1) + ".rels")
var drawingFiles:ArrayList<*>? = null
if (target != null)
 // first retrieve enbedded content in
                        // .rels (images, charts ...)
                        drawingFiles = OOXMLReader.parseRels(OOXMLReader
.wrapInputStream(OOXMLReader.wrapInputStream(zip
.getInputStream(target)))) // obtain a
 // list of
                    // image
                    // file
                    // references
                    // for use
                    // in later
                    // parsing
                    target = OOXMLReader.getEntry(zip, p + f) // now get
 // drawingml
                    // file and
                    // process it
                    parseDrawingXML(bk, drawingFiles, OOXMLReader
.wrapInputStream(zip
.getInputStream(target!!)), zip, p, externalDir)
}
else if (ooxmlElement == "vmldrawing")
{ // legacy
 // drawing
                    // elements
                    target = OOXMLReader.getEntry(zip, p + f)
val vml = parseLegacyDrawingXML(bk, OOXMLReader
.wrapInputStream(zip.getInputStream(target!!)))
target = OOXMLReader.getEntry(zip, p + "_rels/" // get

 // external
                            // objects
                            // linked to
                            // the vml
                            // by
                            // parsing
                            // it's rels
                            + f.substring(f.lastIndexOf("/") + 1) + ".rels")
if (target != null)
{
val embeds = OOXMLReader
.storeEmbeds(zip, target, p, externalDir) // passes
 // thru
                        // embedded
                        // objects
                        this.addOOXMLShape(arrayOf<Any>(vml, embeds))
}
else
this.addOOXMLShape(vml)
 /**/
                }
else if (ooxmlElement == "hyperlink")
{ // hyperlinks
c = cl[i] as Array<String> // don't strip path
for (j in hyperlinks.indices)
{
if (rId == (hyperlinks[j] as Array<String>)[0])
{
val h = hyperlinks[j] as Array<String>
try
{ // target= cl[2], ref= h[1], desc= h[2]
bk.getWorkSheet(this.sheetName)!!
.getCell(h[1]).setURL(rId, h[2], "") // TODO:
 // hyperlink
                                // text
                                // mark
                            }
catch (e:Exception) {
Logger.logErr("OOXMLAdapter.parse: failed setting hyperlink to cell "
+ h[1] + ":" + e.toString())
}

break
}
}
}
else if (OOXMLReader.parsePivotTables && ooxmlElement == "pivotTable")
{ // sheet-parent
 /*
                     * TODO: Do we really need to get rels ????
                     * // must lookup cacheid from rid of
                     * pivotCacheDefinitionX.xml in
                     * pivotTableDefinitionX.xml.rels
                     * target= OOXMLReader.getEntry(zip,p + "_rels/" +
                     * f.substring(f.lastIndexOf("/")+1)+".rels");
                     * ArrayList ptrels=
                     * parseRels(wrapInputStream(wrapInputStream(zip.
                     * getInputStream(target))));
                     * if (ptrels.size() > 1) { // what could this be?
                     * Logger.
                     * logWarn("OOXMLReader.parse: Unknown Pivot Table Association: "
                     * + ptrels.get(1));
                     * }
                     * String pcd= ((String[])ptrels.get(0))[1];
                     * pcd= pcd.substring(pcd.lastIndexOf("/")+1);
                     * Object cacheid= null;
                     * for (int z= 0; z < pivotCaches.size(); z++) {
                     * Object[] o= (Object[]) pivotCaches.get(z);
                     * if (pcd.equals(o[0])) {
                     * cacheid= o[1];
                     * break;
                     * }
                     * }
                     *
                     * target = getEntry(zip,p + f);
                     * PivotTableDefinition.parseOOXML(bk, /*cacheid, * /this,
                     * wrapInputStream(zip.getInputStream(target)));
                     */
                    try
{ // SAVE FOR LATER INPUT -- must do after all sheets
 // are input ...
                        pivotTables.put(p + f, bk
.getWorkSheet(this.sheetName))
}
catch (we:WorkSheetNotFoundException) {}

}
else if (ooxmlElement == "comments")
{ // parse comments
 // or notes
                    target = OOXMLReader.getEntry(zip, p + f)
parseCommentsXML(bk, OOXMLReader
.wrapInputStream(zip.getInputStream(target!!)))

 // Below are elements we do not as of yet handle
                }
else if ((ooxmlElement == "macro"
|| ooxmlElement == "activeX"
|| ooxmlElement == "table"
|| ooxmlElement == "vdependencies"
|| ooxmlElement == "oleObject"
|| ooxmlElement == "image"
|| ooxmlElement == "printerSettings"))
{

var attrs = ""
if ((shExternalLinkInfo != null && shExternalLinkInfo!!.get(rId) != null))
attrs = shExternalLinkInfo!!.get(rId)
OOXMLReader
.handleSheetPassThroughs(zip, bk, this, p, externalDir, c, attrs)
 // OOXMLReader.handlePassThroughs(zip, bk, this, p, c); //
                    // pass-through this file and any embedded objects as well
                }
else
{ // unknown type
Logger.logWarn(("OOXMLAdapter.parse:  XLSX Option Not yet Implemented " + ooxmlElement))
}
}
}
catch (e:IOException) {
Logger.logErr("OOXMLAdapter.parse failed: " + e.toString())
}

shExternalLinkInfo = null
}

/**
 * NOTE: commentsX.xml also needs legacy drawing info (vmlDrawingX.vml)
 * to define the text box itself including position and size, plus the vml elements
 * also define whether the note is hidden
 */
    internal fun parseCommentsXML(bk:WorkBookHandle, ii:InputStream) {
try
{
val factory = XmlPullParserFactory.newInstance()
factory.setNamespaceAware(true)
val xpp = factory.newPullParser()
xpp.setInput(ii, null) // using XML 1.0 specification
var eventType = xpp.getEventType()

val lastTag = java.util.Stack() // keep track of
 // element
            // hierarchy

            val authors = ArrayList()
var addr = ""
var authId = -1
var comment:Unicodestring? = null
 // ignore for now: phonetic properties (phoneticPr),
            // phonetic run (rPh)
            while (eventType != XmlPullParser.END_DOCUMENT)
{
if (eventType == XmlPullParser.START_TAG)
{
val tnm = xpp.getName()
if (tnm == "author")
{
authors.add(OOXMLReader.getNextText(xpp))
}
else if (tnm == "comment")
{
if (comment != null && "" != addr)
{
this.createNote(addr, comment!!, authors
.get(authId) as String)
}
addr = xpp.getAttributeValue("", "ref")
authId = Integer
.valueOf(xpp.getAttributeValue("", "authorId"))
.toInt()
comment = null
}
else if (tnm == "text")
{
 // read in text element
                        lastTag.push(tnm)
val t = Text.parseOOXML(xpp, lastTag, bk)
 // don't reset state vars as can there can be more
                        comment = t.commentWithFormatting
}
}
else if (eventType == XmlPullParser.END_TAG)
{}
eventType = xpp.next()
}
if ("" != comment!!.toString() && "" != addr)
{
this.createNote(addr, comment!!, authors.get(authId) as String)
}
}
catch (e:Exception) {
Logger.logErr("OOXMLAdapter.parseCommentsXML: " + e.toString())
}

return
}

/**
 * parse vml - legacy drawing info e.g. mso shapes and lines + note textboxes
 * <br></br>for now, legacy drawing info is just stored and not parsed into BIFF8 structures
 * <br></br>i.e. store everything but note textboxes at this time; intention is later on
 * to store all mso shapes and objects in BIFF8 records
 * <br></br>this vml is stored at the sheet level in the boundsheet's OOXMLShapes storage
 * <br></br>Notes textboxes are being created upon writeLegacyDrawingXML
 *
 * @param bk
 * @param sheet
 * @param ii
 * @return StringBuffer rep of saved vml
 */
    internal fun parseLegacyDrawingXML(bk:WorkBookHandle, ii:InputStream):StringBuffer {
/**
 * more info:
 * The Shape element is the basic building block of VML. A shape may exist on its own or within a Group
 * element. Shape defines many attributes and sub-elements that control the look and behavior of the shape. A
 * shape must define at least a Path and size (Width, Height). VML 1 also uses properties of the CSS2 style
 * attribute to specify positioning and sizing
 *
 * The ShapeType element defines a definition, or template, for a shape. Such a template is “instantiated” by
 * creating a Shape element that references the ShapeType. The shape can override any value specified by its
 * ShapeType, or define attributes and elements the ShapeType does not provide. A ShapeType may not
 * reference another ShapeType.
 * The attributes and elements a ShapeType uses are identical to those of the Shape element, with these
 * exceptions: ShapeType may not use the Type element, Visibility is always hidden.
 *
 * Regarding Notes:
 * The visible box shown for comments attached to cells is persisted using VML. The comment contents are
 * stored separately as part of SpreadsheetML.
 */
        val savedVml = StringBuffer()
try
{
val factory = XmlPullParserFactory.newInstance()
factory.setNamespaceAware(true)
val xpp = factory.newPullParser()
xpp.setInput(ii, null) // using XML 1.0 specification
var eventType = xpp.getEventType()
 // NOTE: since vml controls visibility (hidden or shown),
            // text box size, etc., notes are created upon VML parsing
            // and edited here for the actual text and formats ... ms's
            // legacy drawing stuff makes for alot of convoluted
            // processing ((;
            val nhs = FastAddVector()
run({ val anhs = bk.getWorkSheet(this.sheetName)!!
.commentHandles
for (i in anhs.indices)
{
nhs.add(anhs[i])
} })
while (eventType != XmlPullParser.END_DOCUMENT)
{
if (eventType == XmlPullParser.START_TAG)
{
var tnm = xpp.getName()
if (tnm == "shapelayout")
{
 // just store
                        savedVml.append(OOXMLReader.getCurrentElement(xpp))
}
else if (tnm == "shapetype")
{
 // if spt==202 (shape-type=text box) id="_x0000_t202"
                        // then it's note textbox
                        if (xpp.getAttributeValue("urn:schemas-microsoft-com:office:office", "spt") != "202")
{// if it's not a note textbox
 // shapetype, store it
                            savedVml.append(OOXMLReader.getCurrentElement(xpp))
}
else
 // ignore element - will be rebuilt upon write
                            OOXMLReader.getCurrentElement(xpp)
}
else if (tnm == "shape")
{ // this is basic
 // several types: can contain images, shapes and notes
                        if (!xpp.getAttributeValue("", "type")
.endsWith("_x0000_t202"))
{// if it's not a note
 // textbox, save it
                            // if type="#_x0000_t202" it's a note textbox
                            savedVml.append(OOXMLReader.getCurrentElement(xpp))
}
else
{// add note here, text and formatting will be
 // input upon Comments parse;
                            var r = -1
var c = -1
var visible = false
val bounds = ShortArray(8)
while (eventType != XmlPullParser.END_DOCUMENT)
{
if (eventType == XmlPullParser.START_TAG)
{
tnm = xpp.getName() // Anchor
if (tnm == "Row")
r = Integer
.valueOf(OOXMLReader
.getNextText(xpp)!!)
.toInt()
else if (tnm == "Column")
c = Integer
.valueOf(OOXMLReader
.getNextText(xpp)!!)
.toInt()
else if (tnm == "Visible")
visible = true
else if (tnm == "Anchor")
{
 // get a string rep of the bounds
                                        var sbounds = OOXMLReader
.getNextText(xpp)
 // prepare for parsing
                                        sbounds = sbounds!!
.replace(("[^0-9,]+").toRegex(), "")
val s = sbounds!!.split((",").toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
for (i in 0..7)
{
bounds[i] = java.lang.Short.valueOf(s[i])
.toShort()
}
}
}
else if (eventType == XmlPullParser.END_TAG)
{
if (xpp.getName() == "shape")
break
}
eventType = xpp.next()
}
val addr = ExcelTools
.formatLocation(intArrayOf(r, c))
for (i in nhs.indices)
{
val nh = nhs.get(i) as CommentHandle
if (nh.address!!.endsWith(addr))
{
if (visible)
nh.show()
nh.textBoxBounds = bounds
nhs.removeAt(i)
break
}
}
}
}
else if (tnm == "xml")
{ // ignore :)
}
else if (tnm == "imagedata")
{}
else
{ // just store
 // --
                        savedVml.append(OOXMLReader.getCurrentElement(xpp))
}
}
else if (eventType == XmlPullParser.END_TAG)
{}
eventType = xpp.next()
}
}
catch (e:Exception) {
Logger.logErr(("OOXMLAdapter.parseLegacyDrawingXML: " + e.toString()))
}

return savedVml
}

/**
 * given drawingML drawing.xml inputstream, parse each twoCellAnchor tag into appropriate image or chart and insert into sheet
 *
 * @param bk
 * @param sheet
 * @param imgFiles list of image or chart files (referenced in drawing.xml via rId)
 * @param ii       InputStream
 * @param zip      Current Open ZipOutputStream
 */
    internal fun parseDrawingXML(bk:WorkBookHandle, drawingFiles:ArrayList<*>?, ii:InputStream, zip:ZipFile, parentDir:String, externalDir:String) {
try
{
val lastTag = java.util.Stack() // keep track of
 // element
            // hierarchy

            val factory = XmlPullParserFactory.newInstance()
factory.setNamespaceAware(true)
val xpp = factory.newPullParser()

xpp.setInput(ii, null) // using XML 1.0 specification
var eventType = xpp.getEventType()
while (eventType != XmlPullParser.END_DOCUMENT)
{
if (eventType == XmlPullParser.START_TAG)
{
val tnm = xpp.getName()
if (tnm == "twoCellAnchor")
{ // beginning of DrawingML
 // for a single image or
                        // chart
                        lastTag.push(tnm)
 // TODO: handle group shapes which combine images,
                        // shapes and/or charts
                        // ********************************************************
                        val t = TwoCellAnchor
.parseOOXML(xpp, lastTag, bk).cloneElement() as TwoCellAnchor
if (t.hasImage())
{
val s = t.embed // rid of embedded object
if (s!!.indexOf("rId") == 0)
{ // should!
val imgFile = OOXMLReader
.parsePathForZip(OOXMLReader
.getFilename(drawingFiles!!, s)!!, parentDir)
val img = ZipEntry(imgFile)
val `is` = BufferedInputStream(
zip.getInputStream(img))
val im = ImageHandle(`is`, this)
this.insertImage(im)
im.name = t.name
im.shapeName = t.descr
im.bounds = TwoCellAnchor
.convertBoundsToBIFF8(this, t
.bounds) // must do after
 // insert
                                im.spPr = t.sppr // set image shape
 // properties
                                im.setEditMovement(t.editAs) // specify
 // how to
                                // resize or
                                // move
                                im.update() // update underlying image record
 // with set data
                            }
}
else if (t.hasChart())
{
val s = t.chartRId
if (s!!.indexOf("rId") == 0)
{ // should!
var chartfilename = OOXMLReader
.getFilename(drawingFiles!!, s)
var name = t.name
if (name == null || name == "null")
{
name = "Untitled Chart"
}
val ch = bk.createChart(name, bk
.getWorkSheet(this.sheetName))
ch!!.relativeBounds = TwoCellAnchor
.convertBoundsToBIFF8(this, t
.bounds) // must do after
 // insert
                                ch!!.setEditMovement(t.editAs) // specify
 // how to
                                // resize or
                                // move
                                ch!!.ooxmlName = name
chartfilename = OOXMLReader
.parsePathForZip(chartfilename!!, parentDir)
val chFile = ZipEntry(chartfilename!!)
 // must account for default chart settings: set
                                // fontx recs to default font for this workbook
                                // ...
                                ch!!.resetFonts() // reset all fonts for the
 // chart
                                ch!!.removeLegend() // not all charts have
 // legends!
                                val ps = chartfilename!!.lastIndexOf("/") + 1
val rels = OOXMLReader
.getEntry(zip, (chartfilename!!
.substring(0, ps) + "_rels/"
+ chartfilename!!.substring(ps)
+ ".rels"))
if (rels != null)
{ // chart file has embeds -
 // usually drawing ml which
                                    // defines userShapes
                                    // xxx TODO: REFACTOR to get these specifics
                                    // out
                                    val chartEmbeds = OOXMLReader
.parseRels(OOXMLReader
.wrapInputStream(zip
.getInputStream(rels!!)))
for (i in chartEmbeds.indices)
{
val dr = chartEmbeds
.get(i) as Array<String>
if (dr[0] == "userShape")
{ // should!
dr[1] = dr[1].substring((dr[1]
.lastIndexOf("/") + 1))
ch!!.addChartEmbed(arrayOf<String>(dr[0], externalDir + dr[1]))
OOXMLReader
.passThrough(zip, (parentDir + dr[1]), (externalDir + dr[1])) // Store
 // Embedded
                                            // Object
                                            // on
                                            // disk
                                            // for
                                            // later
                                            // retrieval
                                        }
else if (dr[0] == "image")
{
var parentp = OOXMLReader
.parsePathForZip(dr[1], parentDir)
parentp = parentp
.substring(0, (parentp
.lastIndexOf("/") + 1))
dr[1] = dr[1].substring((dr[1]
.lastIndexOf("/") + 1))
ch!!.addChartEmbed(arrayOf<String>(dr[0], externalDir + dr[1]))
OOXMLReader.passThrough(zip, (parentp + dr[1]), (externalDir + dr[1])) // save
 // the
                                            // original
                                            // target
                                            // file
                                            // for
                                            // later
                                            // re-packaging
                                        }
else if (dr[0] == "themeOverride")
{
var parentp = OOXMLReader
.parsePathForZip(dr[1], parentDir)
parentp = parentp
.substring(0, (parentp
.lastIndexOf("/") + 1))
dr[1] = dr[1].substring((dr[1]
.lastIndexOf("/") + 1))
ch!!.addChartEmbed(arrayOf<String>(dr[0], externalDir + dr[1]))
val target = OOXMLAdapter
.getEntry(zip, (parentp + dr[1]))
bk.workBook!!.theme!!
.parseOOXML(bk, OOXMLAdapter
.wrapInputStream(zip
.getInputStream(target)))
}
else
{
Logger.logWarn(("OOXMLAdapter.parseDrawingML: unknown chart embed " + dr[0]))
}
}
}
 // do after parsing rels in case there is
                                // override theme colors ...
                                ch!!.parseOOXML(OOXMLReader.wrapInputStream(zip
.getInputStream(chFile)))
}
}
else if (t.hasShape())
{
this.addOOXMLShape(t) // just store shape for later
 // output since prev.
                            // versions do not handle
                            // shapes
                            if (t.embed != null)
{ // if this shape has
 // embedded objects such
                                // as images
                                val imgFile = OOXMLReader
.parsePathForZip(OOXMLReader
.getFilename(drawingFiles!!, t
.embed)!!, parentDir) // look
 // up
                                // embedded
                                // rid
                                // in
                                // content
                                // list
                                // to
                                // get
                                // filename
                                t.embedFilename = imgFile // save embedded
 // filename for
                                // later
                                // retrieval
                                OOXMLReader
.passThrough(zip, imgFile, (externalDir + imgFile)) // Store Embedded
 // Object on disk
                                // for later
                                // retrieval
                            }
}
else
{ // TESTING!
Logger.logErr("OOXMLAdapter.parseDrawingXML: Unknown twoCellAnchor type")
}
}
else if (tnm == "oneCellAnchor")
{ // unclear if this
 // can be root
                        // of charts and
                        // images as
                        // well as
                        // shapes
                        lastTag.push(tnm)
val oca = OneCellAnchor
.parseOOXML(xpp, lastTag, bk).cloneElement() as OneCellAnchor
if (oca.hasImage())
{
val s = oca.embed // rid of embedded object
if (s!!.indexOf("rId") == 0)
{ // should!
val imgFile = OOXMLReader
.parsePathForZip(OOXMLReader
.getFilename(drawingFiles!!, s)!!, parentDir)
val img = ZipEntry(imgFile)
val `is` = BufferedInputStream(
OOXMLReader.wrapInputStream(zip
.getInputStream(img)))
val im = ImageHandle(`is`, this)
this.insertImage(im)
im.name = oca.name
im.shapeName = oca.descr
im.bounds = oca.bounds // must do after
 // insert
                                im.spPr = oca.sppr // set image shape
 // properties
                                im.update() // update underlying image record
 // with set data
                            }
}
else if (oca.hasChart())
{
val s = oca.embed
if (s!!.indexOf("rId") == 0)
{ // should!
var chart = OOXMLReader
.getFilename(drawingFiles!!, s)
var name = oca.name
if (name == null || name == "null")
{
name = "Untitled Chart"
}
val ch = bk.createChart(name, bk
.getWorkSheet(this.sheetName))
ch!!.relativeBounds = oca.bounds
 // ch.setChartTitle(name);
                                chart = OOXMLReader
.parsePathForZip(chart!!, parentDir)
val chFile = ZipEntry(chart!!)
ch!!.parseOOXML(OOXMLReader.wrapInputStream(zip
.getInputStream(chFile)))
}
}
else if (oca.hasShape())
{
this.addOOXMLShape(oca) // just store shape for
 // later output since
                            // prev. versions do not
                            // handle shapes
                        }
else
{ // TESTING!
Logger.logErr("OOXMLAdapter.parseDrawingXML: Unknown oneCellAnchor type")
}
}
else if (tnm == "userShapes")
{ // drawings ONTOP of
 // charts =
                        // Reference to
                        // Chart Drawing
                        // Part
                        Logger.logErr("OOXMLAdapter.parseDrawingXML: USER SHAPE ENCOUNTERED")
}
}
eventType = xpp.next()
}
}
catch (e:Exception) {
Logger.logErr(("OOXMLAdapter.parseDrawingXML: failed " + e.toString()))
}

}

/**
 * clear out object references in prep for closing workbook
 */
    public override fun close() {
wkbook = null
for (info in colinfos.values)
{
if (null != info)
info!!.close()
}
colinfos.clear()
var ii:Iterator<*> = rows.keys.iterator()
while (ii.hasNext())
{
val r = rows.get(ii.next())
r.close()
}
rows.clear()

cellsByRow = TreeMap<CellAddressible, BiffRec>(
CellAddressible.RowMajorComparator())

cellsByCol = TreeMap<CellAddressible, BiffRec>(
CellAddressible.ColumnMajorComparator())
 // TODO: clear recs
        arrayformulas.clear()
 // TODO: clear recs
        transferXfs.clear()
 // TODO: clear recs
        transferFonts!!.clear()
imageMap!!.clear()
charts.clear()
ooxmlObjects.clear()
if (ooxmlShapes != null)
ooxmlShapes!!.clear()

ooAutoFilter = null
mc.clear()
sheetView = null // OOXML sheet view object
sheetPr = null // OOXML sheetPr object
if (lastselection != null)
{
lastselection!!.close()
lastselection = null
}
if (protector != null)
{
protector!!.close()
protector = null
}

if (sheetNameRecs != null)
{
ii = sheetNameRecs!!.keys.iterator()
while (ii.hasNext())
{
val n = sheetNameRecs!!.get(ii.next()) as Name
n.close()
}
sheetNameRecs!!.clear()
}

for (i in cond_formats!!.indices)
{
val c = cond_formats!!.get(i) as Condfmt
c.close()
}
cond_formats!!.clear()

for (i in autoFilters.indices)
{
val a = autoFilters.get(i) as AutoFilter
a.close()
}
autoFilters.clear()

if (lastCell != null)
{
(lastCell as XLSRecord).close()
lastCell = null
}
if (lastRow != null)
{
lastRow!!.close()
lastRow = null
}

if (window2 != null)
{
window2!!.close()
window2 = null
}
if (scl != null)
{
scl!!.close()
scl = null
}
if (pane != null)
{
pane!!.close()
pane = null
}
if (dvalRec != null)
{
dvalRec!!.close()
dvalRec = null
}
if (header != null)
{
header!!.close()
header = null
}
if (footer != null)
{
footer!!.close()
footer = null
}
if (wsBool != null)
{
wsBool!!.close()
wsBool = null
}
if (guts != null)
{
guts!!.close()
guts = null
}
if (dimensions != null)
{
dimensions!!.close()
dimensions = null
}
if (myBof != null)
{
myBof!!.close()
myBof = null
}
if (myEof != null)
{
myEof!!.close()
myEof = null
}
if (myidx != null)
{
myidx!!.close()
myidx = null
}
for (i in printRecs!!.indices)
{
val r = printRecs!!.get(i) as XLSRecord
r.close()
}
printRecs!!.clear()

 // clear out refs by sheet recs
        for (j in SheetRecs.indices)
{
val r = SheetRecs.get(j) as XLSRecord
r.close()
}
SheetRecs.clear()
if (localrecs != null)
localrecs!!.clear()
 // col records

    }

companion object {

/**
 *
 */
    private val serialVersionUID = 8977216410574107840L

 // sheet types from grbit field offset 0
    internal val SHEET_DIALOG:Byte = 0x00
internal val XL4_MACRO:Byte = 0x01
internal val CHART:Byte = 0x02
internal val VBMODULE:Byte = 0x06

 // hidden states from grbit field offset 1
     val VISIBLE:Byte = 0x00
 val HIDDEN:Byte = 0x01
 val VERY_HIDDEN:Byte = 0x02

/**
 * associates external reference info with the r:id of the external reference
 * for instance, oleObject elements are associated with a shape Id that links back to a .vml file entry
 *
 * @param externalobjs
 * @param xpp
 */
    protected fun addExternalInfo(externalobjs:MutableMap<String, String>, xpp:XmlPullParser) {
 // String[] attrs= new String[xpp.getAttributeCount()-1];
        val attrs = ArrayList()
var rId = ""
 // int j= 0;
        for (i in 0 until xpp.getAttributeCount())
{
val n = xpp.getAttributeName(i)
if (n == "id")
rId = xpp.getAttributeValue(i)
else
 // attrs[j++]= n+ "=\"" + xpp.getAttributeValue(i) +"\"";
                attrs.add(n + "=\"" + xpp.getAttributeValue(i) + "\"")
}
var s = Arrays.asList<Any>(*attrs.toTypedArray()).toString() // 1.6 only
 // Arrays.toString(attrs.toArray());
        if (s.length > 2)
{
s = s.substring(1, s.length - 1)
 // 1.6 only s= s.replace(",", ""); // only issue is embedded
            // ,'s in quoted strings, lets assume not!
            s = StringTool.replaceText(s, ",", "") // only issue is embedded
 // ,'s in quoted strings,
            // lets assume not!
        }
externalobjs.put(rId, s)
}
}
}/**
 * Insert an image into the WorkBook
 *
 * @param im
 *//**
 * adds a cell to the Sheet
 *
 * @param obj
 * @param rc
 * @param FORMAT_ID
 * @return
 */