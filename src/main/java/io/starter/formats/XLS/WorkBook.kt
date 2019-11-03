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

import io.starter.OpenXLS.WorkBookException
import io.starter.OpenXLS.*
import io.starter.formats.OOXML.Theme
import io.starter.formats.XLS.charts.Ai
import io.starter.formats.XLS.charts.Chart
import io.starter.formats.XLS.charts.Fontx
import io.starter.formats.XLS.charts.GenericChartObject
import io.starter.formats.XLS.formulas.*
import io.starter.toolkit.FastAddVector
import io.starter.toolkit.Logger

import java.awt.Color
import java.io.*
import java.lang.reflect.Field
import java.lang.reflect.Modifier
import java.util.*

/**
 * <pre>
 * The WorkBook record represents an XLS workbook substream containing worksheets and associated records.
 *
</pre> *
 *
 * @see WorkBook
 *
 * @see Boundsheet
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

 class WorkBook:Serializable, XLSConstants, Book {

/**
 * Gets the date format used by this book.
 */
     var dateFormat:DateConverter.DateFormat = DateConverter.DateFormat.LEGACY_1900
internal set

private var DEBUGLEVEL = 0
private var bofct = 0
private var eofct = 0
private var indexnum = 0
 var defaultIxfe = 15
/**
 * Sets the calculation mode for the workbook.
 *
 * @param CalcMode
 * @see WorkBookHandle.setFormulaCalculationMode
 */
    /**
 * Sets the OpenXLS calculation mode for the workbook.
 *
 * @param CalcMode
 * @see WorkBookHandle.setFormulaCalculationMode
 */
     var calcMode = XLSConstants.CALCULATE_AUTO
private var defaultLanguage = 0                                    // default
 // language
    // code
    // for
    // the
    // current
    // workbook
    private var copying = false
 var isSharedupes = false
/**
 * @return Returns the xfrecs.
 */
     val xfrecs:AbstractList<*> = Vector()
private val indexes = Vector()
private val names = Vector()                            // ALL
 // of
    // the
    // names
    // (including
    // worksheet
    // scoped,
    // etc)
    private val orphanedPtgNames = Vector()
private val externalnames = Vector()
internal var formulas:AbstractList<*> = FastAddVector()
/**
 * Return the font records
 */
     val fontRecs:AbstractList<*> = Vector()
 var colorTable:Array<Color>

/**
 * The list of Format records indexed by format ID.
 */
    private val formats = TreeMap()
 val chartVect:AbstractList<*> = Vector()
/**
 * OOXML-specific
 */
    private val ooxmlObjects = Vector()                            // stores
 // OOXML
    // objects
    // external
    // to
    // workbook
    // e.g.
    // oleObjects,
    /**
 * returns the workbook codename used by vba macros OOXML-specific
 */
    /**
 * sets the workbook codename used by vba macros OOXML-specific
 *
 * @param s
 */
     var codename:String? = null                                    // stores
 // TODO: input into Codename record
 // OOXML
    // codename
    /**
 * returns the list of dxf's (incremental style info) (int) OOXML-specific
 */
    /**
 * sets the list of dxf's (incremental style info) (int) OOXML-specific
 */
     var dxfs:ArrayList<*>? = null                                    // 20090622
 // KSC:
    // stores
    // dxf's
    // (incremental
    // style
    // info)
    // per
    // workbook
    /**
 * returns the first non-hidden
 *
 *
 * sheet (int) in the workbook OOXML-specific
 *
 *
 * Mar 15, 2010
 *
 * @return
 */
    /**
 * sets the first sheet (int) in the workbook OOXML-specific
 *
 *
 * naive implementaiton does not account for hidden sheets
 *
 *
 * Mar 15, 2010
 *
 * @param f
 */
    // first sheet is hidden -- fix
 // all else failed
 var firstSheet = 0
get() {
try
{
if (!getWorkSheetByNumber(field).hidden)
return field
}
catch (x:Exception) {}

for (t in 0 until this.numWorkSheets)
{
try
{
if (!getWorkSheetByNumber(t).hidden)
{
this.firstSheet = t
return field
}
}
catch (x:Exception) {}

}
return field
}                                    // specifies
 // first
    // sheet
    // (ooxml)
    /**
 * return the OOXML theme for this workbook, if any
 *
 * @return
 */
    /**
 * sets the OOXML theme for this 2007 verison workbook
 *
 * @param t
 */
     var theme:Theme? = null

 // Reference Tracking
     var refTracker:ReferenceTracker? = ReferenceTracker()
private set

 // various
    /**
 * returns the Vector of Boundsheets
 */
     val sheetVect:AbstractList<*> = Vector()                            // TODO:
 // remove
    // this
    // variable?
    // its
    // duplicated
    // in
    // workSheets
    /**
 * @return
 */
     val hlinklookup:AbstractList<*> = Vector(
20)
/**
 * @return
 */
     val mergecelllookup:AbstractList<*> = Vector(
20)

 /*
     * Indirect formulas that need to be calculated after load
     * for reftracker
     */
    private var indirectFormulas:AbstractList<*> = ArrayList()
private val supBooks = Vector()

 // We should consider using an ordered collections class for
    // some of these?
    // an enumeration of worksheets for instance will not always
    // be in the same
    // order, causing tests to fail...
    private val workSheets = Hashtable(
100, .950f)
private val bookNameRecs = HashMap()
private val formulashash = Hashtable()

/**
 * Maps number format patterns to their IDs.
 */
    private val formatlookup = Hashtable(
30)

private val ptViews = Hashtable(
20)                                                                                        // TODO:
 // move
    // to
    // sheet
    // -
    // sheet-level
    // pivot
    // table
    // view
    // recordss
    private val ptstream = ArrayList()                        // wb-level
 // pivot
    // cache
    // definitions
    // usually
    // only
    // 1
    /**
 * return the pivot cache, if any
 *
 * @return
 */
    /**
 * set the pivot cache pointer
 *
 * @param pc initialized pivot cache
 */
     var pivotCache:PivotCache? = null                                    // if
 // has
    // pivot
    // tables
    // this
    // is
    // the
    // one
    // and
    // only
    // pivot
    // cache

    /**
 * retrieve the format cache - links string vers. of xf to xf rec
 * used for resusing xf's
 *
 * @return
 * @see FormatHandle.updateXf
 */
     var formatCache = HashMap()

/**
 * set the last processed Index record
 */
    internal var lastINDEX:Index? = null
private set
 var sharedStringTable:Sst? = null
internal set
private var lastBOF:Bof? = null
private var myexternsheet:Externsheet? = null
private var firstBOF:Bof? = null
/**
 * Get a substream by name.
 */
    override var streamer = createByteStreamer()
private set
 var tabID:TabID? = null
private set
private var win1:Window1? = null
private var calcmoderec:CalcMode? = null                                                        // determines
 // recalculation
    // mode
    // for
    // workbook
    // -
    // Manual,
    // Auto
    // ...
    private var drh:DefaultRowHeight? = null
private var currchart:Chart? = null
private var currai:Ai? = null
private var myADDINSUPBOOK:Supbook? = null                                    // for
 // external
    // names
    // SUPBOOK
    /**
 * get a handle to the factory
 */
    /**
 * get a handle to the Reader for this
 * WorkBook.
 */
    override var factory:WorkBookFactory? = null
/**
 * get a handle to the ContinueHandler
 */
    override var continueHandler = ContinueHandler(
this)
private set
private var usersview:Usersviewbegin? = null
private var lasteof:Eof? = null
private var xl2k:BiffRec? = null
 var msodgMerge:AbstractList<*> = Vector()
 var msodg:MSODrawingGroup? = null
 var lastSPID = 1024                                    // 20071030
 // last
    // or
    // next
    // SPID
    // (=
    // shape
    // ID
    // or
    // image
    // ID);
    // incremented
    // upon
    // new
    // images
    // ...
    // appropriate
    // to
    // store
    // at
    // book
    // level
    // (?)
    private var currdrw:MSODrawing? = null

private var protector:BookProtectionManager? = null

/**
 * Gets this sheet's SheetProtectionManager.
 */
     val protectionManager:BookProtectionManager
get() {
if (protector == null)
protector = BookProtectionManager(this)
return protector
}

/**
 * Get the MsoDrawingGroup for this workbook
 *
 * @return msodrawinggroup
 */
    /**
 * Set the msodrawinggroup for this workbook
 *
 * @param msodg
 */
     var msoDrawingGroup:MSODrawingGroup?
get() =msodg
set(msodg) {
this.msodg = msodg
this.msodg!!.workBook = this
this.msodg!!.streamer = this.streamer
}

/**
 * Return some useful statistics about the WorkBook
 *
 * @return
 */
     val stats:String
get() =getStats(false)

 val xlsVersionString:String
get() =lastBOF!!.xlsVersionString

 val charts:Array<Chart>
get() {
val chts = arrayOfNulls<Chart>(this.chartVect.size)
return chartVect.toTypedArray() as Array<Chart>
}

/**
 * return all pivot table views (==Sxview records)
 * <br></br>SxView is the top-level record of a Pivot Table
 * as distinct from the PivotCache (stored data source in a LEOFile Storage)
 * and PivotTable Stream (SxStream top-level record)
 *
 * @return
 */
     val allPivotTableViews:Array<Sxview>
get() {
val sv = arrayOfNulls<Sxview>(ptViews.size)
val x = ptViews.elements()
var t = 0
while (x.hasMoreElements())
{
sv[t++] = x.nextElement() as Sxview
}
return sv
}

 val nPivotTableViews:Int
get() =ptViews.size

/**
 * get the Externsheet
 */
     val externSheet:Externsheet?
get() {
if (myexternsheet == null)
addExternsheet()
return myexternsheet
}

/**
 * Get a string array of external names
 *
 * @return externalNames
 */
     val externalNames:Array<String>
get() {
val n = arrayOfNulls<String>(externalnames.size)
externalnames.toTypedArray()
return n
}

/**
 * Get a collection of all names in the workbook
 */
     val workbookScopedNames:Array<Name>
get() {
val a = ArrayList(this.bookNameRecs.values)
val n = arrayOfNulls<Name>(a.size)
a.toTypedArray()
return n
}

/**
 * returns the List of Formulas in the book
 *
 * @return
 */
     val formulaList:List<*>
get() =formulas

/**
 * Returns the recalculation mode for the Workbook:
 * <br></br>0= Manual
 * <br></br>1= Automatic
 * <br></br>2= Automatic except for multiple table operations
 *
 * @return int
 */
     val recalculationMode:Int
get() =this.calcmoderec!!.recalcuationMode

 val numFonts:Int
get() =fontRecs.size

/**
 * Gets the number of custom number formats registered on this book.
 */
     val numFormats:Int
get() =formats.size

/**
 * sets the selected worksheet
 */
     val selectedSheetNum:Int
get() =win1!!.currentTab

internal var lastFormula:Formula? = null
internal var countryRec:XLSRecord? = null
internal var inChartSubstream = false
internal var chartTemp = ArrayList()

/**
 * get a handle to the Usersviewbegin for the workbook
 */
     val usersviewbegin:Usersviewbegin
get() {
if (usersview == null)
{
usersview = Usersviewbegin()
streamer.addRecord(usersview!!)
this.addRecord(usersview!!, false)
}
return usersview
}

/**
 * Get the number of worksheets in this WorkBook
 */
     val numWorkSheets:Int
get() =this.sheetVect.size

/**
 * get the number of formulas in this WorkBook
 *
 * @return
 */
     val numFormulas:Int
get() =formulas.size

/**
 * get the number of Cells in this WorkBook
 */
     val numCells:Int
get() {
var cellnum = 0
val e = workSheets.elements()
while (e.hasMoreElements())
{
val b = e.nextElement() as Boundsheet
cellnum += b.numCells
}
return cellnum
}

/**
 * get all of the Cells in this WorkBook
 */
     val cells:Array<BiffRec>
get() {
val cellz = FastAddVector()
for (i in 0 until workSheets.size)
{
try
{
val b = this.getWorkSheetByNumber(i)
val cz = b.cells
for (x in cz.indices)
{
cellz.add(cz[x])
}
}
catch (e:Exception) {
Logger.logErr("Error retrieving worksheet for getCells: $e")
}

}
val cellzr = arrayOfNulls<BiffRec>(cellz.size)
cellz.toTypedArray()
return cellzr
}

/**
 * @return Returns the lastbound.
 */
    /**
 * @param lastbound The lastbound to set.
 */
     var lastbound:Boundsheet? = null

override// 2003-vers
 val fileName:String
get() =if (this.factory != null) this.factory!!.fileName else "New Spreadsheet"

 val numXfs:Int
get() =xfrecs.size

/**
 * Get the typename for this object.
 */
    internal val typeName:String
get() ="WorkBook"

/**
 * Returns a map of xml strings representing the XF's in this workbook/Integer of lookup.
 * These are used as a comparitor to determine if additional xf's need to be brought in or
 * not and to give the new XF number if the xf exists.
 *
 *
 * Changed 20080226 KSC: to use XF toString as XML is limited in format, toString is more complete
 *
 * @return map (String XfXml, Integer xfLookup)
 */
     val xfrecsAsString:Map<*, *>
get() {
val retMap = HashMap()
for (xfNum in 1 until this.numXfs)
{
val x = this.getXf(xfNum)
val xml = x!!.toString()
retMap.put(xml, Integer.valueOf(xfNum))
}
return retMap
}

/**
 * Returns a map of xml strings representing the XF's in this workbook/Integer of lookup.
 * These are used as a comparitor to determine if additional xf's need to be brought in or
 * not and to give the new XF number if the xf exists.
 *
 * @return map (String XfXml, Integer xfLookup)
 */
     val fontRecsAsXML:Map<*, *>
get() {
val retMap = HashMap()
for (i in this.fontRecs.indices.reversed())
{
val fnt = this.fontRecs[i] as Font
val xml = "<FONT><" + fnt.xml + "/></FONT>"
retMap.put(xml, Integer.valueOf(fnt.idx))
}
return retMap
}

/**
 * Returns whether this book uses the 1904 date format.
 *
 */
     val is1904:Boolean
@Deprecated("Use {@link #getDateFormat()} instead.")
get() =this.dateFormat == DateConverter.DateFormat.LEGACY_1904

 // OOXML Additions
    /**
 * returns truth of "Excel 2007" format
 */
    /**
 * set truth of "Is Excel 2007"
 * true increases maximums of base storage, etc.
 *
 * @param b
 */
     var isExcel2007 = false

/**
 * Returns all strings that are in the SharedStringTable for this workbook.  The SST contains
 * all standard string records in cells, but may not include such things as strings that are contained
 * within formulas.   This is useful for such things as full text indexing of workbooks
 *
 * @return Strings in the workbook.
 */
     val allStrings:Array<String>
get() {
val al = this.sharedStringTable!!.allStrings
var s = arrayOfNulls<String>(al.size)
s = al.toTypedArray() as Array<String>
return s
}

 // TODO Auto-generated method stub
 // TODO Auto-generated method stub
 var calcId:Int
get() =0
set(i) {

}

protected fun reflectiveClone(source:WorkBook) {
for (field in WorkBook::class.java!!.getDeclaredFields())
{
if (Modifier.isStatic(field.getModifiers()))
continue

try
{
field.set(this, field.get(source))
}
catch (e:IllegalAccessException) {
throw RuntimeException(e)
}

}
}

protected fun createByteStreamer():ByteStreamer {
return ByteStreamer(this)
}

/**
 * init the ImageHandles
 */
    internal fun initImages() {
lastSPID = Math.max(lastSPID, msodg!!.spidMax) // 20090508 KSC:
 // lastSPID should
        // also account for
        // charts
        // [BUGTRACKER 2372
        // copyChartToSheet
        // error]
        // 20071217 KSC: clear out imageMap before inputting!
        for (x in 0 until this.getWorkSheets().size)
{
this.getWorkSheets()[x].imageMap!!.clear()
}
for (i in 0 until msodg!!.msodrawingrecs.size)
{ // 20070914
 // KSC:
            // store
            // msodrawingrecs
            // with
            // MSODrawingGroup
            // instead
            // of
            // here
            val rec = msodg!!.msodrawingrecs[i] as MSODrawing
lastSPID = Math.max(lastSPID, rec.getlastSPID()) // valid for
 // header
            // msodrawing
            // record(s)
            val imgdx = rec.getImageIndex() - 1 // it's 1-based
val imageData = msodg!!.getImageBytes(imgdx)
if (imageData != null)
{
val im = ImageHandle(imageData, rec.sheet)
im.setMsgdrawing(rec) // Link 2 actual Msodrawing rec
im.name = rec.name // set image name from rec ...
im.shapeName = rec.getShapeName() // set shape name as well
 // ...
                im.imageType = msodg!!.getImageType(imgdx)// 20100519 KSC:
 // added!
                rec.sheet!!.imageMap!![im] = Integer.valueOf(imgdx)
}
}
}

 /*
     * Once all the records are parsed, the msodrawinggroup
     * needs to be parsed.
     * We first merge all the msodrawinggroup records. From the
     * till now experiment,
     * there are maximum of two such records, adjacent to each
     * other. If there are two,
     * the size of the first is only 8228 and the remaining data
     * is in second record.
     * However, when writing, it doesn't seem important to break
     * it into two and can be written
     * as one.
     *
     * Once msodrawingrecords are parsed, it is not used.
     * However, we have a msodrawingglobal class that
     * records the current count of msodrawing related records,
     * like shape count, max shape id till now etc.
     * We need to maintain that for writing (creating the
     * msodrawinggroup records) later on.
     */
     fun mergeMSODrawingRecords() {
 // 20070915 KSC: Now don't re-initialize Msodrawing recs;
        // just merge and parse
        if (msodg != null)
{
msodg!!.mergeAndRemoveContinues()
for (i in 1 until msodgMerge.size)
{
msodg!!.mergeRecords(msodgMerge[i] as MSODrawingGroup) // merges
 // and
                // removes
                // secondary
                // msodrawinggroups
                // 20071003 KSC: get rid of secondary msodg's, will be
                // created upon createMSODGContinues
                this.streamer
.removeRecord(msodgMerge[i] as MSODrawingGroup) // remove
 // existing
                // continues
                // from
                // stream
            }
while (msodgMerge.size > 1)
msodgMerge.removeAt(msodgMerge.size - 1)
msodg!!.parse()
this.initImages()
}
}

/**
 * after changing the MSODrawings on a sheet, this is called to
 * update the header Msodrawing rec.
 * <br></br>must sum up all other mso's on the particular sheet SPCONTAINERLEN and update the sheet mso header ...
 * <br></br>this calculation is quite experimental, so far it's working in all known cases ...
 *
 * @param bs
 */
     fun updateMsodrawingHeaderRec(bs:Boundsheet) {
val msdHeader = msodg!!.getMsoHeaderRec(bs)
if (msdHeader != null)
{
var spContainerLength = 0 // count of all other spcontainer lengths
 // (sum=header spgroupcontainer)
            var otherContainerLength = 0 // count of other containers
 // (solvercontainer,etc) added to
            // dgcontainerlength
            var numshapes = 2 // 20100324 KSC: total guess, really
val totalDrawingRecs = msodg!!.msodrawingrecs.size
for (z in 0 until totalDrawingRecs)
{
val rec = msodg!!.msodrawingrecs[z] as MSODrawing
if (rec.sheet == bs)
{
if (rec != msdHeader && !rec.isHeader)
{ // added header
 // check- seems
                        // like *can*
                        // have multiple
                        // header recs
                        // in charts!
                        spContainerLength += rec.spContainerLength
otherContainerLength += rec.solverContainerLength
if (rec.isShape)
 // if it's a shape-type mso, count;
                            // there are other mso-types that
                            // are not SPCONTAINERS; apparently
                            // don't count these ...
                            numshapes++
}
}
}
msdHeader
.updateHeader(spContainerLength, otherContainerLength, numshapes, msdHeader
.getlastSPID())
}
}

/**
 * For workbooks that do not contain an MSODrawing group create a new one,
 * if the drawing group already exists return the existing
 *
 * @return
 */
     fun createMSODrawingGroup():MSODrawingGroup? {
if (msodg != null)
return msodg
this.msoDrawingGroup = MSODrawingGroup
.prototype as MSODrawingGroup?
return msodg
}

/**
 * Return some useful statistics about the WorkBook
 *
 * @return
 */
     fun getStats(usehtml:Boolean):String {
var rex = "\r\n"
if (usehtml)
rex = "<br/>"

var ret = "-------------------------------------------$rex"
ret += "OpenXLS Version:     " + WorkBookHandle.version + rex
ret += "Excel Version:        $xlsVersionString$rex"
ret += "-------------------------------------------------\r\n"
ret += "Statistics for:       $this$rex"
ret += "Number of Worksheets: " + this.numWorkSheets + rex
ret += "Number of Cells:      " + this.numCells + rex
ret += "Number of Formulas:   " + this.numFormulas + rex
ret += "Number of Charts:     " + this.chartVect.size + rex
ret += "Number of Fonts:      " + this.numFonts + rex
ret += "Number of Formats:    " + this.numFormats + rex
ret += "Number of Xfs:        " + this.numXfs + rex
 // ret += "StringTable: " + this.stringTable.toString() +
        // rex;
        ret += "-------------------------------------------------\r\n"
return ret
}

private fun addMergedcells(c:Mergedcells) {
this.mergecelllookup.add(c)
}

private fun addHlink(r:Hlink) {
this.hlinklookup.add(r)
}

/**
 * take care of any lazy updating before output
 */
     fun prestream() {
if (this.msoDrawingGroup != null)
this.msoDrawingGroup!!.prestream()
}

 fun getSupBooks():Array<Supbook> {
val sbs = arrayOfNulls<Supbook>(this.supBooks.size)
return this.supBooks.toTypedArray() as Array<Supbook>
}

/**
 * returns the list of OOXML objects which are external or auxillary to the main workbook
 * e.g. theme, doc properties
 *
 * @return
 */
     fun getOOXMLObjects():List<*> {
return ooxmlObjects
}

/**
 * adds the object-specific signature of the external or auxillary OOXML object
 * Object should be of String[] form,
 * key, path, local path  + filename [, rid,  [extra info], [embedded information]]
 * e.g. theme, doc properties
 *
 * @param o
 */
     fun addOOXMLObject(o:Any) {
if ((o as Array<String>)[0] != "externalLink")
ooxmlObjects.add(o)
else
ooxmlObjects.add(0, o) // ensure ExternalLinks are 1st because they
 // are linked via rId in workbook.xml
    }

/**
 * return the External Supbook record associated with the desired externalWorkbook
 * will create if bCreate
 *
 * @param externalWorkbook String URL (name) of External Workbook
 * @param bCreate          if true, will create an external SUPBOOK record for the externalWorkbook
 * @return Supbook
 */
     fun getExternalSupbook(externalWorkbook:String?, bCreate:Boolean):Supbook? {
var sb:Supbook? = null
if (externalWorkbook == null)
return null
for (i in this.supBooks.indices)
{
if ((this.supBooks.get(i) as Supbook).isExternalRecord)
if (externalWorkbook
.equals((this.supBooks.get(i) as Supbook)
.externalWorkBook!!, ignoreCase = true))
sb = this.supBooks.get(i)
}
if (sb == null && bCreate)
{ // create
sb = Supbook.getExternalPrototype(externalWorkbook) as Supbook
val loc = (supBooks.get(supBooks.size - 1) as Supbook)
.recordIndex // must have at least one global supbook
 // present
            streamer.addRecordAt(sb, loc + 1) // external supbooks appear to be
 // before "normal" supbooks
            // [BugTracker 1434]
            this.supBooks.add(sb) // 20080714 KSC: add at beginning -- correct
 // in all cases?
            // this.addRecord(sb,false); // "" no need
        }
return sb
}

/**
 * return the index into the Supbook records for this supbook
 *
 * @param sb
 * @return int
 */
     fun getSupbookIndex(sb:Supbook):Int {
for (i in this.supBooks.indices)
{
if (this.supBooks.get(i) === sb)
return i
}
return -1
}

 fun getPivotTableView(nm:String):Sxview {
return ptViews.get(nm)
}

 fun addPivotTable(sx:Sxview) {
this.ptViews.put(sx.tableName!!, sx) // Pivot Table View ==Top-level
 // record for a Pivot Table
    }

/**
 * return the Externsheet for this book
 *
 * @param create a new Externsheet if it does not exist
 * @return the Externsheet
 */
     fun getExternSheet(create:Boolean):Externsheet? {
if (myexternsheet == null && create)
{
addExternsheet()
}
return myexternsheet
}
/**
 * default constructor -- do init
 */
    init{

val cm = System.getProperties()[WorkBook.CALC_MODE_PROP]
if (cm != null)
{
try
{
this.calcMode = Integer.parseInt(cm.toString())
}
catch (e:Exception) {
Logger.logWarn("Invalid Calc Mode Setting in System properties:$cm")
}

}
if (System.getProperties()["io.starter.OpenXLS.sharedupes"] != null)
{
this.isSharedupes = System.getProperties()["io.starter.OpenXLS.sharedupes"] == "true"
if (this.isSharedupes)
{
this.setDupeStringMode(WorkBook.SHAREDUPES)
}
}
this.initBuiltinFormats()
 // re-init color table: initial state of color table if
        // Pallete record exists, changes may occur
        colorTable = arrayOfNulls<java.awt.Color>(FormatHandle.COLORTABLE.size)
for (i in FormatHandle.COLORTABLE.indices)
colorTable[i] = FormatHandle.COLORTABLE[i]
}

/**
 * Gets the format ID for a given number format pattern.
 * This lookup is completely case-insensitive. For most patterns this
 * correctly reflects the case-insensitivity of the tokens. Custom patterns
 * containing string literals could be matched incorrectly.
 *
 * @param pattern the number format pattern to look up
 * @return the format ID of the given pattern or -1 if it's not recognized
 */
     fun getFormatId(pattern:String):Short {
val res = formatlookup.get(pattern.toUpperCase()) as Short
return res ?: -1
}

/**
 * Initializes the format lookup to contain the built-in formats.
 */
    private fun initBuiltinFormats() {
val formats = FormatConstantsImpl.builtinFormats

for (i in formats.indices)
formatlookup.put(formats[i][0].toUpperCase(), java.lang.Short
.valueOf(formats[i][1], 16))
}

/**
 * Init names at Post-load
 */
     fun initializeNames() {
for (i in names.indices)
{
val n = names.get(i) as Name
n.parseExpression() // evaluate expression at postload, after sheet
 // recs are loaded
        }
}

/**
 * add a Name object to the collection of names
 */
     fun addName(n:Name):Int {
if (n.getItab().toInt() != 0)
{
 // its a sheet level name
            try
{
val b = this.getWorkSheetByNumber(n.getItab() - 1)// one
 // based
                // pointer
                b.addLocalName(n)
n.setSheet(b)
}
catch (e:WorkSheetNotFoundException) {}

}
else
{
val sName = n.nameA // returns upper case name
val existo = bookNameRecs.get(sName)
if (existo != null)
{ // handle duplicate named ranges
val bnam = n.toString()
if (bnam.indexOf("Built-in:") != 0)
{
try
{
if ((existo as Name).location != null)
 // use
                            // original
                            // - as good
                            // a guess
                            // as any
                            return -1 // an invalid sheet
}
catch (e:Exception) {}

 // if original does not have a location set, use this one
                    // instead
                    this.names.remove(existo)
this.bookNameRecs.remove(sName)
}
}
this.bookNameRecs.put(sName, n)
}
this.names.add(n)
if (myexternsheet != null && n.getExternsheet() == null)
{
try
{
n.setExternsheet(myexternsheet) // update sheet reference
}
catch (e:WorkSheetNotFoundException) {
Logger.logWarn("WorkBookHandle.addName() setting Externsheet failed for new Name: $e")
}

}
return names.size - 1
}

 fun addNameUpdateSheetRefs(n:Name, origWorkBookName:String) {
if (bookNameRecs.get(n.nameA) == null)
{
val newName = Name(this, n.name)
try
{
newName.location = n.location
}
catch (e:Exception) {}

if (myexternsheet != null && newName.getExternsheet() == null)
{
try
{
newName.setExternsheet(myexternsheet) // update sheet
 // reference
                }
catch (e:WorkSheetNotFoundException) {
Logger.logWarn("WorkBookHandle.addName() setting Externsheet failed for new Name: $e")
}

}
newName.updateSheetRefs(origWorkBookName)
}
}

/**
 * Store an external name
 *
 * @param n = String describing the name
 * @return int location of the name
 */
     fun addExternalName(n:String):Int {
this.externalnames.add(n)
return externalnames.size // one-based index
}

/**
 * Get the external name at the specified index.
 *
 * @param t index of the name
 * @return name at the index, empty string if it doesn't exist.
 *
 *
 * Why are we calling getExternalName one based, then removing for ordinal,  internal processes should always
 * be 0,1,2,3...  -NR 1/06
 */
     fun getExternalName(t:Int):String {
return if (t > 0) this.externalnames.get(t - 1) else "" // one-based index
}

/**
 * For workbooks that do not contain an externsheet
 * this creates the externsheet record with one 0000 record
 * and the related Supbook rec
 */
     fun addDefaultExternsheet() {
val sbb = Supbook.getPrototype(this.numWorkSheets) as Supbook
var l = this.sharedStringTable!!.recordIndex
streamer.addRecordAt(sbb, l++)
supBooks.add(sbb)
val ex = Externsheet
.getPrototype(0x0000, 0x0000, this) as Externsheet
streamer.addRecordAt(ex, l)
this.addRecord(ex, false)
}

/**
 * apparently this method adds an External name rec and returns the ilbl
 *
 *
 * Correct structure is
 * Supbook
 * Externname
 * Supbook
 * Externsheet
 *
 * @param s
 * @return
 * @see PtgNameX
 */
     fun getExtenalNameNumber(s:String):Int {
val i = externalnames.indexOf(s)
if (i > -1)
 // got it
            return i + 1
if (this.externSheet == null)
this.addDefaultExternsheet()

 // not found; add a new EXTERNNAME record to list of add-ins
        val n = addExternalName(s)
try
{
var loc:Int
if (myADDINSUPBOOK == null)
{
val sb = Supbook.addInPrototype as Supbook
loc = this.externSheet!!.recordIndex
streamer.addRecordAt(sb, loc++)
this.addRecord(sb, false)
myADDINSUPBOOK = sb
supBooks.add(sb)
val externref = this.externSheet!!.virtualReference
 // Add EXTERNNAME record after ADD-IN SUPBOOK record and
                // after existing EXTERNNAME records
                val exn = Externname.getPrototype(s) as Externname
streamer.addRecordAt(exn, loc++)
this.addRecord(exn, false)
}
else
{
loc = streamer.getRecordIndex(myADDINSUPBOOK!!)
 // Add EXTERNNAME record after ADD-IN SUPBOOK record and
                // after existing EXTERNNAME records
                val exn = Externname.getPrototype(s) as Externname
streamer.addRecordAt(exn, loc + externalnames.size)
this.addRecord(exn, false)
}

}
catch (e:Exception) {
Logger.logWarn("Error adding externname: $e")
}

return n
}

/**
 * Get a collection of all names in the workbook
 */
     fun getNames():Array<Name> {
val n = arrayOfNulls<Name>(names.size)
names.toTypedArray()
return n
}

/**
 * returns the array of Formulas in the book
 *
 * @return
 */
     fun getFormulas():Array<Formula> {
val n = arrayOfNulls<Formula>(formulas.size)
formulas.toTypedArray()
return n
}

/**
 * remove a formula from the book
 *
 * @param fmla
 */
     fun removeFormula(fmla:Formula) {
this.formulashash.remove(fmla.cellAddressWithSheet)
formulas.remove(fmla)
fmla.destroy()
}

/**
 * Sets the recalculation mode for the Workbook:
 * <br></br>0= Manual
 * <br></br>1= Automatic
 * <br></br>2= Automatic except for multiple table operations
 */
     fun setRecalcuationMode(mode:Int) {
this.calcmoderec!!.setRecalculationMode(mode)
}

/**
 * returns a Named range by number
 *
 * @param t
 * @return
 */
     fun getName(t:Int):Name {
return this.names.get(t - 1)
}

/**
 * rename the NamedRange in the lookup map
 *
 * @param t
 * @return
 */
     fun setNewName(oldname:String, newname:String) {
var oldname = oldname
var newname = newname
if (oldname == newname)
return
oldname = oldname.toUpperCase() // case-insensitive
newname = newname.toUpperCase() // ""
val old = bookNameRecs.get(oldname) ?: return
// new name?
bookNameRecs.remove(oldname)
bookNameRecs.put(newname, old)

}

/**
 * Re-assocates ptgrefs that are pointing to a name that has been deleted then
 * is recreated
 *
 * @param name
 */
     fun associateDereferencedNames(name:Name) {
val i = orphanedPtgNames.iterator()
val theName = name.name
while (i.hasNext())
{
val x = i.next() as IlblListener
if (x.storedName.equals(theName, ignoreCase = true))
{
x.ilbl = this.getNameNumber(theName).toShort()
x.addListener()
}
}
}

/**
 * returns a named range by name string
 *
 *
 * This method will first attempt to look in the book names, then the sheet names,
 * obviously different scoped names can have the same identifying name, so this could return
 * one of multiple names if this is the case
 *
 * @param t
 * @return
 */
     fun getName(nameRef:String):Name? {
var nameRef = nameRef
nameRef = nameRef.toUpperCase() // case-insensitive
var o:Any? = this.bookNameRecs.get(nameRef)
if (o == null)
{
val shts = this.getWorkSheets()
for (i in shts.indices)
{
o = shts[i].getName(nameRef)
if (o != null)
return o as Name?
}
}
return o as Name?
}

/**
 * returns a scoped named range by name string
 *
 * @param t
 * @return
 */
     fun getScopedName(nameRef:String):Name? {
return bookNameRecs.get(nameRef.toUpperCase()) ?: return null
}

/**
 * Returns the ilbl of the name record associated with the string passed in.
 * If the name does not exist, it get's created without a location reference.
 * This is needed to support formula creation with non-existent names referenced.
 *
 * @param t, the name record to search for
 * @return the index of the name
 */
     fun getNameNumber(nameStr:String?):Int {
for (i in names.indices)
{
val n = this.names.get(i) as Name
if (n.name.equals(nameStr!!, ignoreCase = true))
return i + 1
}
 // no name exists, we need to create one.
        val myName:Name
myName = Name(this, nameStr)
 // myName = new Name(this, true);
        // myName.setName(nameStr);
        val nmx = this.getNames()
var namepos = -1
if (nmx.size >= 1)
{
namepos = nmx[nmx.size - 1].recordIndex
}
else
{
namepos = this.externSheet!!.recordIndex
}
namepos++
this.streamer.addRecordAt(myName, namepos)
return this.getNameNumber(nameStr)
}

/**
 * Get's the index for this particular front.
 *
 *
 * NOTE:  this doesn't actually get a "matching" font, it has to be the exact font.
 * 20070826 KSC: changed to match font characterstics, not just return exact matching font
 */
     fun getFontIdx(f:Font):Int {
 // 20070819 KSC: Try this to see if better! Matches 6 key
        // attributes (size, name, color, etc.)
        for (i in fontRecs.indices.reversed())
{ // start from the back so
 // don't initially match
            // defaults...
            if (f.matches(fontRecs[i] as Font))
return if (i > 3) i + 1 else i
}
 // return fonts.indexOf(f);
        return -1
}

/**
 * Get's the index for this font, based on matching through
 * xml strings.  If the font doesn't exist in the book it returns -1;
 *
 * @return KSC: is this method necessary now with above getFontIdx changes?
 */
     fun getMatchingFontIndex(f:Font):Int {
val fontmap = this.fontRecsAsXML
val o = fontmap.get("<FONT><" + f.xml + "/></FONT>")
if (o != null)
{
val I = o as Int?
return I!!.toInt()
}
else
{
return -1
}
}

/**
 * InsertFont inserts a Font record into the workbook level stream,
 * For some reason, the addFont only puts it into an array that is never accessed
 * on output.  This may have a reason, so I am not overwriting it currently, but
 * let's check it out?
 */
     fun insertFont(x:Font):Int {
val insertIdx = this.getFont(this.numFonts).recordIndex
 // perform default add rec actions
        this.streamer.addRecordAt(x, insertIdx + 1)
x.idx = -1 // flag to add into font array
this.addRecord(x, false) // also adds to font array so no need for
 // additional addFont below
        return fontRecs.indexOf(x)
}

/**
 * add a Font object to the collection of Fonts
 */
     fun addFont(f:Font):Int {
fontRecs.add(f)
return if (fontRecs.size > 4) fontRecs.size else fontRecs.size - 1
}

/**
 * Get the font at the specified index.  Note that the number 4 does not exist, so index correctly based of that.
 *
 *
 * So,  if you call getFont(5), you are really doing getFont(4) from the internal array
 *
 * @param t
 * @return
 */
     fun getFont(t:Int):Font {
var t = t
if (t >= 4)
{
t--
}
if (this.fontRecs.size >= t)
{
if (t >= fontRecs.size)
{
Logger.logWarn("font " + t
+ " not found. Workbook contains only: " + fontRecs.size
+ " defined fonts.")
return fontRecs[0] as Font
}
else
{
return fontRecs[t] as Font
}
}
return this.fontRecs[0] as Font
}

/**
 * Inserts a newly created Format record into the workbook.
 * This method handles assigning the format ID and adding the record to the
 * workbook. If the record is already part of the workbook use
 * [.addFormat] instead.
 */
     fun insertFormat(format:Format):Int {
val last:Format
try
{
last = formats.get(formats.lastKey())
}
catch (e:NoSuchElementException) {
 /*
             * There are no other Format records in the workbook.
             * This shouldn't happen because most (all?) Excel files
             * contain
             * Format records for the locale-specific (and thus not
             * implied)
             * built-in formats. If it does happen, either we need to
             * re-assess
             * the above assumption or this method was called before the
             * Format
             * records were parsed. Either way we need to know about it.
             */
            throw AssertionError("WorkBook.insertFormat called but no "
+ "Format records exist. This should not happen. Please "
+ "report this error to support@extentech.com.")
}

 // Add it to the streamer and workbook
        streamer.addRecordAt(format, last.recordIndex + 1)
addRecord(format, false)

 // Give it a format ID
        if (format.ifmt.toInt() == -1)
format.ifmt = Math.max(last.ifmt + 1, 164).toShort()

 // Add it to the format lookups
        addFormat(format)

return format.ifmt.toInt()
}

@Throws(FormulaNotFoundException::class)
 fun getFormula(cellAddress:String):Formula {

return formulashash.get(cellAddress) as Formula ?: throw FormulaNotFoundException(
                "no formula found at $cellAddress")
}

/**
 * Adds an existing format record to the list of known formats.
 * This method does not add the record to the workbook! If the format is
 * not already in the workbook use [.insertFormat] instead.
 *
 * @param format the Format record to add
 */
     fun addFormat(format:Format):Int {
val ifmt = java.lang.Short.valueOf(format.ifmt)

 // Add it to the format record lookup
        formats.put(ifmt, format)

 // Add it to the format string lookup
        formatlookup.put(format.format!!.toUpperCase(), ifmt)

return format.ifmt.toInt()
}

/**
 * Gets a custom number format by its format ID.
 */
     fun getFormat(id:Int):Format {
return formats.get(java.lang.Short.valueOf(id.toShort()))
}

/**
 * associate default row/col size recs
 */
    internal fun setDefaultRowHeightRec(dr:DefaultRowHeight) {
this.drh = dr
}

/**
 * set Default row height in twips (1/20 of a point)
 */
    // should be a double as Excel units are 1/20 of what is
    // stored in defaultrowheight
    // e.g. 12.75 is Excel Units, twips = 12.75*20 = 256
    // (approx)
    // should expect users to use Excel units and target method
    // do the 20* conversion
     fun setDefaultRowHeight(t:Int) {
drh!!.setDefaultRowHeight(t)
}

/**
 * set Default col width for all worksheets in the workbook,
 *
 *
 * Default column width can also be set on individual worksheets
 */
     fun setDefaultColWidth(t:Int) {
val b = this.getWorkSheets()
for (i in b.indices)
{
b[i].defaultColumnWidth = t
}
}

/**
 * sets the selected worksheet
 */
     fun setSelectedSheet(bs:Boundsheet) {
val bsx = this.getWorkSheets()
for (t in bsx.indices)
{
if (bsx[t] != bs)
bsx[t].setSelected(false)
}
this.win1!!.setCurrentTab(bs)
}

/**
 * for those cases where a formula calculation adds a new string rec
 * need to explicitly set lastFormula before calling addRecord
 *
 * @param f
 */
     fun setLastFormula(f:Formula) {
lastFormula = f
}

/**
 * Associate a record with its containers and related records.
 */
    override fun addRecord(rec:BiffRec, addtorec:Boolean):BiffRec {
var addtorec = addtorec
val opc = rec.opcode
rec.streamer = streamer
rec.workBook = this

var bs:Boundsheet? = null
var lbplypos:Long? = 0L

 // get the relevant Boundsheet for this rec
        if (rec is Bof)
{
if (this.getFirstBof() == null)
this.setFirstBof(rec)
if (bofct == eofct)
{ // not a chart or other non Sheet Bof
this.setLastBOF(rec)
}
if (rec.isChartBof)
{
inChartSubstream = true
}
}

if (this.lastBOF == null)
Logger.logWarn("WorkBook: NULL Last BOF")
var lb = this.lastBOF!!.lbPlyPos
if (!this.lastBOF!!.isValidBIFF8)
lb += 8
lbplypos = lb // use last

bs = getSheetFromRec(rec, lbplypos)
if (bs != null)
lbplypos = bs.lbPlyPos

if (bs != null)
{ // &&){
lastbound = bs
if (addtorec)
rec.setSheet(bs)// we don't include Bof or other Book-recs
 // because it lives in the Streamer recvec
            if (!copying)
{
if (lastFormula != null && opc == XLSConstants.STRINGREC)
{
lastFormula!!.addInternalRecord(rec)
}
else if (lastFormula != null && opc == XLSConstants.ARRAY)
{
lastFormula!!.addInternalRecord(rec)
}
else if (rec.isValueForCell)
{
if (currchart == null)
bs.addCell(rec as CellRec)

}
}/*
             * ((rec.isValueForCell())
             * &&
             */
}

if (inChartSubstream)
{
if (currchart == null)
{
if (rec.opcode == XLSConstants.CHART)
{
chartVect.add(rec)
bs?.addChart(rec as Chart)
currchart = rec as Chart
currchart!!.setPreRecords(chartTemp)
chartTemp = ArrayList() // clear out
}
else
{
chartTemp.add(rec)
}
}
else
{
currchart!!.addInitialChartRecord(rec)
if (rec.opcode == XLSConstants.EOF)
{
currchart!!.initChartRecords() // finished
currchart = null
inChartSubstream = false
}
}
addtorec = false
}

 // Rows, valrecs, dbcells, and muls are stored in the row,
        // not the byte streamer
        if (opc == XLSRecord.DBCELL || opc == XLSRecord.ROW
|| rec.isValueForCell || opc == XLSRecord.MULRK
 /* || opc==MULBLANK */
                || opc == XLSConstants.CHART || opc == XLSRecord.FILEPASS
|| opc == XLSRecord.SHRFMLA || opc == XLSRecord.ARRAY
|| opc == XLSRecord.STRINGREC)
{
addtorec = false
}

 // add it to the record stream
        if (addtorec)
{

if (lbplypos.toLong() > 0)
streamer.addRecord(rec)
else
streamer.records.add(rec)
}

when (opc) {
XLSConstants.AUTOFILTER -> bs!!.autoFilters.add(rec)

XLSConstants.CONDFMT -> {
bs!!.conditionalFormats!!.add(rec)
(rec as Condfmt).initializeReferences()
}

XLSConstants.CF -> {
val cfmt = bs!!.conditionalFormats!![bs.conditionalFormats!!.size - 1] as Condfmt
cfmt.addRule(rec as Cf)
}

XLSConstants.MERGEDCELLS -> {
bs!!.addMergedCellsRec(rec as Mergedcells)
this.addMergedcells(rec)
}

 // give protection records to the relevant ProtectionManager
            XLSConstants.PASSWORD, XLSConstants.PROTECT, XLSConstants.PROT4REV, XLSConstants.OBJPROTECT, XLSConstants.SCENPROTECT, XLSConstants.FEATHEADR -> {
val manager:ProtectionManager
if (bs != null)
manager = bs.protectionManager
else
manager = this.protectionManager
manager.addRecord(rec)
}

XLSConstants.DVAL -> if (bs != null)
bs.dvalRec = rec as Dval

XLSConstants.DV -> if (bs != null)
{
if (bs.dvalRec != null)
{
bs.dvalRec!!.addDvRec(rec as Dv)
}
}

XLSConstants.INDEX -> {
val id = rec as Index
id.indexNum = indexnum++
indexes.add(indexes.size, id)
this.lastINDEX = id
if (bs == null)
{
Logger.logWarn("ERROR: WorkBook.addRecord( Index ) error: BAD LBPLYPOS.  The wrong LB:$lbplypos")
try
{
bs = this.getWorkSheetByNumber(indexnum - 1)
Logger.logInfo(" The RIGHT LB:" + bs.lbPlyPos)
}
catch (e:WorkSheetNotFoundException) {
Logger.logInfo("problem getting WorkSheetByNumber: $e")
}

}
bs!!.sheetIDX = id
}

XLSConstants.ROW -> {
val rw = rec as Row
bs?.addRowRec(rw)
}

XLSConstants.FORMULA -> {
this.addFormula(rec as Formula)
lastFormula = rec
}

XLSConstants.ARRAY -> {
val arr = rec as Array
bs?.addArrayFormula(arr)
arr.parentRec = lastFormula // [BugTracker 1869] link array
}

 /*
             * case SHRFMLA : done in shrfmla.init
             * Shrfmla form = (Shrfmla) rec;
             * try{ // throws exceptipon during pullparse
             * form.setHostCell( lastFormula );
             * }catch(Exception e){;}
             * break;
             */

            XLSConstants.DATE1904 -> if ((rec as NineteenOhFour).is1904)
{
this.dateFormat = DateConverter.DateFormat.LEGACY_1904
}

 /*
             * case PALETTE : // palette now correctly read into
             * COLORTABLE
             * this.pal = (Palette) rec;
             * break;
             */

            XLSConstants.HLINK -> {
val hl = rec as Hlink
this.addHlink(hl)
}

XLSConstants.DSF -> {
val dsf = rec as Dsf
if (dsf.fDSF == 1)
{
Logger.logErr("DOUBLE STREAM FILE DETECTED!")
Logger.logErr("  OpenXLS is compatible with Excel 97 and above only.")
throw WorkBookException(
"ERROR: DOUBLE STREAM FILE DETECTED!  OpenXLS is compatible with Excel 97 + only.",
WorkBookException.DOUBLE_STREAM_FILE)
}
}

XLSConstants.GUTS -> if (bs != null)
{
bs.guts = rec as Guts
}

XLSConstants.DBCELL -> {}

XLSConstants.BOF -> {
if (DEBUGLEVEL > 5)
Logger.logInfo("BOF:$bofct - $rec")
if (eofct == bofct)
{
bs?.setBOF(rec as Bof)
}
bofct++
}

XLSConstants.EXTERNSHEET -> myexternsheet = rec as Externsheet

XLSConstants.DEFCOLWIDTH -> if (bs != null)
{
bs.defColWidth = rec as DefColWidth
}

XLSConstants.EOF -> {
this.lasteof = rec as Eof
eofct++
if (eofct == bofct)
{
bs?.setEOF(rec)
eofct--
bofct--
}
}

XLSConstants.SELECTION // only used for Recvec index
 -> bs!!.lastselection = rec as Selection

XLSConstants.COUNTRY -> {
 // Added to save position of 1st bound sheet, which is 1
                // record
                // before COUNTRY RECORD (= 2 before 1st SUPBOOK record -
                // true in all cases?)
                countryRec = rec as XLSRecord
 // USA=1, Canada=1, Japan=81, China=86, Thailand= 66, Korea=
                // 82, India=91 ...
                this.defaultLanguage = (rec as Country).defaultLanguage
}

XLSConstants.SUPBOOK // KSC: must store ordinal positions of SupBooks, for
 -> {
 // adding Externsheets
                supBooks.add(rec)
if (myADDINSUPBOOK == null)
{ // see if this is the ADD-IN SUPBOOK
 // rec
                    val sb = rec as Supbook
if (sb.isAddInRecord)
myADDINSUPBOOK = sb
}
}

XLSConstants.BOUNDSHEET -> {
val sh = rec as Boundsheet

 /*
                 * Here we need to set the selected variable,
                 * but not mess with selected tabs
                 * when all of the sheets aren't in the book yet.
                 * -jm
                 */
                var ctab = 1 // default- select 1st sheet if no Windows1 record
if (win1 != null)
{// Windows1 record is optional 20101004
 // TestCorruption.TestNPEOnOpen
                    ctab = win1!!.currentTab
}
val shts = sheetVect.size
if (ctab == shts)
sh.selected = true

this.addWorkSheet(sh.lbPlyPos, sh)
}

XLSConstants.MULRK -> {
val mul = rec as Mulrk
val xit = mul.recs.iterator()
while (xit.hasNext())
{
this.addRecord(xit.next() as Rk, false)
}
}

XLSConstants.SST -> this.sharedStringTable = rec as Sst

XLSConstants.EXTSST -> (rec as Extsst).setSst(this.sharedStringTable!!)

XLSConstants.SXSTREAMID -> this.ptstream.add(rec) // Pivot Stream

XLSConstants.SXVS, XLSConstants.DCONREF, XLSConstants.DCONNAME, XLSConstants.DCONBIN -> try
{
val sid = ptstream.get(ptstream.size - 1) as SxStreamID
sid.addSubrecord(rec)
}
catch (e:Exception) {}

XLSConstants.SXVIEW -> addPivotTable(rec as Sxview) // Pivot Table View ==Top-level

 // all* possible records associated with SxView (=PivotTable
            // View) (*hopefully)
            XLSConstants.SXVD,
 // case SXVI: // subrecords of SxVD
                // case SXVDEX:
            XLSConstants.SXIVD, XLSConstants.SXPI, XLSConstants.SXDI, XLSConstants.SXLI, XLSConstants.SXEX, XLSConstants.SXVIEWEX9, XLSConstants.QSISXTAG -> try
{
val sx = this.ptViews.values
.toTypedArray()[this.ptViews.size - 1] as Sxview
sx.addSubrecord(rec)
}
catch (e:Exception) {}

XLSConstants.TABID -> this.tabID = rec as TabID

XLSConstants.NAME -> addName(rec as Name)

XLSConstants.CALCMODE -> this.calcmoderec = rec as CalcMode

XLSConstants.WINDOW1 -> this.win1 = rec as Window1

XLSConstants.WINDOW2 -> if (bs != null)
bs.window2 = rec as Window2

XLSConstants.SCL // scl is for zoom
 -> if (bs != null)
bs.scl = rec as Scl

XLSConstants.PANE -> if (bs != null)
bs.pane = rec as Pane
XLSConstants.EXCEL2K -> xl2k = rec

XLSConstants.PHONETIC ->
 // TODO: this isn't necessary anymore! look at and remove
                if (this.currdrw != null)
{
this.currdrw!!.mystery = rec as Phonetic
}

XLSConstants.MSODRAWINGGROUP -> {
if (msodg == null)
msodg = rec as MSODrawingGroup
msodgMerge.add(rec)
}

XLSConstants.MSODRAWING -> {
rec.setSheet(bs)
if (msodg != null)
msodg!!.addMsodrawingrec(rec as MSODrawing)
}

XLSConstants.COLINFO -> bs!!.addColinfo(rec as Colinfo)

XLSConstants.USERSVIEWBEGIN -> this.usersview = rec as Usersviewbegin

XLSConstants.WSBOOL -> if (bs != null)
{
bs.wsBool = rec as WsBool
}

 // Handle continue records which are actually masked Mso's
            XLSConstants.CONTINUE -> if ((rec as Continue).maskedMso != null)
{
rec.maskedMso!!.setSheet(bs)
if (msodg != null)
msodg!!.addMsodrawingrec(rec.maskedMso)
}

XLSConstants.XF -> try
{
this.addXf(rec as Xf)
}
catch (e:Exception) {
 // throws exceptions during PullParse
                }

}// formula to it's parent formula
 // rec
 // record for a Pivot Table
 // do what???io.starter.toolkit.Logger.log("PROBLEM with
 // MSODG!");
 // DO NOTHING

 // finish up
        rec.setIndex(lastINDEX)
rec.setXFRecord()
return rec
}

/**
 * Dec 15, 2010
 *
 * @param rec
 * @return
 */
    override fun getSheetFromRec(rec:BiffRec, lbplypos:Long?):Boundsheet? {
var bs:Boundsheet? = null

if (rec.sheet != null)
{
bs = rec.sheet
}
else if (lbplypos != null)
{
bs = this.getWorkSheet(lbplypos)
}
else
bs = lastbound
return bs
}

/**
 * set the Debug level
 */
    override fun setDebugLevel(i:Int) {
this.DEBUGLEVEL = i
}

/**
 * get the cell by the following String Pattern
 *
 *
 * BiffRec c = getCell("SheetName!C17");
 */
    @Throws(CellNotFoundException::class, WorkSheetNotFoundException::class)
 fun getCell(cellname:String):BiffRec {
val semi = cellname.indexOf("!")
val sname = cellname.substring(0, semi)
val cname = cellname.substring(semi + 1)
return getCell(sname, cname)
}

/**
 * get the cell by the following String Pattern
 *
 *
 * BiffRec c = getCell("Sheet1", "C17");
 */
    @Throws(CellNotFoundException::class, WorkSheetNotFoundException::class)
 fun getCell(sheetname:String, cellname:String):BiffRec {
var cellname = cellname
cellname = cellname.toUpperCase()
try
{
val bs = this.getWorkSheetByName(sheetname)
return bs.getCell(cellname) ?: throw CellNotFoundException("$sheetname:$cellname")
}
catch (a:WorkSheetNotFoundException) {
throw WorkSheetNotFoundException("$sheetname not found")
}
catch (e:NullPointerException) {
throw CellNotFoundException("$sheetname:$cellname")
}

}

/**
 * add a Boundsheet to the WorkBook
 */
    private fun addWorkSheet(lbplypos:Long?, sheet:Boundsheet?) {
if (sheet == null)
{
Logger.logWarn("WorkBook.addWorkSheet() attempting to add null sheet.")
return
}
else
{
this.lastbound = sheet
if (DEBUGLEVEL > 10)
Logger.logInfo("Workbook Adding Sheet: " + sheet.sheetName
+ ":" + lbplypos)
workSheets.put(lbplypos!!, sheet)
sheetVect.add(sheet)
}
}

/**
 * @param n
 * @return
 */
     fun removeName(n:Name):Boolean {
if (this.names.contains(n))
{
this.names.remove(n)
if (n.getItab().toInt() != 0)
{
try
{
this.getWorkSheetByNumber(n.getItab() - 1)
.removeLocalName(n)
}
catch (e:WorkSheetNotFoundException) {}

}
else
{
this.bookNameRecs.remove(n.toString().toUpperCase()) // case-insensitive
}
}
val al = n.ilblListeners
this.orphanedPtgNames.addAll(al)

this.updateNameIlbls()
return this.streamer.removeRecord(n)
}

/**
 * Add a sheet-scoped name record to the boundsheet
 *
 *
 * Note this is not that primary repository for names, it just contains the name records
 * that are bound to this book, adding them here will not add them to the workbook;
 *
 * @param bookNameRecs
 */
     fun addLocalName(name:Name) {
bookNameRecs.put(name.nameA, name)
}

/**
 * Remove a sheet-scoped name record from the boundsheet.
 *
 *
 * Note this is not that primary repository for names, it just contains the name records
 * that are bound to this book, removing them here will not remove them completely from the workbook.
 *
 *
 * In order to do that you will need to call book.removeName
 *
 * @param bookNameRecs
 */
     fun removeLocalName(name:Name) {
bookNameRecs.remove(name.nameA)
}

/**
 * After any changes in the name records
 * this method needs to be called in order to
 * update ilbl records
 */
     fun updateNameIlbls() {
for (i in names.indices)
{
val n = names.get(i) as Name
n.updateIlblListeners()
}
}

/**
 * remove a Boundsheet from the WorkBook
 */
     fun removeWorkSheet(sheet:Boundsheet) {

var sheetNum = sheet.sheetNum
 // remove the sheet
        // automatically deletes Named ranges scoped to the sheet
        val namesOnSheet = sheet.allNames
for (i in namesOnSheet.indices)
this.removeName(namesOnSheet[i])

 // Remove Externsheet ref before removing sheet
        // update any Externsheet references...
        try
{
val ext = this.externSheet
ext?.removeSheet(sheet.sheetNum)
}
catch (e:WorkSheetNotFoundException) {
Logger.logInfo("could not update Externsheet reference from "
+ sheet.toString() + " : " + e.toString())
}

sheet.removeAllRecords()
streamer.removeRecord(sheet)
workSheets.remove(sheet.lbPlyPos)
sheetVect.remove(sheet)
 // we need to reset the lastbound for adding new worksheets.
        // Currently assume it is
        // the last one in the vector.
        if (sheetVect.size > 0)
{
lastbound = sheetVect[sheetVect.size - 1] as Boundsheet
lastBOF = lastbound!!.myBof
}

 // decrement the tab ids...
        this.tabID!!.removeRecord()
this.updateScopedNamedRanges()

 // update wb chart cache - remove charts referenced by
        // deleted sheet
        for (i in chartVect.size - 1 downTo 0)
{
if ((this.chartVect[i] as Chart).sheet == sheet)
this.chartVect.removeAt(i)
}

if (this.numWorkSheets == 0)
return  // empty book
try
{ // set the next sheet selected...
while (sheetNum <= this.numWorkSheets)
{
val s2 = this.getWorkSheetByNumber(sheetNum++)
s2.setSelected(true)
if (!s2.hidden)
break

}
}
catch (e:WorkSheetNotFoundException) {
try
{
val s2 = this.getWorkSheetByNumber(0)
s2.setSelected(true)
}
catch (ee:Exception) {
throw WorkBookException(
"Invalid WorkBook.  WorkBook must contain at least one Sheet.",
WorkBookException.RUNTIME_ERROR)
}

}

}

/**
 * Updates all the name records in the workbook that are bound to a
 * worksheet scope (as opposed to a workbook scope).  Name records use
 * their own non-externsheet based sheet references, so need to be modified
 * whenever a sheet delete (or non-last sheet insert) operation occurs
 */
    private fun updateScopedNamedRanges() {
for (i in sheetVect.indices)
{
(sheetVect[i] as Boundsheet).updateLocalNameReferences()
}
}

/**
 * returns the Boundsheet identified by its
 * offset to the BOF record indicating the
 * start of the Boundsheet data stream.
 *
 *
 * used internally to access the Sheets to
 * ensure that the lbplypos is correct -- essential
 * to proper operation of XLS file.
 *
 * @param Long lbplypos of Boundsheet
 */
    private fun getWorkSheet(lbplypos:Long?):Boundsheet {
return workSheets.get(lbplypos!!)
}

/**
 * returns the Boundsheet with the specific name
 *
 * @param String name of Boundsheet
 */
    @Throws(WorkSheetNotFoundException::class)
 fun getWorkSheetByName(bstr:String?):Boundsheet {
var bstr = bstr
try
{
if (bstr!!.startsWith("'") || bstr.startsWith("\""))
bstr = bstr.substring(1, bstr.trim{ it <= ' ' }.length - 1)
val bs = sheetVect.iterator()
while (bs.hasNext())
{
val bsi = bs.next() as Boundsheet
val bsin = bsi.sheetName
 // TODO: check if we can have dupe names different case
                if (bsin.equals(bstr, ignoreCase = true))
return bsi
}
}
catch (ex:Exception) {
Logger.logWarn("WorkBook.getWorkSheetByName failed: $ex")
}

throw WorkSheetNotFoundException(
"Worksheet $bstr not found in $this")
}

/**
 * returns the Boundsheet with the specific Hashname
 *
 * @param String hashname of Boundsheet
 */
    @Throws(WorkSheetNotFoundException::class)
 fun getWorkSheetByHash(s:String):Boundsheet? {
val bs = this.getWorkSheets()
for (i in bs.indices)
{
if (bs[i].getSheetHash().equals(s, ignoreCase = true))
return bs[i]
}
return null
}

/**
 * returns the Boundsheet at the specific index
 *
 * @param int index of Boundsheet
 */
    @Throws(WorkSheetNotFoundException::class)
 fun getWorkSheetByNumber(i:Int):Boundsheet {
var bs:Boundsheet? = null
try
{
bs = sheetVect[i] as Boundsheet
}
catch (e:ArrayIndexOutOfBoundsException) {}

if (bs == null)
{ // External Sheet Ref NOT FOUND
throw WorkSheetNotFoundException(
"$i not found")
}
return bs
}

/**
 * returns the boundsheets for this book as an array
 */
    internal fun getWorkSheets():Array<Boundsheet> {
val ret = arrayOfNulls<Boundsheet>(sheetVect.size)
return sheetVect.toTypedArray() as Array<Boundsheet>
}

/**
 * set the last BOF read in the stream
 */
    internal fun setLastBOF(b:Bof) {
lastBOF = b
}
/** return the last BOF read in the stream  */
    // Bof lastBOF{return lastBOF;}

    /**
 * get a handle to the first BOF to perform offset functions which don't know where the
 * start of the file is due to the compound file format.
 *
 *
 * Referred to in Boundsheet as the 'lbPlyPos', this
 * is the position of the BOF for the Boundsheet relative
 * to the *first* BOF in the file (the firstBOF of the WorkBook)
 *
 * @see Boundsheet
 */
    override fun setFirstBof(b:Bof) {
firstBOF = b
}

internal fun getFirstBof():Bof? {
return firstBOF
}

override fun toString():String {
return this.fileName
}

/**
 * Returns whether the sheet selection tabs should be shown.
 */
     fun showSheetTabs():Boolean {
return win1!!.showSheetTabs()
}

/**
 * Sets whether the sheet selection tabs should be shown.
 */
     fun setShowSheetTabs(show:Boolean) {
win1!!.setShowSheetTabs(show)
}

/**
 * set the first visible tab
 */
     fun setFirstVisibleSheet(bs2:Boundsheet) {
win1!!.setFirstTab(bs2.sheetNum)
}

 // Associate related records

    /**
 * return the XF record at the specified index
 */
     fun getXf(i:Int):Xf? {
return if (xfrecs.size < i - 1) null else xfrecs[i] as Xf
}

/**
 * InsertXF inserts an XF record into the workbook level stream,
 * For some reason, the addXf only puts it into an array that is never accessed
 * on output.  This may have a reason, so I am not overwriting it currently, but
 * let's check it out?
 */
    internal fun insertXf(x:Xf):Int {
val insertIdx = this.getXf(this.numXfs - 1)!!.recordIndex
 // perform default add rsec actions
        this.streamer.addRecordAt(x, insertIdx + 1)
this.addRecord(x, false) // updates xfrecs + formatcache
x.ixfe = x.idx
return x.idx
}

/**
 * internally used in preparation for reading an 2007 and above workbook
 */
     fun removeXfRecs() {
 // must keep the 1st xf rec as default
        for (i in xfrecs.size - 1 downTo 1)
{
val xf = xfrecs[i] as Xf
this.streamer.removeRecord(xf)
xfrecs.removeAt(i)

}
}

/**
 * TODO: Does this function as desired?   See comment for insertXf() above...
 * tracks existing xf recs, used when testing whether xfrec exists or not ...
 * -NR 1/06
 * ya should - called now from addRecord every time an xf record is added
 * NOTE: this is the only place addXf is called
 *
 * @param xf
 * @return
 */
    internal fun addXf(xf:Xf):Int {
xfrecs.add(xf)
xf.idx = xfrecs.size - 1 // flag that it's been added to records
this.updateFormatCache(xf) // links tostring of xf to xf rec for
 // updating/reuse purposes
        return xf.idx
}

/**
 * formatCache:
 * links tostring of xf to xf rec for updating/reuse purposes
 *
 * @param xf
 * @see FormatHandle.updateXf
 *
 * @see WorkBook.addXf
 */
     fun updateFormatCache(xf:Xf) {
if (xf.idx != -1)
{ // if this xf has been already added to the
 // workbook
            if (formatCache.containsValue(xf))
{ // xf signature has
 // changed/it's been updated
                val ii = formatCache.keys.iterator() // remove and
 // update below
                while (ii.hasNext())
{
val key = ii.next() as String
val x = formatCache.get(key) as Xf
if (x == xf)
{
formatCache.remove(key)
break
}
}
}
val formatStr = xf.toString()

if (!formatCache.containsKey(formatStr))
formatCache.put(formatStr, xf)
}
}

 fun setSubStream(s:ByteStreamer) {
this.streamer = s
}

/**
 * Write the contents of the WorkBook bytes to an OutputStream
 *
 * @param _out
 */
    override fun stream(_out:OutputStream):Int {
return streamer.streamOut(_out)
}

/**
 * Copies a complete boundsheet within the workbook
 *
 *
 * If a name exists that refers directly to this sheet then duplicate it, otherwise workbook scoped names are
 * not copied
 */
    @Throws(Exception::class)
 fun copyWorkSheet(SourceSheetName:String, NewSheetName:String) {
var origSheet:Boundsheet? = null
origSheet = this.getWorkSheetByName(SourceSheetName)
val chts = origSheet.charts // 20080630 KSC: Added
for (i in chts.indices)
{
val cxi = chts[i] as Chart
cxi.populateForTransfer()
}
val inbytes = origSheet.sheetBytes
this.addBoundsheet(inbytes, SourceSheetName, NewSheetName, null, false)
val bnd = this.getWorkSheetByName(NewSheetName)
 // handle moving the built-in name records. These handle
        // such items as print area, header/footer, etc
        val ns = this.getNames()
for (i in ns.indices)
{ // 20100404 KSC: take out +1?
if (ns[i].getItab().toInt() == origSheet.sheetNum + 1)
{
 // it's a built in record, move it to the new sheet
                val sheetnum = bnd.sheetNum
val xref = this.getExternSheet(true)!!
.insertLocation(sheetnum, sheetnum)
val n = ns[i].clone() as Name
n.setExternsheetRef(xref)
n.updateSheetReferences(bnd)
n.setSheet(bnd)
n.setItab((bnd.sheetNum + 1).toShort())
this.insertName(n)
}
}
}

/**
 * Inserts a newly created Name record into the correct location in the streamer.
 *
 * @param n
 */
     fun insertName(n:Name) {
var namepos = -1
val nmx = getNames()
if (nmx.size > 0)
{
if (nmx[nmx.size - 1].sheet != null)
{
namepos = nmx[nmx.size - 1].recordIndex
}
else
{
namepos = getExternSheet(true)!!.recordIndex + nmx.size
}
}
else
{
namepos = getExternSheet(true)!!.recordIndex
}
namepos++
streamer.addRecordToBookStreamerAt(n, namepos)
addRecord(n, false)
}

/**
 * Copies an existing Chart to another WorkSheet
 *
 * @param chartname
 * @param sheetname
 */
    @Throws(ChartNotFoundException::class, WorkSheetNotFoundException::class)
 fun copyChartToSheet(chartname:String, sheetname:String) {
val ct = this.getChart(chartname)
val sht = this.getWorkSheetByName(sheetname)
val bt = ct.chartBytes
sht.addChart(bt, chartname, ct.coords)
}

/**
 * Inserts a serialized boundsheet chart into the workboook
 */
     fun addChart(destChart:Chart, NewChartName:String, boundsht:Boundsheet):Chart {
destChart.workBook = this
destChart.setSheet(boundsht)
val recs = destChart.xlSrecs
for (x in recs.indices)
{
val rec = recs[x] as XLSRecord
rec.workBook = this
rec.setSheet(boundsht)
if (rec.opcode == XLSConstants.MSODRAWING)
{
addChartUpdateMsodg(rec as MSODrawing, boundsht)
continue
}
if (rec !is Bof)
 // TODO: error/problem with the BOF
                // record!!!
                rec.init()
if (rec is Dimensions)
destChart.setDimensions(rec)
try
{
(rec as GenericChartObject).parentChart = destChart
}
catch (e:ClassCastException) { // Scl, Obj and others are not
 // chart objects
            }

}
destChart.title = NewChartName
destChart.id = boundsht.lastObjId + 1 // track last obj id per sheet
 // ...
        this.chartVect.add(destChart)
boundsht.charts.add(destChart) // should really have two lists???
return destChart
}

/**
 * updates Mso (MSODrawingGroup + Msodrawing) records upon add/copy worksheet and add/copy charts
 * NOTE: this code is mainly garnered via trial and error, works
 *
 * @param mso Msodrawing record that is being added or copied
 * @param sht Boundsheet
 */
     fun addChartUpdateMsodg(mso:MSODrawing, sht:Boundsheet) {
if (msodg == null)
{
this.msoDrawingGroup = MSODrawingGroup
.prototype as MSODrawingGroup?
msodg!!.initNewMSODrawingGroup() // generate and add required records
 // for drawing records
        }
msodg!!.addMsodrawingrec(mso)
var hdr = msodg!!.getMsoHeaderRec(sht)
if (hdr != null && hdr != mso)
{ // already have a header rec
if (sht.charts.size > 0)
{
mso.makeNonHeader()
hdr.setNumShapes(hdr.getNumShapes() + 1)
}
}
else if (hdr == null)
{
mso.setIsHeader()
hdr = mso
}
this.updateMsodrawingHeaderRec(sht)
msodg!!.dirtyflag = true // flag to reset SPIDs on write
msodg!!.spidMax = ++lastSPID
msodg!!.updateRecord()
}

/**
 * JM -
 * Add the requisite records in the book streamer for the chart. \
 * Supbook, externsheet & msodrawinggroup
 *
 *
 * I think this is due to the fact that the referenced series are usually stored
 * in the fashon 'Sheet1!A4:B6' The sheet1 reference requires a supbook, though the
 * reference is internal.
 */
     fun addPreChart() {
this.addExternsheet()
if (msodg == null)
{
this.msoDrawingGroup = MSODrawingGroup
.prototype as MSODrawingGroup?
msodg!!.initNewMSODrawingGroup() // generate and add required records
 // for drawing records
        }

}

/**
 * remove an existing chart from the workbook
 * NOTE: STILL EXPERIMENTAL TESTS OK IN BASIC CIRCUMSTANCES BUT MUST BE TESTED FURTHER
 */
    @Throws(ChartNotFoundException::class)
 fun deleteChart(chartname:String, sheet:Boundsheet) {
val chart = this.getChart(chartname)
 // TODO: Update Dimensions record??
        val recs = chart.xlSrecs
 // first rec SHOULD BE MsoDrawing!!!
        try
{
val rec = recs[0] as MSODrawing
msodg!!.removeMsodrawingrec(rec, sheet, true) // also remove
 // associated Obj
            // record
        }
catch (e:Exception) {
Logger.logErr("deleteChart: expected Msodrawing record")
}

 /*
         * shouldn't be necessary to remove chart recs as they are
         * separated upon init of workbook and reassebmbled upon
         * write
         */
        this.removeChart(chartname)
}

/**
 * Inserts a serialized boundsheet into the workboook.
 *
 * @throws ClassNotFoundException
 * @throws IOException
 * @param    inbytes                original sheet bytes
 * @param    NewSheetName        new Sheet Name
 * @param    origWorkBookName    original WorkBook Name (nec. for resolving possible external references)     *
 * @boolean SSTPopulatedBoundsheet - a boundsheet that has all of it's sst data saved off in LabelSST records.
 * one would use this when moving a sheet from a workbook that was not an original or created from getNoSheetWorkBook.
 * Do not use this if the data already exists in the SST, you are just causing bloat!
 */
    @Throws(IOException::class, ClassNotFoundException::class)
 fun addBoundsheet(inbytes:ByteArray?, origSheetName:String, NewSheetName:String, origWorkBookName:String?, SSTPopulatedBoundsheet:Boolean) {
var destSheet:Boundsheet? = null
val bais = ByteArrayInputStream(inbytes!!)
val bufstr = BufferedInputStream(bais)
val o = ObjectInputStream(bufstr)
destSheet = o.readObject() as Boundsheet

if (destSheet != null)
{
this.addBoundsheet(destSheet, origSheetName, NewSheetName, origWorkBookName, SSTPopulatedBoundsheet)
}
}

/**
 * change the tab order of a boundsheet
 */
     fun changeWorkSheetOrder(bs:Boundsheet, idx:Int) {
 // reorder the sheet vector
        if (idx >= 0 && idx < sheetVect.size)
{
sheetVect.remove(bs)
sheetVect.add(idx, bs)
for (x in sheetVect.indices)
{
val bs1 = sheetVect[x] as Boundsheet
val udpatewin1 = bs1.selected()
if (udpatewin1)
bs1.setSelected(true)
}
}

var insertLoc = Integer.MAX_VALUE
 // remove the existing boundsheet records in the streamer
        for (i in sheetVect.indices)
{
val bound = sheetVect[i] as Boundsheet
val position = bound.recordIndex
insertLoc = Math.min(insertLoc, position)
streamer.removeRecord(sheetVect[i] as XLSRecord)
}
 // enter the boundsheet records back in the streamer in
        // correct order
        for (i in sheetVect.indices)
{
val bound = sheetVect[i] as Boundsheet
streamer.addRecordAt(bound, insertLoc + i)
}
}

/**
 * add a deserialized boundsheet  to this workbook.
 *
 * @param    bound                new (copied) sheet
 * @param    newSheetName        new sheetname
 * @param    origWorkBookName    original WorkBook Name (nec. for resolving possible external references)     *
 * @boolean SSTPopulatedBoundsheet - the boundsheet has all of it's sst data saved off in LabelSST records.
 * one would use this when moving a sheet from a workbook that was not an original or created from getNoSheetWorkBook.
 * Do not use this if the data already exists in the SST, you are just causing bloat!
 */
     fun addBoundsheet(bound:Boundsheet, origSheetName:String, newSheetName:String, origWorkBookName:String?, SSTPopulatedBoundsheet:Boolean) {
var newSheetName = newSheetName
bound.streamer = streamer
val old_allowdupes = this.isSharedupes
this.setDupeStringMode(ALLOWDUPES)

bound.mc.clear()
bound.workBook = this

 // Check if sheetname already exists!
        try
{
while (this.getWorkSheetByName(newSheetName) != null)
newSheetName = newSheetName + "Copy"
}
catch (we:WorkSheetNotFoundException) {
 /* good !!! */
        }

bound.sheetName = newSheetName

 // get a hold of the lbplypos number that we will need for
        // the new boundsheet
        var recvecOffset = streamer.recVecSize - 1
var x:XLSRecord? = null
if (lastbound != null)
{
try
{ // lastbound must be reset because other operations could
 // alter
                lastbound = this
.getWorkSheetByNumber(this.numWorkSheets - 1)
x = lastbound!!.sheetRecs[lastbound!!.sheetRecs.size - 1] as XLSRecord
}
catch (e:Exception) {}

}
else if (countryRec != null)
{
x = streamer.getRecordAt(countryRec!!.recordIndex) as XLSRecord
}
else
x = this.lasteof
 // last record is a junkrec. We are going to move that down
        // and put in the new BOF here
        // TODO: recvecOffset position when no sheets??????
        if (x!!.opcode != XLSConstants.EOF)
{
recvecOffset -= 1
}
var newloc = x.offset
 // modify the boundsheet rec for its new location/info/name
        var listenerpos = -1
var newoffset = -1
if (lastbound != null)
{
listenerpos = lastbound!!.recordIndex
newoffset = lastbound!!.offset + lastbound!!.length + 4
 // offset + reclen + headerlen
        }
else if (x != null)
{
listenerpos = x.recordIndex - 1 // account for +1 below
newoffset = x.offset + x.length + 4
}
else
{
listenerpos = recvecOffset - 1
newoffset = streamer.getRecordAt(recvecOffset).length + 4
}

 // put the serialized recs from localrecs into the normal
        // SheetRecs
        bound.setLocalRecs(FastAddVector()) // reset localrecs
val newrecs = bound.sheetRecs
val newbof = newrecs[0] as Bof

newloc += newbof.length + 4
newbof.offset = newloc
bound.setBOF(newbof)
this.addRecord(newbof, false)

recvecOffset += 1 // move it past that last Eof
newoffset = newloc + newbof.length + 4
lastbound = bound
 // insert the actual boundsheet record into the recvec
        streamer.addRecordAt(bound, listenerpos + 1)

this.addRecord(bound, false)

 // modify the TabID record to reflect new sheet
        tabID!!.addNewRecord()
recvecOffset = newbof.recordIndex

var tout = 0
copying = true

 // Add an externsheet ref for the new sheet
        if (this.myexternsheet == null)
this.addExternsheet()
try
{
val sheetref = this.numWorkSheets - 1
myexternsheet!!.insertLocation(sheetref, sheetref)
}
catch (e:Exception) {
Logger.logWarn("Adding new sheetRef failed in addBoundsheet()$e")
}

 // update the chart references + add to wb
        val chts = bound.charts
for (i in chts.indices)
{
val chart = chts[i] as Chart // obviously algorithm has
 // changed and chart is NOT
            // removed :) [discovered by
            // Shigeo/Infoteria/formatbroken273193.sce]
            // // 20080702 KSC: since it's
            // removed, don't inc index
            chart.updateSheetRefs(bound.sheetName, origWorkBookName)
this.chartVect.add(chart)
}

bound.lastObjId = 1 // see if resetting obj id helps in file open
 // errors; if so, must reset Note obj id's as
        // well ...

        /********** This loop handles Boundsheet records contained in the sheet level streamer, that is, not the valrecs  */
        val numrecs = newrecs.size
for (z in 1 until numrecs)
{
val xl = newrecs[z] as XLSRecord
this.addRecord(xl, false)
if (DEBUGLEVEL > 5)
try
{
Logger.logInfo("Copying: " + xl.toString() + ":"
+ newoffset + ":" + xl.length)
}
catch (e:Exception) {}

if (xl is Codename)
{
xl.setName(newSheetName)
}
else if (xl is Name)
{
 // Name records specify data ranges -- update to point to
                // new sheet
                val refnum = myexternsheet!!.getcXTI()
xl.setExternsheetRef(refnum)
}
else if (xl is Cf)
{ // must check Conditional Format
 // formula refs and handle any
                // external references
                try
{
updateFormulaPtgRefs(xl
.formula1, origSheetName, newSheetName, origWorkBookName)
 // NOTE: FORMULA2 can be null -- TODO: should check here
                    updateFormulaPtgRefs(xl
.formula2, origSheetName, newSheetName, origWorkBookName)
}
catch (e:Exception) {}

}
else if (xl.opcode == XLSConstants.OBJ)
{
(xl as Obj).objId = bound.lastObjId++
}
else if (xl.opcode == XLSConstants.MSODRAWING || xl.opcode == XLSConstants.CONTINUE && (xl as Continue).maskedMso != null)
{ // 20100510
 // KSC:
                // handle
                // masked
                // mso's
                val mso:MSODrawing?
if (xl.opcode == XLSConstants.MSODRAWING)
mso = xl as MSODrawing
else
mso = (xl as Continue).maskedMso
if (msodg == null)
{
this.msoDrawingGroup = MSODrawingGroup
.prototype as MSODrawingGroup?
msodg!!.initNewMSODrawingGroup() // generate and add required
 // records for drawing
                    // records
                    msodg!!.addMsodrawingrec(mso) // only add when msodg is null
 // b/c otherwise it's added
                    // via the addRecord
                    // statement above
                }
if (mso!!.getImageIndex() > 0)
{ // add image bytes as well, if
 // any
                    val im = bound
.getImageByMsoIndex(mso.getImageIndex())
val idx = msodg!!.addImage(im!!.imageBytes, im
.imageType, false)
bound.imageMap!![im] = Integer.valueOf(idx) // 20100518
 // KSC:
                    // makes
                    // more
                    // sense?
                    // im.getImageIndex()));
                    // // add
                    // new image
                    // to map
                    // and link
                    // to actual
                    // imageIndex
                    // - moved
                    // from
                    // above
                    if (idx != mso.getImageIndex())
mso.updateImageIndex(idx)
}
mso.spid = this.lastSPID
msodg!!.spidMax = ++this.lastSPID
 // resets drawing id's - necessarily correct?
                // msodg.dirtyflag= true; // flag to reset SPIDs on write
            }
xl.setOffset(newoffset)
tout += xl.length
newoffset += xl.length
}
if (msodg != null)
{// Moved from above so don't udpate at every mso
 // addition
            // necessary? all mso sub-records on the sheet should have
            // stayed the same ...this.updateMsodrawingHeaderRec(bound);
            msodg!!.updateRecord()
}
/*************** END handling of boundsheet streamer records  */

        /*************** HANDLE Formats + PtgRefs in Cell Records  */
        updateTransferedCellReferences(bound, origSheetName, origWorkBookName)

 // associate the records in the sheet
        this.isSharedupes = old_allowdupes

if (SSTPopulatedBoundsheet)
{
 // bring over the sst
            val sst = this.sharedStringTable
val b = bound.cells
for (i in b.indices)
{
b[i].workBook = this
if (b[i].opcode == XLSConstants.LABELSST)
{
val s = b[i] as Labelsst
s.insertUnsharedString(sst)
}
}
}

if (this.numWorkSheets > 1)
bound.setSelected(false)
else
bound.setSelected(true)

if (DEBUGLEVEL > 5)
Logger.logInfo("changesize for  new boundsheet: "
+ bound.sheetName + ": " + tout)
copying = false
}

/**
 * traverses all rows and their associated cells in the newly transfered sheet,
 * ensuring formula/cell references and format references are correctly transfered
 * into the current workbook
 *
 * @param bound source sheet
 */
    private fun updateTransferedCellReferences(bound:Boundsheet, origSheetName:String, origWorkBookName:String?) {
val localFonts = this.fontRecsAsXML as HashMap<*, *>
val boundFonts = bound.transferFonts // ALL fonts in the source
 // workbook
        val localXfs = this.xfrecsAsString as HashMap<*, *>
val boundXfs = bound.transferXfs
 // Set the workbook on all the cells
        val rows = bound.rows
for (i in rows.indices)
{
rows[i].workBook = this
if (rows[i].ixfe != this.defaultIxfe)
transferFormatRecs(rows[i], localFonts, boundFonts, localXfs, boundXfs) // 20080709
 // KSC:
            // handle
            // default
            // ixfe
            // for
            // row
            val rowcells = rows[i].cells.iterator()
var aMul:Mulblank? = null
var c:Short = 0
while (rowcells.hasNext())
{
val b = rowcells.next() as BiffRec
if (b.opcode == XLSConstants.MULBLANK)
{
if (aMul == b)
c++
else
{
aMul = b as Mulblank
c = aMul.getColFirst().toShort()
}
aMul.setCurrentCell(c)
}
b.workBook = this // Moved to before updateFormulaPtgRefs
 // [BugTracker 1434]
                if (b is Formula)
{ // Examine Ptg Refs to handle
 // external sheet references not
                    // contained in this workbook
                    updateFormulaPtgRefs(b, origSheetName, bound
.sheetName, origWorkBookName)
if (b.shared != null)
b.shared!!.workBook = this

}
 // 20080226 KSC: transfer format, fonts and xf here instead
                // of populateWorkbookWithRemoteData()
                transferFormatRecs(b, localFonts, boundFonts, localXfs, boundXfs)
}
}
 // 20080226 KSC: handle xf's for columns
        for (co in bound.colinfos)
{
transferFormatRecs(co, localFonts, boundFonts, localXfs, boundXfs)
}
val c = bound.charts
for (i in c.indices)
{
val cht = c[i] as Chart
val fontrefs = cht.fontxRecs
for (x in fontrefs.indices)
{
val fontx = fontrefs[x] as Fontx
var fid = fontx.ifnt
if (fid > 3)
{
fid = bound.translateFontIndex(fid, localFonts)
fontx.ifnt = fid
}
}
}
}

/**
 * examine all Ptg's referenced by this formula, looking for hanging or missing sheet references
 * if found, sets sheet reference to the current sheet (TODO: a better way?)
 *
 * @param f Formula Rec
 */
    private fun updateFormulaPtgRefs(f:Formula?, origSheetName:String, newSheetName:String, origWorkBookName:String?) {
try
{
if (f == null)
return  // 20100222 KSC
f.populateExpression()
val p = f.cellRangePtgs
for (k in p.indices)
{
if (p[k] is PtgRef)
{
val ptg = p[k] as PtgRef
try
{
if (ptg !is PtgArea3d || ptg
.firstSheet == ptg.lastSheet)
{
val sheetName = ptg.sheetName
if (sheetName == origSheetName)
ptg.sheetName = newSheetName
ptg.addToRefTracker()
 /*
                             * changed to use above. don't understand this:
                             * if (!sheetName.equals(origSheetName)) {
                             * this.getWorkSheetByName(ptg.getSheetName());
                             * ptg.setSheetName(newSheetName);
                             * } else
                             * ptg.setSheetName(newSheetName);
                             */
                        }
else
{ // uncommon case of two sheet range
val pref = ptg
 // this.getWorkSheetByName(pref.getFirstPtg().getSheetName());
                            // don't understand this
                            // this.getWorkSheetByName(pref.getLastPtg().getSheetName());
                            ptg.location = ptg.toString() // reset ixti if
 // nec.
                        }
}
catch (we:WorkSheetNotFoundException) {
Logger.logWarn("External Reference encountered upon updating formula references:  Worksheet Reference Found: " + ptg.sheetName!!)
ptg.setExternalReference(origWorkBookName)
}

}
else if (p[k] is PtgExp)
{
val ptgexp = p[k] as PtgExp
try
{
val pe = ptgexp.convertedExpression // will fail
 // if
                        // ShrFmla
                        // hasn't
                        // been
                        // input yet
                        for (j in pe.indices)
{
if (pe[j] is PtgRef)
{
val ptg = pe[j] as PtgRef
try
{
if (ptg is PtgArea3d)
{ // PtgRef3d,
 // etc.
                                        this.getWorkSheetByName(ptg
.sheetName)
ptg.location = ptg.toString() // reset
 // ixti
                                        // if
                                        // nec.
                                    }
 // otherwise, we're good
                                }
catch (we:WorkSheetNotFoundException) {
Logger.logWarn("External References Not Supported:  UpdateFormulaReferences: External Worksheet Reference Found: " + ptg.sheetName!!)
ptg.setExternalReference(origWorkBookName)
}

}
}
}
catch (e:Exception) {
 // if links to "main" ShrFmla, won't be set yet and will
                        // give exception - see Shrfmla WorkBook.addRecord
                    }

}
}
}
catch (e:Exception) {
Logger.logErr("WorkBook.updateFormulaRefs: error parsing expression: $e")
}

}

/**
 * given a record in an previously external workbook, ensure that xf and font records
 * are correctly input into the current workbook and that the pointers are correctly updated
 *
 * @param b          BiffRec
 * @param localFonts HashMap of string version of all fonts, font nums in current workbook
 * @param boundFonts List of string version of all fonts, font nums in external workbook
 * @param localXfs   HashMap of string version of all xfs, xf nums in current workbook
 * @param boundXfs   List of string version of all xfs, xf nums in external workbook
 */
    private fun transferFormatRecs(b:BiffRec, localFonts:HashMap<String, Int>, boundFonts:List<*>?, localXfs:HashMap<String, Int>, boundXfs:List<*>) {
val oldXfNum = b.ixfe
val localNum = this
.transferFormatRecs(oldXfNum, localFonts, boundFonts, localXfs, boundXfs)
if (localNum != -1)
b.ixfe = localNum
}

/**
 * given a record in an previously external workbook, ensure that xf and font records
 * are correctly input into the current workbook and that the pointers are correctly updated
 *
 * @param b          BiffRec
 * @param localFonts HashMap of string version of all fonts, font nums in current workbook
 * @param boundFonts List of string version of all fonts, font nums in external workbook
 * @param localXfs   HashMap of string version of all xfs, xf nums in current workbook
 * @param boundXfs   List of string version of all xfs, xf nums in external workbook
 */
    private fun transferFormatRecs(oldXfNum:Int, localFonts:HashMap<String, Int>, boundFonts:List<*>?, localXfs:HashMap<String, Int>, boundXfs:List<*>):Int {
var localNum = -1
if (boundXfs.size > oldXfNum)
{// if haven't populatedForTransfer i.e.
 // haven't opted to transfer formats
            // ...
            val origxf = boundXfs[oldXfNum] as Xf // clone xf so modifcations
 // don't affect original
            if (origxf != null)
{
/** FONT  */
                // must handle font first in order to create xf below
                // see if referenced xf + fonts are already in workbook; if
                // not, add
                val localfNum:Int
 // check to see if the font needs to be added
                var fnum = origxf.ifnt.toInt()
if (fnum > 3)
fnum--
val thisFont = boundFonts!![fnum] as Font
val xmlFont = "<FONT><" + thisFont.xml + "/></FONT>"
val fontNum = localFonts[xmlFont]
if (fontNum != null)
{ // then get the fontnum in this book
localfNum = fontNum.toInt()
}
else
{ // it's a new font for this workbook, add it in
localfNum = this.insertFont(thisFont) + 1
localFonts[xmlFont] = Integer.valueOf(localfNum)
}

/** XF  */
                val localxf = FormatHandle
.cloneXf(origxf, origxf.font, this) // clone xf so
 // modifcations
                // don't
                // affect
                // original
                // input "local" versions of format and font

                /** FORMAT  */
                val fmt = origxf.format // number format - is null if
 // format is general ...
                if (fmt != null)
 // add if necessary
                    localxf.formatPattern = fmt.format // adds new
 // format
                // pattern if
                // not found
                localxf.setFont(localfNum)

 // now check out to see if xf needs to be added
                val xmlxf = localxf.toString()
val xfNum = localXfs[xmlxf]
if (xfNum == null)
{ // insert it into the book
localNum = this.insertXf(localxf)
localXfs[xmlxf] = Integer.valueOf(localNum)
}
else
 // already exists in the destination
                    localNum = xfNum.toInt()

}
}
return localNum
}

 fun setStringEncodingMode(mode:Int) {
this.sharedStringTable!!.setStringEncodingMode(mode)
}

 fun setDupeStringMode(mode:Int) {
if (mode == ALLOWDUPES)
this.isSharedupes = false
else if (mode == SHAREDUPES)
this.isSharedupes = true
}

/**
 * Returns a Chart Handle
 *
 * @return ChartHandle a Chart in the WorkBook
 */
    @Throws(ChartNotFoundException::class)
 fun getChart(chartname:String):Chart {
val cv = this.chartVect
var cht:Chart? = null
 // Get by MSODG Drawing Name
        for (x in cv.indices)
{
cht = cv[x] as Chart
val titlemso = cht.msodrawobj
if (titlemso != null)
{
val mson = titlemso.name // shapeName;
if (mson!!.equals(chartname, ignoreCase = true))
return cht
}
}
val untitled = chartname == "[Untitled]"
 // Try to get by title
        for (x in cv.indices)
{
cht = cv[x] as Chart
val cname = cht.title
if (cname.equals(chartname, ignoreCase = true))
return cht
else if (untitled && cname == "")
return cht
}
throw ChartNotFoundException(chartname)
}

/**
 * removes the desired chart from the list of charts
 *
 * @param chartname
 * @throws ChartNotFoundException
 */
    @Throws(ChartNotFoundException::class)
 fun removeChart(chartname:String) {
val cv = this.chartVect
var cht:Chart? = null
for (x in cv.indices)
{
cht = cv[x] as Chart
if (cht.title.equals(chartname, ignoreCase = true))
{
cv.removeAt(x)
return
}
}
throw ChartNotFoundException(chartname)
}

/**
 * NOT 100% IMPLEMENTE YET
 * creates the initial records for a Pivot Cache
 * <br></br>A Pivot Cache identifies the data used in a Pivot Table
 * <br></br>NOTE: only SHEET cache sources are supported at this time
 *
 * @param ref       String reference: either reference or named range
 * @param sheetName String sheetname
 * @param cacheid   if > 0, the desired cacheid (useful only in OOXML parsing)
 * @return int cacheid
 */
     fun addPivotStream(ref:String, sheetName:String, sid:Int):Int {
var sid = sid
 // in wb substream, DIRECTLY AFTER STYLE records:
        // STYLE/STYLEEX [TableStyle TableStyleElement] [Palette]
        // [ClrtClient]
        if (sid < 0)
sid = 0 // initial cache id if none already present
val records = this.streamer.records
var z = -1
var i = records.size - 1
while (i > 0 && z == -1)
{
val opcode = (records[i] as BiffRec).opcode.toInt()
if (opcode == XLSConstants.SXADDL.toInt())
{ // find last cache id and increment
 /*
                 * while (i > 0 && opcode!=SXSTREAMID)
                 * opcode= ((BiffRec) records.get(i--)).getOpcode();
                 * if (opcode==SXSTREAMID) {
                 * sid= ((SxStreamID) records.get(i+1)).getStreamID() + 1;
                 * }
                 */
                z = i + 1
}
else if (opcode == 4188)
 // ClrtClient
                z = i + 1
else if (opcode == XLSConstants.PALETTE.toInt())
z = i + 1
else if (opcode == 2192)
 // TableStyleElement
                z = i + 1
else if (opcode == 2194)
 // StyleEx
                z = i + 1
else if (opcode == XLSConstants.STYLE.toInt())
 // Style
                z = i + 1
i--
}

val tx = TableStyles.prototype as TableStyles? // see if
 // this is
        // really
        // necessary
        // ...
        this.streamer.addRecordAt(tx!!, z++)
val sxid = SxStreamID.prototype as SxStreamID?
this.streamer.addRecordAt(sxid!!, z++)
this.ptstream.add(sxid) // Pivot Cache -
sxid.setStreamID(sid) // cache id
this.streamer.records
.addAll(z, sxid.addInitialRecords(this, ref, sheetName))

return sid
}

/**
 * adds the Pivot Cache Directory Storage +Stream records necessary to
 * define the pivot cache (==pivot table data) for pivot table(s)
 * <br></br>NOTE: at this time only 1 pivot cache is supported
 *
 * @param ref Cell Range which identifies pivot table data range
 * @param wbh
 * @param sId Stream or cachid Id -- links back to SxStream set of records
 */
     fun addPivotCache(ref:String, wbh:WorkBookHandle, sId:Int) {
if (pivotCache == null)
{
pivotCache = PivotCache()
pivotCache!!.createPivotCache(factory!!.leoFile
.directoryArray, wbh, ref, sId)
}
}

/**
 * returns the start of the stream defining the desired pivot cache
 *
 * @param cacheid
 * @return
 */
     fun getPivotStream(cacheid:Int):SxStreamID? {
 // int z= 0;
        for (i in ptstream.indices)
{
val sid = (ptstream.get(i) as SxStreamID).streamID.toInt()
if (sid == cacheid)
 // if (z++==cacheid)
                return ptstream.get(i)
}
 /*
         * List records= this.getStreamer().records;
         * for (int i= 0; i < records.size(); i++) {
         * int opcode= ((BiffRec) records.get(i)).getOpcode();
         * if (opcode==SXSTREAMID) {
         * int sid= ((SxStreamID) records.get(i)).getStreamID() + 1;
         * if (sid==cacheid)
         * return (SxStreamID) records.get(i);
         * }
         * }
         */
        return null
}

 fun addIndirectFormula(f:Formula) {
indirectFormulas.add(f)
}

/**
 * Initialize the indirect functions in this workbook by calculating the formulas
 */
     fun initializeIndirectFormulas() {
val i = indirectFormulas.iterator() // contains all INDIRECT
 // funcvars + params
        while (i.hasNext())
{
val f = i.next() as Formula
f.calculateIndirectFunction()
}
indirectFormulas = ArrayList() // clear out
}

/**
 * Inserts an externsheet into the recvec, provided one does not yet exist.
 * also calls add supBook
 */
     fun addExternsheet() {
if (myexternsheet == null)
{
val numsheets = this.numWorkSheets
val sb = Supbook.getPrototype(numsheets) as Supbook
 // put it in after the last boundsheet record
            try
{
val b = this.getWorkSheetByNumber(numsheets - 1)
var loc = b.recordIndex + 1
if (streamer.getRecordAt(loc).opcode == XLSConstants.COUNTRY)
loc++
streamer.addRecordAt(sb, loc) // 20080306 KSC: do first 'cause
 // externsheet now references
                // global sb store
                this.addRecord(sb, false)
val ex = Externsheet
.getPrototype(0, 0, this) as Externsheet
streamer.addRecordAt(ex, loc + 1)// 20080306 KSC: must inc loc
 // since now inserting after
                // sb
                this.addRecord(ex, false)
myexternsheet = ex
}
catch (e:WorkSheetNotFoundException) {
Logger.logWarn("WorkBook.addExternSheet() locating Sheet for adding Externsheet failed: $e")
}

}
}

/**
 * add formula to book and init the ptgs
 *
 * @param rec
 */
     fun addFormula(rec:Formula) {
this.formulas.add(rec)
val shn = rec.sheet!!.sheetName + "!" + rec.cellAddress
this.formulashash.put(shn, rec)
}

/**
 * returns true if the default language selected in Excel is one of
 * the DBCS (Double-Byte Code Set) languages
 * <br></br>
 * The languages that support DBCS include
 * <br></br>
 * Japanese, Chinese (Simplified), Chinese (Traditional), and Korean
 * <pre>
 * Language        Country code    Countries/regions
 * -------------------------------------------------------------
 *
 * Arabic                966       (Saudi Arabia)
 * Czech                 42        (Czech Republic)
 * Danish                45        (Denmark)
 * Dutch                 31        (The Netherlands)
 * English               1         (The United States of America)
 * Farsi                 98        (Iran)
 * Finnish               358       (Finland)
 * French                33        (France)
 * German                49        (Germany)
 * Greek                 30        (Greece)
 * Hebrew                972       (Israel)
 * Hungarian             36        (Hungary)
 * Indian                91        (India)
 * Italian               39        (Italy)
 * Japanese              81        (Japan)
 * Korean                82        (Korea)
 * Norwegian             47        (Norway)
 * Polish                48        (Poland)
 * Portuguese (Brazil)   55        (Brazil)
 * Portuguese            351       (Portugal)
 * Russian               7         (Russian Federation)
 * Simplified Chinese    86        (People's Republic of China)
 * Spanish               34        (Spain)
 * Swedish               46        (Sweden)
 * Thai                  66        (Thailand)
 * Traditional Chinese   886       (Taiwan)
 * Turkish               90        (Turkey)
 * Urdu                  92        (Pakistan)
 * Vietnamese            84        (Vietnam)
</pre> *
 *
 * @return boolean
 */
     fun defaultLanguageIsDBCS():Boolean {
return (this.defaultLanguage == 81 || this.defaultLanguage == 886
|| this.defaultLanguage == 86 || this.defaultLanguage == 82)
 // PROBLEM WITH THIS: POSSIBLE TO BE SET AS DBCS DEFAULT
        // LANGUAGE
        // BUT HAVE NON-DBCS TEXT or VISA VERSA

        /*
         * In a double-byte character set, some characters require
         * two bytes,
         * while some require only one byte.
         * The language driver can distinguish between these two
         * types of characters by designating
         * some characters as "lead bytes."
         * A lead byte will be followed by another byte (a
         * "tail byte") to create a
         * Double-Byte Character (DBC).
         * The set of lead bytes is different for each language.
         *
         * Lead bytes are always guaranteed to be extended
         * characters; no 7-bit ASCII characters
         * can be lead bytes.
         * The tail byte may be any byte except a NULL byte.
         * The end of a string is always defined as the first NULL
         * byte in the string.
         * Lead bytes are legal tail bytes; the only way to tell if
         * a byte is acting as a
         * lead byte is from the context.
         */
    }

/**
 * clear out ALL sheet object refs
 */
     fun closeSheets() {
for (i in sheetVect.size - 1 downTo 1)
{
val b = sheetVect[i] as Boundsheet
if (b.streamer != null)
{ // do separately because may call
 // boundsheet close
                b.streamer!!.close()
b.streamer = null
}
b.close()
if (tabID != null)
tabID!!.removeRecord()
}
val recs = this.streamer.biffRecords
val resetVars = recs[0] != null
sheetVect.clear()
for (i in formulas.indices)
{
val f = formulas[i] as Formula
f.close()
}
this.formulas.clear()
if (this.refTracker != null)
this.refTracker!!.clearCaches()
this.formulashash.clear()
 // TODO: handle
        this.indirectFormulas.clear()
this.chartVect.clear()
this.chartTemp.clear()
if (firstBOF != null)
{
firstBOF!!.close()
this.firstBOF = null
}
if (lasteof != null)
{
lasteof!!.close()
lasteof = null
}
if (resetVars)
{ // just clearing out sheets instead of closing workbook
 // reset lasteof for possible new insertion of sheets (if
            // not removing workbook) ...
            var i = recs.size - 1
while (i > 0 && lasteof == null)
{
if (recs[i] != null)
{
if ((recs[i] as BiffRec).opcode == XLSConstants.EOF)
lasteof = recs[i] as Eof
}
i--
}
}
if (lastBOF != null)
{
lastBOF!!.close()
lastBOF = null
}
if (resetVars)
{ // just clearing out sheets instead of closing workbook
if ((recs[0] as BiffRec).opcode == XLSConstants.BOF)
{// should!!
lastBOF = recs[0] as Bof
firstBOF = recs[0] as Bof
}
}
if (lastbound != null)
{
lastbound!!.close()
this.lastbound = null
}
if (lastFormula != null)
{
lastFormula!!.close()
this.lastFormula = null
}
this.workSheets.clear()
}

protected fun closeRecords() {}

/**
 * clear out object references in prep for closing workbook
 */
     fun close() {
closeSheets()

if (this.isExcel2007)
{
try
{
val externalDir = OOXMLAdapter
.getTempDir(this.factory!!.fileName)
OOXMLAdapter.deleteDir(File(externalDir))
}
catch (e:Exception) {}

}
continueHandler.close()
continueHandler = ContinueHandler(this)

 // clear out array list references
        val ii = bookNameRecs.keys.iterator()
while (ii.hasNext())
{
val n = bookNameRecs.get(ii.next()) as Name
n.close()
}
this.bookNameRecs.clear()
for (i in xfrecs.indices)
{
val x = xfrecs[i] as Xf
x.close()
}
this.xfrecs.clear()
this.formatCache.clear()
this.formatlookup.clear()
this.formats.clear()
this.fontRecs.clear()
if (this.dxfs != null)
this.dxfs!!.clear()
for (i in msodgMerge.indices)
{
val m = msodgMerge[i] as MSODrawingGroup
m.close()
}
this.msodgMerge.clear()
if (this.msodg != null)
{
msodg!!.close()
msodg = null
}
 // TODO: handle
        this.mergecelllookup.clear()
this.hlinklookup.clear()
this.externalnames.clear()
for (i in names.indices)
{
val n = names.get(i) as Name
n.close()
}
this.names.clear()
for (i in indexes.indices)
{
val ind = indexes.get(i) as Index
ind.close()
}
this.indexes.clear()
if (lastINDEX != null)
{
lastINDEX!!.close()
lastINDEX = null
}
if (this.sharedStringTable != null)
{
this.sharedStringTable!!.close()
this.sharedStringTable = null
}

if (countryRec != null)
{
this.countryRec!!.close()
this.countryRec = null
}
if (win1 != null)
{
win1!!.close()
win1 = null
}

if (this.drh != null)
{
this.drh!!.close()
drh = null
}

 // integration point for subclasses
        this.closeRecords()

this.continueHandler = ContinueHandler(this)
if (this.currai != null)
{
this.currai!!.close()
this.currai = null
}
if (calcmoderec != null)
{
calcmoderec!!.close()
calcmoderec = null
}
this.currchart = null
this.currdrw = null
if (this.protector != null)
{
this.protector!!.close()
this.protector = null
}
if (this.myexternsheet != null)
{
this.myexternsheet!!.close()
this.myexternsheet = null
}
if (this.refTracker != null)
{
this.refTracker!!.close()
this.refTracker = null
}
if (tabID != null)
{
tabID!!.close()
tabID = null
}
 // TODO: deal
        this.factory = null
streamer = createByteStreamer()
if (xl2k != null)
(xl2k as XLSRecord).close()
xl2k = null

}

companion object {

private const val serialVersionUID = 2282017774412632087L

 // public constants
     var STRING_ENCODING_AUTO = 0
 var STRING_ENCODING_UNICODE = 1
 var STRING_ENCODING_COMPRESSED = 2
 var ALLOWDUPES = 0
 var SHAREDUPES = 1

 /*
     * This is here only to allow client code to compile, should
     * be removed
     */
     var CONVERTMULBLANKS = "deprecated"
}

}
