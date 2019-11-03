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

import io.starter.formats.XLS.formulas.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.UnsupportedEncodingException
import java.util.ArrayList
import java.util.Stack


/**
 * **Name: Defined Name (218h)**<br></br>
 *
 *
 * Name records describe a name in the workbook
 *
 *
 * <pre>
 * offset  name            size    contents
 * ---
 * 4       grbit           2       option flags
 * 6       chKey           1       Keyboard Shortcut
 * 7       cch             1       length of name text
 * 8       cce             2       length of name definition *stored in Excel parsed format
 * 10      ixals           2       index to sheet containing name
 * 12      itab            2       NAME SCOPE -- 0= workbook, 1+= sheet
 * 14      cchCustMenu     1       length of custom menu text
 * 15      cchDescript     1       length of description text
 * 16      cchHelpTopic    1       length of help topic text
 * 17      cchStatusText   1       length of status bar text
 * 18      rgch            var     name text
 * var     rgce            var     name definition
 * var     rcchCustMenu    var     cust menu text
 * var     rgchDescr       var     description text
 * var     rgchHelpTopic   var     help text
 * var     rgchStatusText  var     status bar text
 *
</pre> *
 */
class Name : XLSRecord {
    internal var grbit: Short = -1
    internal var DEBUG = false
    var isBuiltIn = false
        internal set
    internal var rgch = ""            //     name text
    internal var rgce = ""            //     name definition
    internal var rcchCustMenu = ""    //     cust menu text
    /**
     * get the descriptive text
     */
    internal var description = ""       //     description text
    internal var rgchHelpTopic = ""   //     help text
    internal var rgchStatusText = ""  //     status bar text
    internal var chKey: Byte = -1
    internal var cch: Byte = -1
    internal var cce: Short = -1   // 2
    internal var ixals: Short = -1   // 2
    internal var itab: Short = -1   // 2
    internal var cchCustMenu: Byte = -1
    internal var cchDescript: Byte = -1
    internal var cchHelpTopic: Byte = -1
    internal var cchStatusText: Byte = -1
    var builtInType: Byte = -1
    private var expression: Stack<*>? = null
    internal var externsheet: Externsheet? = null
    private var ptga: Ptg? = null
    val ilblListeners = ArrayList()
    /**
     * @return Returns the cachedOOXMLExpression.
     */
    /**
     * @param cachedOOXMLExpression The cachedOOXMLExpression to set.
     */
    var cachedOOXMLExpression: String? = null
    private var expressionbytes: ByteArray? = null        // for deferred Name expression init

    /* this byte array differs from the prototype in that it has no description field.  It is
        used by formulas inserting names that don't exist within the workbook.  The name will
        show up in the formula, but nowhere else on the excel file
    **/
    private val FILLER_NAME_BYTES = byteArrayOf(0x0, 0x0, 0x0, 0x1, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0)

    // 20090811 KSC: this prototype contains a range; clear out	private byte[] PROTOTYPE_NAME_BYTES = {0x0,0x0,0x0,0x6,0xb,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x61,0x64,0x66,0x61,0x64,0x66,0x3b,0x1,0x0,0x21,0x0,0x21,0x0,0x1,0x0,0x3,0x0};
    private val PROTOTYPE_NAME_BYTES = byteArrayOf(0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0)

    val boundSheets: Array<Boundsheet>?
        get() {
            if (ptga == null)
                try {
                    this.initPtga()
                } catch (e: Exception) {
                }

            if (ptga is PtgRef3d) {
                val p3d = ptga as PtgRef3d?
                val b = p3d!!.getSheet(this.workBook!!)
                val ret = arrayOfNulls<Boundsheet>(1)
                ret[0] = b
                return ret
            } else if (ptga is PtgArea3d) {
                val p3d = ptga as PtgArea3d?
                return p3d!!.getSheets(this.workBook!!)
            } else if (ptga is PtgMemFunc) {
                val p = ptga as PtgMemFunc?
                return p!!.getSheets(this.workBook)
            }
            return null
        }

    /**
     * Return the expression string for this name record.
     *
     * @return
     */
    val expressionString: String
        get() {
            if ((this.expression == null || this.expression!!.size == 0) && this.cachedOOXMLExpression != null)
                return "=" + this.cachedOOXMLExpression!!
            else if (this.expression != null)
                return FormulaParser.getExpressionString(this.expression!!)
            return "="
        }

    /**
     * Return the location of the Name record.  This seems to be slightly wrong as it only
     * returns the location of PTGA.  It's possible to have more complex records than this.
     *
     * @return
     * @throws Exception
     */
    // it's a NameX or some other non-Cell Name
    // 20080228 KSC: return Exception rather than null upon Deleted Named Range
    // JM - why not return loc? 'Cause it's deleted!!! :) // 20071203 KSC
    //20080214 KSC: returns correct string
    // +":"+ ptga.getLocation();; // BAD -- returns a 2d range for a 1d ref
    // PtgMemFunc ...
    var location: String?
        @Throws(Exception::class)
        get() {
            if (ptga == null) {
                try {
                    this.initPtga()
                } catch (e: Exception) {
                    Logger.logWarn("Name.getLocation() failed: $e")
                }

            }
            if (ptga == null)
                return null
            if (ptga is PtgRefErr3d)
                throw io.starter.formats.XLS.CellNotFoundException("Named Range " + this.name + " has been deleted or it's referenced cell is invalid")
            if (ptga is PtgArea3d)
                return ptga!!.location
            return if (ptga is PtgRef3d) ptga!!.location else ptga!!.toString()
        }
        set(newloc) = setLocation(newloc, true)

    /**
     * Is this a string referencing Name?
     *
     * @return
     */
    val isStringReference: Boolean
        get() = this.expression!!.size == 1 && this.expression!![0] is PtgStr

    override var workBook: WorkBook?
        get
        set(b) {
            super.workBook = b
        }

    /**
     * get the display name
     */
    /**
     * set the display name
     *
     *
     * Affects the following byte values:
     * 7       cch             1       length of name text
     * 18      rgch            var     name text
     */
    // 20100604 KSC: added handling of unicode
    // search for dreferenced names to rehook up
    var name: String
        get() = rgch
        set(newname) = try {
            var modnamelen = 0
            val dta = this.getData()
            val namebytes: ByteArray
            var isuni = false
            if (this.getByteAt(14).toInt() == 0x1) {
                namebytes = newname.toByteArray(charset(XLSConstants.UNICODEENCODING))
                isuni = true
            } else {
                namebytes = newname.toByteArray(charset(XLSConstants.DEFAULTENCODING))
            }
            modnamelen = namebytes.size
            var bodlen = dta!!.size
            bodlen -= cch.toInt()
            bodlen += modnamelen
            val newbytes = ByteArray(bodlen)
            System.arraycopy(dta, 0, newbytes, 0, 15)
            System.arraycopy(namebytes, 0, newbytes, 15, modnamelen)
            System.arraycopy(dta, cch + 15, newbytes, modnamelen + 15, dta.size - (cch + 15))
            cch = modnamelen.toByte()
            if (!isuni)
                newbytes[3] = cch
            else
                newbytes[3] = (cch / 2).toByte()
            this.setData(newbytes)
            rgch = newname
        } catch (e: UnsupportedEncodingException) {
            Logger.logWarn("UnsupportedEncodingException in setting NamedRange name: $e")
        }

    /**
     * return the case-insensitive version of the display name
     *
     * @return
     */
    // Case-insensitive
    val nameA: String
        get() = rgch.toUpperCase()

    internal var calc_id = 1

    /**
     * return the calculated value of this Name
     * if it contains a parsed Expression (Formula)
     *
     * @return
     * @throws FunctionNotSupportedException
     */
    val calculatedValue: Any
        @Throws(FunctionNotSupportedException::class)
        get() = FormulaCalculator.calculateFormula(this.expression!!)

    /**
     * get the definition text
     * the definition is stored
     * in Excel parsed format
     */
    internal/*
		StringBuffer sb = new StringBuffer();
		Ptg[] ep = new Ptg[expression.size()];
		ep = (Ptg[]) expression.toArray(ep);
		for(int t = 0;t<ep.length;t++){
			sb.append(ep[t].getString());
		}
		return sb.toString();*/ val definition: String
        get() {
            if (ptga == null) initPtga()
            return ptga!!.string
        }

    override//try{
    //	return this.getLocation();
    //}catch(Exception e){
    // ok
    //}
    val cellAddress: String
        get() = if (this.sheet != null) this.sheet.toString() + "!" + this.name else this.name

    /**
     * Returns an array of ptgs that represent any BiffRec ranges in the formula.
     * Ranges can either be in the format "C5" or "Sheet1!C4:D9"
     */
    // a single ref
    //		return ExpressionParser.getCellRangePtgs(expression);
    val cellRangePtgs: Array<Ptg>
        @Throws(FormulaNotFoundException::class)
        get() {
            if (ptga == null)
                initPtga()
            return ptga!!.components ?: return arrayOf<Ptg>(ptga)
        }

    constructor() {
        // default constructor
    }

    constructor(bk: WorkBook, namestr: String) {
        val bl = PROTOTYPE_NAME_BYTES
        this.setData(bl)
        this.opcode = XLSConstants.NAME
        this.length = bl.size.toShort()
        this.workBook = bk
        try {
            this.setExternsheet(bk.externSheet)
        } catch (x: WorkSheetNotFoundException) {
            Logger.logWarn("Name could not reference WorkBook Externsheet.$x")
        }

        this.init(false)
        this.name = namestr
        bk.insertName(this)
    }

    /**
     * Used for default name entry in formulas.  The name will not exist in the workbook outside
     * of the formula it is used in.  It has no location.
     *
     * @param bk Workbook containing name
     * @param b  Use whatever, just a flag to use this different constructor
     */
    constructor(bk: WorkBook, b: Boolean) {
        val bl = FILLER_NAME_BYTES
        this.setData(bl)
        this.opcode = XLSConstants.NAME
        this.length = bl.size.toShort()
        this.workBook = bk
        try {
            this.setExternsheet(bk.externSheet)
        } catch (x: WorkSheetNotFoundException) {
            Logger.logWarn("Name could not reference WorkBook Externsheet.$x")
        }

        this.init()
        bk.insertName(this)
    }

    /**
     * Store ptgName references to this Name record
     * so they can be accessed
     */
    fun addIlblListener(ptgname: IlblListener) {
        ilblListeners.add(ptgname)
    }

    fun removeIlblListener(ptgname: IlblListener) {
        ilblListeners.remove(ptgname)
    }

    fun updateIlblListeners() {
        val ilbl = this.workBook!!.getNameNumber(this.name).toShort()
        val i = ilblListeners.iterator()
        while (i.hasNext()) {
            (i.next() as IlblListener).ilbl = ilbl
        }
    }

    /**
     * Initialize the Name record
     *
     * @see io.starter.formats.XLS.XLSRecord.init
     */
    override fun init() {
        init(true)        // default= init Expression
    }

    /**
     * init Name record
     *
     * @param initExpression true if should parse formula/ref expression (will be false on wb load)
     */
    fun init(initExpression: Boolean) {
        super.init()
        this.getData()
        //  Logger.logInfo("[" + ByteTools.getByteString(data, false) + "]");

        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        chKey = this.getByteAt(2)
        cch = this.getByteAt(3)
        cce = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        ixals = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        itab = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())

        cchCustMenu = this.getByteAt(10)
        cchDescript = this.getByteAt(11)
        cchHelpTopic = this.getByteAt(12)
        cchStatusText = this.getByteAt(13)

        if (grbit and 0x20 == 0x20) isBuiltIn = true

        var pos = 15
        if (this.getByteAt(14).toInt() == 0x1) {
            cch *= 2
        }// rich byte;
        // get the Name
        try {
            val namebytes = this.getBytesAt(pos, cch.toInt())
            if (this.getByteAt(14).toInt() == 0x1) {
                rgch = String(namebytes!!, XLSConstants.UNICODEENCODING)
            } else if (isBuiltIn) {
                builtInType = namebytes!![0]
                when (builtInType) {
                    CONSOLIDATE_AREA -> rgch = "Built-in: CONSOLIDATE_AREA"

                    AUTO_OPEN -> rgch = "Built-in: AUTO_OPEN"

                    AUTO_CLOSE -> rgch = "Built-in: AUTO_CLOSE"

                    EXTRACT -> rgch = "Built-in: EXTRACT"

                    DATABASE -> rgch = "Built-in: DATABASE"

                    CRITERIA -> rgch = "Built-in: CRITERIA"

                    PRINT_AREA -> rgch = "Built-in: PRINT_AREA"

                    PRINT_TITLES -> rgch = "Built-in: PRINT_TITLES"

                    RECORDER -> rgch = "Built-in: RECORDER"

                    DATA_FORM -> rgch = "Built-in: DATA_FORM"

                    AUTO_ACTIVATE -> rgch = "Built-in: AUTO_ACTIVATE"

                    AUTO_DEACTIVATE -> rgch = "Built-in: AUTO_DEACTIVATE"

                    SHEET_TITLE -> rgch = "Built-in: SHEET_TITLE"

                    _FILTER_DATABASE -> rgch = "Built-in: _FILTER_DATABASE"
                }
            } else {
                rgch = String(namebytes!!)
            }
            if (DEBUG) Logger.logInfo(this.name)
            pos += cch.toInt()

            // get the parsed expression
            /*byte[] */
            expressionbytes = this.getBytesAt(pos, cce.toInt())
            if (initExpression)
                parseExpression()

        } catch (e: Exception) {
            if (DEBUGLEVEL > -1)
                Logger.logWarn("problem reading Name record expression for Name:" + this.name + " " + e)
        }

    }

    /**
     * parse Expression separately from init
     */
    fun parseExpression() {
        if (expressionbytes != null && expression == null) {
            expression = ExpressionParser.parseExpression(expressionbytes, this)
            if (DEBUGLEVEL == XLSConstants.DEBUG_LOW) Logger.logInfo(this.name + ":" + this.definition)
            if (expression == null) {
                val gpg = PtgMystery()
                gpg.init(expressionbytes!!)
                expression = Stack()
                expression!!.push(gpg)
            }
            expressionbytes = null
        }
    }

    /**
     * Return the name which identifies this Name record.  If this is a built in record
     * it will return a generic version of what the built-in is doing.
     *
     * @see io.starter.formats.XLS.XLSRecord.toString
     */
    override fun toString(): String {
        return this.name
    }

    /**
     * Get the expression for this Name record
     *
     * @return
     */
    fun getExpression(): Stack<*>? {
        return this.expression
    }

    /**
     * set the expression for this Name record
     */
    fun setExpression(x: Stack<*>) {
        expression = x
        this.updatePtgs()
    }

    /**
     * Essentially a wrapped setLocation call that is handled for initialization of names
     * from OOXML files.   Our parsing is not up to snuff for some more complex name records,
     * and this is a workaround until we are able to handle those expressions
     *
     *
     * TODO: parse complex expressions better and get rid of this method
     *
     * @param xpression
     */
    fun initializeExpression(xpression: String) {
        try {
            this.setLocation(xpression, false)
        } catch (e: FunctionNotSupportedException) {
            Logger.logWarn("Unable to parse Name record expression: $xpression")
            this.cachedOOXMLExpression = xpression
        }

    }

    /**
     * Set the location for the first ptg in the expression
     *
     * @param newloc
     * @param clearAffectedCells true if should clear ptga formula cached vals
     * @throws FunctionNotSupportedException TODO
     */
    @Throws(FunctionNotSupportedException::class)
    fun setLocation(newloc: String, clearAffectedCells: Boolean) {
        if (newloc.indexOf("{") > -1)
            throw FunctionNotSupportedException("Unable to parse string expression for name record $newloc")
        if (ptga == null)
            this.initPtga()
        if (ptga != null)
            (ptga as PtgRef).removeFromRefTracker()
        //	    ptga = new PtgArea3d(false);
        try {// can be a named value constant
            val f = Formula()
            f.workBook = this.wkbook
            val ptgs = FormulaParser.getPtgsFromFormulaString(f, newloc)
            if (ptgs.size == 1) {    // usual case of 1 ref or 1 PtgMemFunc (complex expression)
                ptga = ptgs.pop() as Ptg
                //  KSC: memory usage changes: now parent rec nec. for reference tracking; if change must update
                if (ptga is PtgRef && (ptga as PtgRef).useReferenceTracker)
                    (ptga as PtgRef).updateInRefTracker(this)
                ptga!!.parentRec = this
            } else {    // less common case of a formula expression
                expression = ptgs
                ptga = null    // otherwise will overwrite expression - see below for handling
            }
        } catch (e: Exception) {    // usually some #REF! error
            Logger.logErr("Name.setLocation: Error processing location $e")
        }

        if (ptga is PtgRef) {    // ensure that references are absolute
            (ptga as PtgRef).isColRel = false
            (ptga as PtgRef).isRowRel = false
            // TODO: get PtgMemFunc's components and ensure references are absolute
            // clear affected cells so that formulas which reference the named range get recalced with the new value(s)
            if (clearAffectedCells) {    // will only be false when initializing an OOXML workbook, which does not need to clear affected cells since it's initializing
                try {
                    val b = (ptga as PtgRef).refCells
                    for (i in b!!.indices)
                        if (b[i] != null)
                            this.workBook!!.refTracker!!.clearAffectedFormulaCells(b[i])
                } catch (e: NullPointerException) {
                }
                // if cells aren't present ...
            }
        }
        if (ptga != null) {
            expression = Stack()
            expression!!.add(ptga)    // update expression with new Ptg -- assume there's only 1 ptg!!
            try {
                this.externsheet!!.addPtgListener((ptga as IxtiListener?)!!)
            } catch (e: Exception) {
                // ptg constants will fail here, ptgNames ...
            }

        } else {
            initPtga()    // will calculate and set ptga to
        }
        this.updatePtgs()
    }

    // remove the name
    // why the boolean - NR
    override fun remove(b: Boolean): Boolean {
        val ret = super.remove(true)
        this.wkbook!!.removeName(this)
        return ret
    }

    /**
     * set the Externsheet rec
     */
    @Throws(WorkSheetNotFoundException::class)
    internal fun setExternsheet(e: Externsheet?) {
        this.externsheet = e
        if (e == null) this.externsheet = wkbook!!.getExternSheet(true)
    }

    /**
     * Initializes ptga.
     *
     *
     * Seems to be an init method to make this
     */
    internal fun initPtga() {
        if (expression == null)
            this.init()
        if (expression == null || expression!!.size == 0)
            return
        val p: Ptg
        //if the usual case of 1 reference-type (area, ref, memfunc...) ptg:
        if (this.expression!!.size == 1) {
            p = expression!![0] as Ptg
            if (p.isReference)
                ptga = p
        } else { // otherwise it's a formula expression
            p = FormulaCalculator.calculateFormulaPtg(this.expression!!)
            if (p.isReference)
                ptga = p
        }
        /*
		if (expression != null && expression.size() > 0){
			// this may be an invalid rec
			for(int t=0;t<expression.size();t++){
				Ptg p = (Ptg) expression.get(t);
				p.setParentRec(this);
				if (p.getIsReference())
					ptga= p;
				// otherwise it's a constant named range expression
			}
		}
*/
    }

    /**
     * set the Externsheet reference
     * for any associated PtgArea3d's
     */
    fun setExternsheetRef(x: Int) {
        // TODO: this doesn't account for formula expressions ...
        for (t in expression!!.indices) {
            val p = expression!![t] as Ptg
            if (p is PtgArea3d) {
                if (DEBUGLEVEL > 1) Logger.logInfo("PtgArea3d encountered in Ai record.")
                p.setIxti(x.toShort())
                ptga = p
            }
            if (p is PtgRef3d) {
                if (DEBUGLEVEL > 1) Logger.logInfo("PtgRef3d encountered in Ai record.")
                p.setIxti(x.toShort())
                ptga = p
            }
        }
    }

    /**
     * update Ptga ixti for moved/copied worksheets
     */
    fun updateSheetRefs(origWorkBookName: String) {
        if (ptga == null) {
            this.initPtga()
        }
        if (ptga is PtgArea3d) {    // PtgRef3d,etc
            val p = ptga as PtgArea3d?
            try {
                this.workBook!!.getWorkSheetByName(p!!.sheetName)
                ptga!!.location = ptga!!.toString()
            } catch (we: WorkSheetNotFoundException) {
                Logger.logWarn("External References Not Supported:  UpdateSheetReferences: External Worksheet Reference Found: " + p!!.sheetName!!)
                p.setExternalReference(origWorkBookName)
            }

        }
    }


    /**
     * do any pre-streaming processing such as expensive
     * index updates or other deferrable processing.
     */
    override fun preStream() {
        try {
            updatePtgs()
        } catch (e: Exception) {
            Logger.logWarn("problem updating Name record expression for Name:" + this.name)
        }

    }

    /***  Update the record byte array with the modified ptg records
     */
    fun updatePtgs() {
        if (expression == null)
        // happens upon init
            return
        val rkdata = this.getData()
        var offset = 15 + cch // the start of the parsed expression
        var sz = offset
        val sz2 = rkdata!!.size - (offset + cce)
        cce = 0
        // add up the size of the expressions
        for (i in expression!!.indices) {
            val ptg = expression!!.elementAt(i) as Ptg
            cce += ptg.length.toShort()
        }
        sz += cce.toInt()
        sz += sz2
        var updated = ByteArray(sz)
        System.arraycopy(rkdata, 0, updated, 0, offset)
        val cbytes = ByteTools.shortToLEBytes(cce)
        updated[4] = cbytes[0]
        updated[5] = cbytes[1]
        // 20090317 KSC: added handling for PtgArrays
        var hasArray = false
        var arraybytes = ByteArray(0)
        for (i in expression!!.indices) {
            val ptg = expression!!.elementAt(i) as Ptg
            val b: ByteArray
            // 20090317 KSC: added handling for PtgArrays
            if (ptg is PtgArray) {
                b = ptg.preRecord
                arraybytes = ByteTools.append(ptg.postRecord, arraybytes)
                hasArray = true
            } else {
                b = ptg.record
            }
            try {
                System.arraycopy(b, 0, updated, offset, b.size/*20071206 KSC: Not necessarily the same ptg.getLength()*/)
            } catch (e: Exception) {
                Logger.logInfo("setting ExternalSheetValue in Name rec: value: " + ptg.opcode + ": " + e)
            }

            offset = offset + ptg.length
        }
        // 20090317 KSC: added handling for PtgArrays
        if (hasArray) {
            updated = ByteTools.append(arraybytes, updated)
        }

        // add the rest if any
        if (sz2 > 0)
            System.arraycopy(rkdata, rkdata.size - sz2, updated, offset, sz2)
        this.setData(updated)
    }

    /**
     * Returns the ptg that matches the string location sent to it.
     * this can either be in the format "C5" or a range, such as "C4:D9"
     */
    fun getPtgsByLocation(loc: String): List<*>? {
        try {
            return ExpressionParser.getPtgsByLocation(loc, expression!!)
        } catch (e: FormulaNotFoundException) {
            Logger.logInfo("updating Chart Series Location failed: $e")
        }

        return null
    }

    /**
     * Set all ptg3ds to the new sheet
     * <br></br>Used when copying worksheets ..
     *
     * @param newSheet
     */
    fun updateSheetReferences(newSheet: Boundsheet) {
        for (i in expression!!.indices) {
            val o = expression!!.elementAt(i)
            if (o is PtgMemFunc) {
                val s = o.subExpression
                for (x in s!!.indices) {
                    val ox = s.elementAt(x)
                    if (ox is PtgArea3d) {
                        ox.setReferencedSheet(newSheet)
                    }// do we have other types we need to handle here? (within memfunc)
                }
            } else if (o is PtgArea3d) {// do we have other types we need to handle here? (outside memfunc);
                o.setReferencedSheet(newSheet)
            }
            // do nothing
        }
        this.updatePtgs()
    }


/**
 * Return an array of ptgs that make up this Name record
 *
 * @return
 */
/*	public Ptg[] getComponents(){
		if (ptga==null)
			initPtga();
		return ptga.getComponents();
/*
//		if (ptga!=null) { // then this is a location-type ptg (ref or memfunc ...)
			Ptg[] p = new Ptg[expression.size()];
			for (int i=0;i<expression.size();i++){
				p[i] = (Ptg)expression.elementAt(i);
			}
			return p;
//		} else {	// calculate formula - assume ends up as a reference
//			Ptg p= FormulaCalculator.calculateFormulaPtg(this.expression);
//			return new Ptg[] {p};
//		}
 * */
 //	}

    /**
 * locks the Ptg at the specified location
 */

     fun setLocationPolicy(loc:String, l:Int):Boolean {
val dx = this.getPtgsByLocation(loc)
val lx = dx!!.iterator()
while (lx.hasNext())
{
val d = lx.next() as Ptg
if (d != null)
{
d!!.locationPolicy = l
}
}
return true
}

 fun getExternsheet():Externsheet? {
return externsheet
}

 fun getPtga():Ptg? {
if (ptga == null)
initPtga()
return ptga
}

/**
 * Set this name record as a built in type.  Use the static bytes
 * to set the built in type.
 *
 * @param builtinType
 */
     fun setBuiltIn(builtinType:Byte) {
 // 20100215 KSC: redo
        grbit = grbit or 0x20    // set built-in bt
if (builtinType == _FILTER_DATABASE)
 // TODO: is this the only one?
            grbit = grbit or 0x1    // set hidden bit
val grbytes = ByteTools.shortToLEBytes(grbit)
val newData = ByteArray(16)
newData[0] = grbytes[0]
newData[1] = grbytes[1]
newData[3] = 0x1            // cch
newData[15] = builtinType
this.setData(newData)
this.init()
}

 fun getIxals():Short {
return ixals
}

 fun setIxals(ixals:Short) {
this.ixals = ixals
val b = ByteTools.shortToLEBytes(ixals)
val rkdata = this.getData()
rkdata[6] = b[0]
rkdata[7] = b[1]
this.setData(rkdata)
}

/**
 * Return the named range scope (0= workbook, 1 or more= sheet)
 *
 * @return Named Range Scope
 */
     fun getItab():Short {
return itab
}

 fun setItab(itab:Short) {
this.itab = itab
val b = ByteTools.shortToLEBytes(itab)
this.getData()[8] = b[0]
this.getData()[9] = b[1]
}

/**
 * sets a namestring to a constant (non-reference) value
 *
 * @param bk
 * @param namestr String name
 * @param value   The expression statement
 * @param scope   1 based reference to sheet scope, 0 is for workbook (default)
 */
     constructor(bk:WorkBook, namestr:String, value:String, scope:Int) {
val bl = PROTOTYPE_NAME_BYTES
this.setData(bl)
this.opcode = XLSConstants.NAME
this.length = (bl.size).toShort()
this.setItab(scope.toShort())
this.workBook = bk
try
{
this.setExternsheet(bk.externSheet)
}
catch (x:WorkSheetNotFoundException) {
Logger.logWarn("Name could not reference WorkBook Externsheet." + x.toString())
}

this.init()
this.name = namestr
bk.insertName(this)    // calls addRecord which calls addName
this.initializeExpression(value)
}

/**
 * Set the scope (itab) of this name
 *
 * @param newitab
 * @throws WorkSheetNotFoundException
 */
    @Throws(WorkSheetNotFoundException::class)
 fun setNewScope(newitab:Int) {
if (this.itab.toInt() == 0)
{
this.workBook!!.removeLocalName(this)
}
else
{
this.workBook!!.getWorkSheetByNumber(this.itab - 1).removeLocalName(this)
}
if (newitab == 0)
{
this.workBook!!.addLocalName(this)
}
else
{
this.workBook!!.getWorkSheetByNumber(newitab - 1).addLocalName(this)
}
this.setItab(newitab.toShort())
}

/**
 * clear out object references in prep for closing workbook
 */
    public override fun close() {
this.externsheet = null
if (ptga != null)
{
if (ptga is PtgRef)
ptga!!.close()
else
ptga!!.close()
ptga = null
}
if (expression != null)
{
while (!expression!!.isEmpty())
{
var p:GenericPtg? = expression!!.pop() as GenericPtg
if (p is PtgRef)
p!!.close()
else
p!!.close()
p = null
}
}
super.close()
}

protected fun finalize() {
this.close()
}

companion object {

/**
 *
 */
    private val serialVersionUID = -7868028144327389601L

protected val CONSOLIDATE_AREA:Byte = 0x0
protected val AUTO_OPEN:Byte = 0x1
protected val AUTO_CLOSE:Byte = 0x2
protected val EXTRACT:Byte = 0x3
protected val DATABASE:Byte = 0x4
protected val CRITERIA:Byte = 0x5
 val PRINT_AREA:Byte = 0x6
 val PRINT_TITLES:Byte = 0x7
protected val RECORDER:Byte = 0x8
protected val DATA_FORM:Byte = 0x9
protected val AUTO_ACTIVATE:Byte = 0xA
protected val AUTO_DEACTIVATE:Byte = 0xB
protected val SHEET_TITLE:Byte = 0xC
 val _FILTER_DATABASE:Byte = 0xD
}
}