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

import io.starter.OpenXLS.*
import io.starter.formats.OOXML.*
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserException
import org.xmlpull.v1.XmlPullParserFactory

import java.io.*
import java.util.*
import java.util.zip.ZipEntry
import java.util.zip.ZipFile

/**
 * Breaking out functionality for reading out of OOXMLAdapter
 */
class OOXMLReader : OOXMLAdapter(), OOXMLConstants {


    // records used in parsing styles.xml
    internal var borders: ArrayList<*>
    internal var fontmap: ArrayList<Int>
    internal var fills: ArrayList<Fill>
    internal var dxfs: ArrayList<Dxf>
    internal var fmts: HashMap<*, *>
    internal var nXfs: Int = 0
    //	private int defaultXf= 0;	// usual for OOXML files; however, those which are converted from XLS may have default xf as 15

    /** */
    /**
     * Parsing/Reading OOXML Input Section /
     */

    /**
     * OOXML parseNBind - reads in an OOXML (Excel 7) workbook
     *
     * @param bk    WorkBookHandle - workbook to input
     * @param fName OOXML filename (must be a ZIP file in OPC format)
     * @throws XmlPullParserException
     * @throws IOException
     * @throws CellNotFoundException
     */
    @Throws(XmlPullParserException::class, IOException::class, CellNotFoundException::class)
    fun parseNBind(bk: WorkBookHandle, fName: String) {
        val zip = ZipFile(fName)

        OOXMLAdapter.inputEncoding = System.getProperty("file.encoding")
        if (OOXMLAdapter.inputEncoding == null) OOXMLAdapter.inputEncoding = "UTF-8"
        // KSC: replaced with above       isUTFEncoding= (System.getProperty("file.encoding").startsWith("UTF"));
        // clear out state vars
        bk.workBook!!.isExcel2007 = true
        val origcalcmode = bk.workBook!!.calcMode
        bk.workBook!!.calcMode = WorkBook.CALCULATE_EXPLICIT   // don't calculate formulas on input
        // TODO: read in format type (doc, macro-enabled doc, template, macro-enabled tempate) from CONTENT_LIST ???
        val rels = OOXMLAdapter.getEntry(zip, "_rels/.rels")
        if (DEBUG)
            Logger.logInfo("parseNBind about to call parseRels on: $rels")
        mainContentList = OOXMLAdapter.parseRels(OOXMLAdapter.wrapInputStream(zip.getInputStream(rels)))
        bk.workBook!!.factory!!.fileName = fName
        bk.setDupeStringMode(WorkBookHandle.SHAREDUPES)

        /* KSC: remove Xf recs first  -- NOTE has some issues for XLS->XLSX -- must fix !! */
        bk.workBook!!.removeXfRecs()
        bk.workBook!!.defaultIxfe = 0

        externalDir = OOXMLAdapter.getTempDir(bk.workBook!!.factory!!.fileName)
        val formulas = ArrayList()     // set in parseSheetXML, must process formulas after all sheets/cells have been added
        val hyperlinks = ArrayList()   // set in parseSheetXML, links with hyperlink target info in sheetX.xml.rels            ""
        val inlineStrs = HashMap()               // set in parseSheetXML, stores inlinestring text with addresses for entry after all sheets have been added
        parseBookLevelElements(bk, null, zip, mainContentList, "", formulas, hyperlinks, inlineStrs, null, null)
        zip.close()
        // if hasn't been streamed, delete temp dir
        if (!bk.workBook!!.factory!!.fileName.endsWith(".tmp"))
            OOXMLAdapter.deleteDir(File(externalDir))    // don't save temp files (pass-through's) -- can reinstate when needed
        bk.workBook!!.calcMode = origcalcmode  // reset
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
    protected fun parseBookLevelElements(bk: WorkBookHandle, sheet: WorkSheetHandle?, zip: ZipFile, cl: ArrayList<*>, parentDir: String, formulas: ArrayList<*>, hyperlinks: ArrayList<*>, inlineStrs: HashMap<*, *>?, pivotCaches: HashMap<String, String>?, pivotTables: HashMap<String, WorkSheetHandle>?) {
        var sheet = sheet
        var pivotCaches = pivotCaches
        var pivotTables = pivotTables
        var p: String    // target path
        var target: ZipEntry?
        var sst = ArrayList() // set in parseSSTXML, used in parsing sheet XML

        try {
            // parse content list for <elementName, target's path, rId>
            for (i in cl.indices) {
                val c = cl[i] as Array<String>

                if (DEBUG)
                    Logger.logInfo("OOXMLReader.parse: " + c[0] + ":" + c[1] + ":" + c[2])

                p = StringTool.getPath(c[1])
                p = parsePathForZip(p, parentDir)

                val ooxmlElement = c[0]
                if (ooxmlElement != "hyperlink")
                // if it's a hyperlink reference, don't strip path info :)
                    c[1] = StringTool.stripPath(c[1])
                val f = c[1]    // root filename
                val rId = c[2]

                if (ooxmlElement == "styles") {
                    target = OOXMLAdapter.getEntry(zip, p + f)
                    parseStylesXML(bk, OOXMLAdapter.wrapInputStream(zip.getInputStream(target!!)))
                } else if (ooxmlElement == "sst") {
                    target = OOXMLAdapter.getEntry(zip, p + f)
                    sst = Sst.parseOOXML(bk, OOXMLAdapter.wrapInputStream(zip.getInputStream(target!!)))
                } else if (ooxmlElement == "sheet") {
                    // sheet.xml
                    target = OOXMLAdapter.getEntry(zip, p + f)

                    try {
                        var sheetnum = 1
                        try {
                            val s = rId.substring(3)        // in form of "rIdXX" where XX is the sheet number
                            sheetnum = Integer.valueOf(s).toInt() - 1  // embed attribute, specifies rId, important in OOXML
                        } catch (e: Exception) {
                            Logger.logWarn("OOXMLAdapter couldn't get sheet number from rid:$rId")
                        }

                        sheet = bk.getWorkSheet(sheetnum)
                        // since we're adding a lot of cells, put sheet in fast add mode    // put statement here AFTER sheet is set :)
                        sheet!!.fastCellAdds = true

                        sheet.mysheet!!.parseOOXML(bk, sheet, OOXMLAdapter.wrapInputStream(zip.getInputStream(target!!)), sst, formulas, hyperlinks, inlineStrs)

                        // sheet.xml.rels
                        target = OOXMLAdapter.getEntry(zip, p + "_rels/" + f.substring(f.lastIndexOf("/") + 1) + ".rels")
                        if (target != null) {
                            try {
                                val pts = HashMap<String, WorkSheetHandle>()
                                sheet.mysheet!!.parseSheetElements(bk, zip, OOXMLAdapter.parseRels(OOXMLAdapter.wrapInputStream(OOXMLAdapter.wrapInputStream(zip.getInputStream(target)))), p, externalDir, formulas, hyperlinks, inlineStrs, pts)
                                if (pts.size > 0) {
                                    pivotTables!!.putAll(pts)
                                }
                            } catch (e: Exception) {
                                Logger.logWarn("OOXMLAdapter.parse problem parsing rels in: $bk $e")
                            }

                        }
                        // reset fast add mode
                        sheet.fastCellAdds = false   // 20090713 KSC: moved from below
                    } catch (we: WorkSheetNotFoundException) {
                        Logger.logErr("OOXMLAdapter.parse: $we")
                    }

                } else if (ooxmlElement == "document") { // main workbook document
                    // workbook.xml
                    target = OOXMLAdapter.getEntry(zip, p + f)

                    if (DEBUG)
                        Logger.logInfo("About to parseWBOOXML:$bk")
                    pivotCaches = HashMap()
                    parsewbOOXML(zip, bk, OOXMLAdapter.wrapInputStream(zip.getInputStream(target!!)), p, pivotCaches)    // seets, defined names, pivotcachedefinition ...

                    // now parse wb content - sheets and their sub-contents (charts, images, oleobjects...)
                    pivotTables = HashMap()
                    parseBookLevelElements(bk, sheet, zip, wbContentList, p, formulas, hyperlinks, inlineStrs, pivotCaches, pivotTables)

                    // after all sheet data has been added, now can add inline strings, if any
                    if (inlineStrs != null)
                        addInlineStrings(bk, inlineStrs)
                    // after all sheet data has been added, now can add formulas
                    addFormulas(bk, formulas)
                    // after all sheet data and formulas, NOW can add pivot Tables
                    addPivotTables(bk, zip, pivotTables)
                } else if (parsePivotTables && ooxmlElement == "pivotCacheDefinition") {    // workbook-parent + pivotTable-parent
                    //pivotCaches.add(new int[] {cid, id});
                    target = OOXMLAdapter.getEntry(zip, p + f)
                    PivotCacheDefinition.parseOOXML(bk, pivotCaches!![rId], OOXMLAdapter.wrapInputStream(zip.getInputStream(target!!)))
                    target = OOXMLAdapter.getEntry(zip, p + "_rels/" + f.substring(f.lastIndexOf("/") + 1) + ".rels")
                    if (target != null) {    // pivotCacheRecords ...
                        try {
                            parseBookLevelElements(bk, sheet, zip, OOXMLAdapter.parseRels(OOXMLAdapter.wrapInputStream(OOXMLAdapter.wrapInputStream(zip.getInputStream(target)))), p, formulas, hyperlinks, inlineStrs, pivotCaches, pivotTables)
                        } catch (e: Exception) {
                            Logger.logWarn("OOXMLAdapter.parse problem parsing rels in: $bk $e")
                        }

                    }
                } else if (parsePivotTables && ooxmlElement == "pivotCacheRecords") {    // pivotcacheDefinition-parent
                } else if (ooxmlElement == "theme" || ooxmlElement == "themeOverride") { // read in theme colors
                    target = OOXMLAdapter.getEntry(zip, p + f)
                    if (target != null) {
                        if (bk.workBook!!.theme == null)
                            bk.workBook!!.theme = Theme.parseThemeOOXML(bk, OOXMLAdapter.wrapInputStream(zip.getInputStream(target)))
                        else
                            bk.workBook!!.theme!!.parseOOXML(bk, OOXMLAdapter.wrapInputStream(zip.getInputStream(target)))    // theme overrides
                    }
                    handlePassThroughs(zip, bk, p, externalDir, c)
                    // Below are elements we do not as of yet handle
                } else if (ooxmlElement == "props"
                        || ooxmlElement == "exprops"
                        || ooxmlElement == "custprops"
                        || ooxmlElement == "connections"
                        || ooxmlElement == "calc"
                        || ooxmlElement == "vba"
                        || ooxmlElement == "externalLink") {

                    handlePassThroughs(zip, bk, p, externalDir, c)   // pass-through this file and any embedded objects as well
                } else {    // unknown type
                    Logger.logWarn("OOXMLReader.parse:  XLSX Option Not yet Implemented $ooxmlElement")
                }
            }
        } catch (e: IOException) {
            Logger.logErr("OOXMLReader.parse failed: $e")
        }

    }

    /**
     * given workbook.xml inputstream, parse OOXML into array list of content (only sheets and names at this point; eventually will handle docProps ...)
     *
     * @param bk         WorkBookHandle
     * @param ii         inputStream
     * @param namedRange ArrayList to hold named ranges (must be added after all sheet data)
     * @return
     */
    internal fun parsewbOOXML(zip: ZipFile, bk: WorkBookHandle, ii: InputStream, p: String, pivotCaches: HashMap<String, String>): ArrayList<*> {
        val namedRanges = ArrayList()  //must save and parse after all sheets have been added
        val contentList = ArrayList()
        val sheets = ArrayList()

        // set the default date format
        bk.workBook!!.dateFormat = DateConverter.DateFormat.OOXML_1900

        try {
            val factory = XmlPullParserFactory.newInstance()
            factory.isNamespaceAware = true
            val xpp = factory.newPullParser()

            xpp.setInput(ii, null) // using XML 1.0 specification
            var eventType = xpp.eventType
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    val tnm = xpp.name
                    if (tnm == "sheet") {
                        var name = "" // name, sheetId, rId, state=hidden
                        var id = 0      // sheetId is used??
                        var rId = 0
                        var hidden = ""
                        for (i in 0 until xpp.attributeCount) {
                            val nm = xpp.getAttributeName(i)
                            val v = xpp.getAttributeValue(i)
                            if (nm.equals("name", ignoreCase = true))
                                name = v
                            else if (nm.equals("SheetId", ignoreCase = true))
                                id = Integer.valueOf(v).toInt() - 1
                            else if (nm.equals("id", ignoreCase = true))
                            // rId
                                rId = Integer.valueOf(v.substring(3)).toInt() - 1
                            else if (nm == "state") {
                                hidden = v
                            }
                        }
                        // sheets may very well NOT be in order so must create after all sheets have been accounted for bk.createWorkSheet(name, id);
                        for (i in sheets.size until rId) {
                            sheets.add("")
                        }
                        sheets.add(rId, arrayOf(name, hidden))
                        contentList.add(arrayOf("sheet", name))
                    } else if (tnm == "workbookPr") { // TODO: get other attributes such as date1904
                        for (i in 0 until xpp.attributeCount) {
                            val attrName = xpp.getAttributeName(i)
                            val attrValue = xpp.getAttributeValue(i)
                            if (attrName.equals("codeName", ignoreCase = true))
                                bk.workBook!!.codename = attrValue
                            else if (attrName.equals("dateCompatibility", ignoreCase = true) && attrValue == "1")
                                bk.workBook!!.dateFormat = DateConverter.DateFormat.LEGACY_1900
                            else if (attrName.equals("date1904", ignoreCase = true) && attrValue == "1")
                                bk.workBook!!.dateFormat = DateConverter.DateFormat.LEGACY_1904
                        }
                    } else if (tnm == "workbookView") {   // TODO: handle other workbookview attributes
                        var n = ""
                        for (i in 0 until xpp.attributeCount) {
                            n = xpp.getAttributeName(i)
                            if (n.equals("firstSheet", ignoreCase = true))
                                bk.workBook!!.firstSheet = Integer.valueOf(xpp.getAttributeValue(i)).toInt()
                            else if (n == "showSheetTabs") {
                                val b = xpp.getAttributeValue(i) != "0"
                                bk.workBook!!.setShowSheetTabs(b)
                            }//else if (n.equalsIgnoreCase("activeTab"))
                            //bk.getWorkBook().setActiveTab(Integer.valueOf(xpp.getAttributeValue(i)).intValue());
                        }
                        /*                   } else if (tnm.equals("externalReference")) {  // TODO: HANDLE! nothing really to do here as it just denotes an rId, handled upon encountering externalLink */
                    } else if (tnm == "definedName") {    // have to process after sheets have been added, so save info
                        // attributes:  (not used by us) comment, customMenu, description, help, shortcutKey, statusBar, vbProcedure, workbookParameter, publishToServer
                        // (to deal with later) function, functionGroupId -- Specifies a boolean value that indicates that the defined name refers to a user-defined function/add-in ...
                        // xlm (External Function)
                        // (should deal with) hidden, localSheetId
                        var nm = ""
                        var id = ""
                        for (i in 0 until xpp.attributeCount) {
                            val n = xpp.getAttributeName(i)
                            if (n == "name")
                            // can be built-in:  _xlnm.Print_Area, _xlnm.Print_Titles, _xlnm.Criteria, _xlnm ._FilterDatabase, _xlnm .Extract, _xlnm .Consolidate_Area, _xlnm .Database, _xlnm .Sheet_Title
                                nm = xpp.getAttributeValue(i)
                            else if (n == "localSheetId")
                            // Specifies the sheet index in this workbook where data from an external reference is displayed.
                                id = xpp.getAttributeValue(i)
                        }
                        var name = OOXMLAdapter.getNextText(xpp)    // value can be a function, a
                        // has an external wb specification, remove as messes up parsing
                        if (id != "" && name!!.startsWith("[")) { // remove external denotation [#]sheet!range
                            var n = 0
                            while ((n = name!!.indexOf("[")) > -1) {
                                name = name!!.substring(0, n) + name.substring(n + name.substring(n).indexOf("]") + 1)
                            }
                        }
                        namedRanges.add(arrayOf(nm, id, name))
                    } else if (tnm == "pivotCache") {
                        var cid = ""
                        var rid = ""
                        for (i in 0 until xpp.attributeCount) {
                            val n = xpp.getAttributeName(i)
                            if (n == "cacheId") {
                                cid = xpp.getAttributeValue(i)
                            } else if (n == "id") {
                                rid = xpp.getAttributeValue(i)
                            }
                        }
                        pivotCaches[rid] = cid
                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                }
                eventType = xpp.next()
            }
        } catch (e: Exception) {
            Logger.logErr("OOXMLAdapter.parsewbOOXML failed: $e")
        }

        for (i in sheets.indices) {
            val obx = sheets.get(i)

            var sh: Array<String>? = null
            if (obx is String) {
                sh = arrayOfNulls<String>(1)
                sh[0] = obx.toString()
            } else {
                sh = obx as Array<String>
            }
            val name = sh!![0]
            if (name != null && name != "") {
                bk.createWorkSheet(name, i)
                try {
                    bk.workBook!!.defaultIxfe = 0    // OOXML default xf= 0, 2003-vers, default= 15 see below
                    if (sh!![1] == "hidden")
                        bk.getWorkSheet(i).hidden = true
                    else if (sh!![1] == "veryHidden")
                        bk.getWorkSheet(i).veryHidden = true
                } catch (e: Exception) {
                    // shouldn't!
                }

            }
        }
        // workbook.xml.rels
        val target = OOXMLAdapter.getEntry(zip, p + "_rels/workbook.xml.rels")
        try {
            wbContentList = OOXMLAdapter.parseRels(OOXMLAdapter.wrapInputStream(OOXMLAdapter.wrapInputStream(zip.getInputStream(target))))
        } catch (e: IOException) {
            Logger.logWarn("OOXMLReader.parseWbOOXML: $e")
        }

        // for workbook contents, MUST PROCESS themes before styles, sst and styles, etc. before SHEETS
        reorderWbContentList(wbContentList)

        // add all named ranges
        addNames(bk, namedRanges)

        return contentList
    }

    /**
     * given Styles.xml OOXML input stream, parse and input into workbook
     *
     * @param bk WorkBookHandle
     * @param ii InputStream
     */
    internal fun parseStylesXML(bk: WorkBookHandle, ii: InputStream) {
        try {
            borders = ArrayList()
            fontmap = ArrayList()
            fills = ArrayList()
            dxfs = ArrayList()
            fmts = HashMap()
            nXfs = 0                                  // position in xfrecs array is vital as cells will reference the styleId/xfId
            var indexedColor = 0                          // index into COLOR_TABLE

            val factory = XmlPullParserFactory.newInstance()
            factory.isNamespaceAware = true
            val xpp = factory.newPullParser()

            xpp.setInput(ii, null) // using XML 1.0 specification
            var eventType = xpp.eventType
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    val tnm = xpp.name
                    if (tnm == "font") {
                        val f = Font.parseOOXML(xpp, bk)
                        val idx = FormatHandle.addFont(f, bk)
                        fontmap.add(Integer.valueOf(idx))
                    } else if (tnm == "dxfs") {   // differential formatting (conditional formatting) style
                    } else if (tnm == "dxf") { // incremental style info -- for conditional save
                        val d = Dxf.parseOOXML(xpp, bk).cloneElement() as Dxf
                        dxfs.add(d)
                    } else if (tnm == "fill") {
                        val f = Fill.parseOOXML(xpp, false, bk) as Fill
                        fills.add(f)    //new Object[] { Integer.valueOf(fp), fgColor, bgColor});
                    } else if (tnm == "numFmt") {
                        var fmtId = 0
                        var newFmtId = 0
                        var xmlFormatPattern = ""
                        for (i in 0 until xpp.attributeCount) {
                            val nm = xpp.getAttributeName(i)
                            if (nm == "numFmtId") {
                                fmtId = Integer.valueOf(xpp.getAttributeValue(i)).toInt()
                            } else if (nm == "formatCode") {
                                xmlFormatPattern = xpp.getAttributeValue(i)
                                xmlFormatPattern = Xf.unescapeFormatPattern(xmlFormatPattern)
                            }
                        }
                        newFmtId = Xf.addFormatPattern(bk.workBook!!, xmlFormatPattern).toInt()
                        fmts[Integer.valueOf(fmtId)] = Integer.valueOf(newFmtId)  // map our format id with original
                    } else if (tnm == "border") { // TODO: use Border element to parse
                        val b = Border.parseOOXML(xpp, bk).cloneElement() as Border
                        borders.add(b)
                    } else if (tnm == "cellXfs") {
                        while (eventType != XmlPullParser.END_DOCUMENT) {
                            if (eventType == XmlPullParser.START_TAG) {
                                this.parseCellXf(xpp, bk)
                            } else if (eventType == XmlPullParser.END_TAG && xpp.name == "cellXfs")
                                break
                            eventType = xpp.next()
                        }
                    } else if (tnm == "rgbColor") {
                        // save custom indexed colors
                        val clr = "#" + xpp.getAttributeValue(0).substring(2)
                        //io.starter.toolkit.Logger.log(clr);
                        // usually the same as COLORTABLE but sometimes different too :)
                        try {
                            bk.workBook!!.colorTable[indexedColor++] = FormatHandle.HexStringToColor(clr)
                        } catch (e: ArrayIndexOutOfBoundsException) {
                            // happens?
                        }

                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                    val endTag = xpp.name
                    if (endTag == "worksheet")
                    // we're done!
                        break
                } else if (eventType == XmlPullParser.TEXT) {
                }
                if (eventType != XmlPullParser.END_DOCUMENT)
                    eventType = xpp.next()
            }
            if (dxfs.size > 0)
                bk.workBook!!.dxfs = dxfs

        } catch (e: Exception) {
            Logger.logErr("OOXMLReader.parseStylesXML: $e")
        }

    }


    /**
     * Parse the cellXF's section of the styles.xml file
     *
     * @param xpp
     * @param bk
     */
    private fun parseCellXf(xpp: XmlPullParser, bk: WorkBookHandle) {
        val tnm = xpp.name
        if (tnm == "xf") {
            var f = 0
            var fmtId = -1
            var fillId = -1
            var borderId = -1
            for (i in 0 until xpp.attributeCount) {
                val nm = xpp.getAttributeName(i)
                // all id's are 0-based
                if (nm == "fontId") {
                    f = Integer.valueOf(xpp.getAttributeValue(i)).toInt()
                } else if (nm == "numFmtId") {
                    fmtId = Integer.valueOf(xpp.getAttributeValue(i)).toInt()
                } else if (nm == "fillId") {
                    fillId = Integer.valueOf(xpp.getAttributeValue(i)).toInt()
                } else if (nm == "borderId") {
                    borderId = Integer.valueOf(xpp.getAttributeValue(i)).toInt()
                }
            }
            f = fontmap[f]  // FONT
            var xf: Xf? = null
            if (nXfs < bk.workBook!!.xfrecs.size)
            // either alter existing default xf or create new xf
                xf = bk.workBook!!.xfrecs[nXfs] as Xf
            if (xf == null)
            // if it doesn't exist, create new otherwise overwrite orig
                xf = Xf.updateXf(null, f, bk.workBook)
            else {
                xf.setFont(f)
                xf.setFormat(0.toShort())
            }
            if (fmtId > 0) { // NUMBER FORMAT 0 is default
                if (fmts.get(Integer.valueOf(fmtId)) != null)
                // map it
                    fmtId = (fmts.get(Integer.valueOf(fmtId)) as Int).toInt()
                xf!!.setFormat(fmtId.toShort())
            }
            if (borderId > -1) {     // BORDER
                val b = borders[borderId] as Border
                xf!!.setAllBorderLineStyles(b.borderStyles)    //bs);
                xf.setAllBorderColors(b.borderColorInts)
            }
            if (fillId > 0) {    // FILL 0 is default
                xf!!.setFill(fills[fillId])
            }
            // is xf 15 the default? (will happen if converted from xls) ******* very important to avoid unnecessary blank creation *******
            // see TestCorruption.TestStackOverflow
            if (nXfs == 15 && xf!!.toString() == bk.workBook!!.getXf(0)!!.toString())
                bk.workBook!!.defaultIxfe = 15
            nXfs++

        } else if (tnm == "protection") {
            val xf = bk.workBook!!.getXf(nXfs - 1)
            for (j in 0 until xpp.attributeCount) {
                val n = xpp.getAttributeName(j)
                val v = xpp.getAttributeValue(j)
                if (n == "hidden")
                    xf!!.isFormulaHidden = v == "1"
                else if (n == "locked")
                    xf!!.isLocked = v == "1"
            }
        } else if (tnm == "alignment") {
            val xf = bk.workBook!!.getXf(nXfs - 1)
            for (j in 0 until xpp.attributeCount) {
                val n = xpp.getAttributeName(j)
                val v = xpp.getAttributeValue(j)
                if (n == "horizontal") {
                    val ha = sLookup(v, OOXMLConstants.horizontalAlignment)
                    xf!!.horizontalAlignment = ha
                } else if (n == "vertical") {
                    val va = sLookup(v, OOXMLConstants.verticalAlignment)
                    xf!!.verticalAlignment = va
                } else if (n == "indent")
                    xf!!.indent = Integer.valueOf(v).toInt()
                else if (n == "wrapText")
                    xf!!.wrapText = true
                else if (n == "textRotation")
                    xf!!.rotation = Integer.valueOf(v).toInt()
                else if (n == "shrinkToFit")
                    xf!!.isShrinkToFit = true
                else if (n == "readingOrder")
                    xf!!.rightToLeftReadingOrder = Integer.valueOf(v).toInt()
            }
        }

    }


    /**
     * look up string index in string array
     *
     * @param s
     * @param sarr String[]
     * @return int index into sarr
     */
    private fun sLookup(s: String?, sarr: Array<String>?): Int {
        if (sarr != null && s != null) {
            for (i in sarr!!.indices) {
                if (s == sarr!![i])
                    return i
            }
        }
        return -1
    }

    /**
     * given a list of all named ranges in the workbook, add all
     *
     * @param bk
     * @param namedRanges
     */
    internal fun addNames(bk: WorkBookHandle, namedRanges: ArrayList<*>) {
        // now input named ranges before processing sheet data
        for (j in namedRanges.indices) {
            val s = namedRanges[j] as Array<String>
            if (!(s[0] == "" && s[2] == "")) {
                try {
                    if (s[0].indexOf("_xlnm") == 0) {    // it's a built-in
                        val sh = s[2].substring(0, s[2].indexOf("!"))
                        //                      String[] addresses= StringTool.splitString(s[2], ",");
                        //                      for (int k= 0; k < addresses.length; k++) {
                        if (s[0] == "_xlnm.Print_Area")
                            try {
                                bk.getWorkSheet(sh)!!.getMysheet()!!.printArea = s[2]//addresses[k]);
                            } catch (e: OutOfMemoryError) {
                                // System.gc();
                                Logger.logWarn("OOXMLAdapter.parse OOME setting PrintArea")
                            }
                        else if (s[0] == "_xlnm.Print_Titles")
                            try {
                                bk.getWorkSheet(sh)!!.getMysheet()!!.printTitles = s[2] //addresses[k]);
                            } catch (e: OutOfMemoryError) {
                                // System.gc();
                                Logger.logWarn("OOXMLAdapter.parse OOME setting PrintTitles")
                            }

                        // TODO: handle other built-in named ranges
                        // _xlnm._FilterDatabase, _xlnm.Criteria, _xlnm.Extract
                        //                      }
                    } else {
                        if (!s[2].startsWith("[")) { // skip names in external workbooks
                            var scope = 0
                            if (s[1] != "") scope = Integer.parseInt(s[1]) + 1
                            Name(bk.workBook, s[0], s[2], scope)
                        }
                    }
                } catch (es: NumberFormatException) {
                    // this is usually a named range that is currently #REF!
                } catch (e: Exception) {
                    //Logger.logErr("OOXMLAdapter.parse: failed creating Named Range:" + e.toString() + s[0] + ":" + s[2]);
                }

            } else
                Logger.logErr("OOXMLAdapter.parse: failed retrieving Named Range")
        }
    }

    /**
     * given a HashMap of inline Strings per cell address, set cell value to string
     * <br></br>NOTE: cells must exist with proper format before calling this method
     *
     * @param bk
     * @param inlineStrs HashMap
     */
    internal fun addInlineStrings(bk: WorkBookHandle, inlineStrs: HashMap<*, *>) {
        val ii = inlineStrs.keys.iterator()
        while (ii.hasNext()) {
            val cellAddr = ii.next() as String
            val s = inlineStrs.get(cellAddr) as String
            val rc = ExcelTools.getRowColFromString(cellAddr)
            try {
                val ch = bk.getCell(cellAddr)   // should have been added already
                ch.`val` = s
            } catch (ex: Exception) {
            }

        }
    }

    /**
     * intercept Sheet adds and hand off to parse event listener as needed
     */
    protected fun sheetAdd(sheet: WorkSheetHandle?, `val`: Any, cachedval: Any?, r: Int, c: Int, fmtid: Int): CellHandle {
        val ch = sheetAdd(sheet!!, `val`, r, c, fmtid)
        (ch!!.cell as Formula).setCachedValue(cachedval)
        return ch
    }

    /**
     * given an array list of every formula in the workbook, iterate list, parse and add approrpriately
     *
     * @param bk
     * @param formulas
     */
    internal fun addFormulas(bk: WorkBookHandle, formulas: ArrayList<*>) {
        // after sheets, now can input formulas
        var sheet: WorkSheetHandle? = null
        val sharedFormulas = HashMap()
        for (j in formulas.indices) {
            val s = formulas[j] as Array<String>
            //formulas:  0=sheetname, 1= cell address, 2=formula including =, 3=shared formula index, 4=array refs, 5=formula type, 6=calculate always flag, 7=format id, 8=cached value
            if (s[0] == "" || s[1] == "" || s.size < 8)
                continue // no address or formula - should ever happen?
            try {
                // for clarity, assign values to most common ops
                val addr = s[1]
                val rc = ExcelTools.getRowColFromString(addr)
                val fStr = s[2]
                var type = s[5]
                var fType = ""
                if (s[5].indexOf('/') > 0) {
                    type = s[5].split("/".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()[0]
                    fType = s[5].split("/".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()[1]
                }
                var fmtid = 0
                try {
                    fmtid = Integer.valueOf(s[7]).toInt()
                } catch (e: Exception) {
                }

                var cachedValue: Any = s[8]
                if (type == "n")
                    try {
                        cachedValue = Integer.valueOf(cachedValue as String)
                    } catch (e: NumberFormatException) {
                        cachedValue = Double(cachedValue as String)
                    }
                else if (type == "b")
                    cachedValue = java.lang.Boolean.valueOf(cachedValue as String)
                // type e -- input calculation exception?
                var ch: CellHandle? = null  // normal case but may be created * as a blank * if part of a merged cell range or dv ...
                try {
                    sheet = bk.getWorkSheet(s[0])
                    ch = sheet!!.getCell(addr)   // if exists, grab it;
                } catch (ex: Exception) {
                }

                if (fStr == "null") { // when would this ever occur?
                    Logger.logWarn("OOXMLAdapter.parse: invalid formula encountered at $addr")
                }

                if (fType == "array") {
                    /*
                       * For a multi-cell formula, the r attribute of the top-left cell
                       * of the range 1 of cells to which that formula applies
                         shall designate the range of cells to which that formula applies
                       */
                    var arrayref: IntArray? = null
                    if (s[4] != null) {   // if has the ref attribute means its the PARENT array formula
                        sheet!!.mysheet!!.addParentArrayRef(s[1], s[4])
                        arrayref = ExcelTools.getRangeRowCol(s[4])
                    } else
                        arrayref = rc
                    /* must enter array formulas for each cell in range denoted by array ref*/
                    for (r in arrayref!![0]..arrayref[2]) {
                        for (c in arrayref[1]..arrayref[3]) {
                            try {
                                ch = sheet!!.getCell(r, c)   // if exists, grab it;
                            } catch (ex: Exception) {
                            }

                            if (ch == null)
                                ch = sheetAdd(sheet, "{$fStr}", cachedValue, r, c, fmtid)
                            else {
                                ch.formatId = fmtid // if exists may be part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
                                ch.setFormula("{$fStr}", cachedValue) // set cached value so don't have to recalculate; just sets cached value if formula is already set
                            }
                        }
                    }
                } else if (fType == "datatable") {
                    if (ch == null)
                        ch = sheetAdd(sheet, fStr, cachedValue, rc[0], rc[1], fmtid)
                    else {
                        ch.formatId = fmtid // if exists may be part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
                        ch.setFormula(fStr, cachedValue)      // set cached value so don't have to recalculate; just sets cached value if formula is already set
                    }
                } else if (fType == "shared" && s[3] != "") {        // meaning that if it's set as shared but doesn't have a shared index, make regular function -- is what excel 2007 does :)
                    // Shared Formulas: there is the "master" shared formula which defines the formula + the range (=ref) of cells that the formula refers to
                    // For references to the shared formula, the si index denotes the shared formula it refers to
                    // one takes the master formula cell, compares with the current cell's address and increments the references in the master shared
                    // formula accordingly -- algorithm of comparison and movement can be tricky
                    val si = Integer.valueOf(s[3])
                    if (!sharedFormulas.containsKey(si)) {
                        // represents the "master" formula of a shared formula, movement is based upon relationship of subsequent cells to this cell
                        if (ch == null)
                            ch = sheetAdd(sheet, fStr, cachedValue, rc[0], rc[1], fmtid)
                        else {
                            ch.formatId = fmtid // if exists may be part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
                            ch.setFormula(fStr, cachedValue)  // set cached value so don't have to recalculate; just sets cached value if formula is already set
                        }
                        // see if it's a 3d range
                        val range = ExcelTools.getRangeCoords(s[3])
                        range[0] -= 1
                        range[2] -= 1
                        val expressionStack = cloneStack(ch.formulaHandle.formulaRec!!.expression!!)
                        sharedFormulas.put(si, arrayOf(expressionStack, rc, range))
                    } else { // found shared formula- means already created; must get original and "move" based on position of this - the child - shared formula
                        val o = sharedFormulas.get(si) as Array<Any>
                        val ss = cloneStack(o[0] as Stack<*>)
                        val rcOrig = o[1] as IntArray
                        Formula.incrementSharedFormula(ss, rc[0] - rcOrig[0], rc[1] - rcOrig[1], o[2] as IntArray)

                        if (ch == null) {
                            ch = sheetAdd(sheet, "=0", null, rc[0], rc[1], fmtid) // add a basic formula; will be "overwritten" by expression, set below
                            ch.setFormula(ss, cachedValue) // must set child shared formulas via expression rather than via formula string as original formula string must be incremented
                        } else {
                            ch.setFormula(ss, cachedValue)   // must set child shared formulas via expression rather than via formula string as original formula string must be incremented
                            ch.formatId = fmtid // if exists may be part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
                        }
                    }
                } else {// it's a regular function
                    if (ch == null)
                    // use parser-aware method
                        ch = sheetAdd(sheet, fStr, cachedValue, rc[0], rc[1], fmtid)
                    else {
                        ch.formatId = fmtid // if exists most likely is part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
                        ch.setFormula(fStr, cachedValue)      // set cached value so don't have to recalculate; just sets cached value if formula is already set
                    }

                }
                if (s[6] != null && ch != null) {  // for formulas such as =TODAY
                    val br = ch.cell
                    if (br is Formula)
                        ch.formulaHandle.calcAlways = true
                }
            } catch (e: FunctionNotSupportedException) {
                Logger.logErr("OOXMLAdapter.parse: failed setting formula " + s[1] + " to cell " + s[0] + ": " + e.toString())
            } catch (e: Exception) {
                Logger.logErr("OOXMLAdapter.parse: failed setting formula " + s[1] + " to cell " + s[0] + ": " + e.toString())
            }

        }
    }

    /**
     * after all sheet data, etc is added, now add pivot tables
     *
     * @param bk          WorkBookHandle
     * @param zip         open ZipFile
     * @param pivotTables Strings name pivot table files within zip
     */
    @Throws(IOException::class)
    internal fun addPivotTables(bk: WorkBookHandle, zip: ZipFile, pivotTables: HashMap<String, WorkSheetHandle>) {
        val ii = pivotTables.keys.iterator()
        while (ii.hasNext()) {
            val key = ii.next()
            val target = zip.getEntry(key)
            /*            target= getEntry(zip,p + "_rels/" + c[1].substring(c[1].lastIndexOf("/")+1)+".rels");
        	ArrayList ptrels= parseRels(wrapInputStream(wrapInputStream(zip.getInputStream(target))));
        	if (ptrels.size() > 1) {	// what could this be?
        		Logger.logWarn("OOXMLReader.parse: Unknown Pivot Table Association: " + ptrels.get(1));
        	}
        	String pcd= ((String[])ptrels.get(0))[1];
        	pcd= pcd.substring(pcd.lastIndexOf("/")+1);
        	Object cacheid= null;
            for (int z= 0; z < pivotCaches.size(); z++) {
        		Object[] o= (Object[]) pivotCaches.get(z);
        		if (pcd.equals(o[0])) {
        			cacheid= o[1];
        			break;
        		}
            }

        	target = getEntry(zip,p + c[1]);*/
            val sheet = pivotTables[key]
            PivotTableDefinition.parseOOXML(bk, /*cacheid, */sheet.getMysheet(), OOXMLAdapter.wrapInputStream(zip.getInputStream(target)))


        }
    }

    /**
     * utility method which looks up a string rid and returns the associated object
     * in a list of Object[] s
     *
     * @param lst source ArrayList
     * @param rid String rid
     * @return
     */
    private fun lookupRid(lst: ArrayList<*>, rid: String): Any? {
        for (i in lst.indices) {
            val o = lst[i] as Array<Any>
            if (rid == o[0])
                return o[1]
        }
        return null
    }

    companion object {

        var parsePivotTables = true        // KSC: TESTING -- only make true in testing


        /**
         * pass-through current OOXML element/file - i.e. save file to external directory on disk
         * because it cannot be processed into our normal BIFF8 machinery
         * also, if current OOXML element has an associated .rels file containing links to other files known as "embeds",
         * store embeds on disk and link information to for later retrieval
         *
         * <br></br>Possible pass-through types:
         * <br></br>props
         * <br></br>exprops
         * <br></br>custprops
         * <br></br>connections
         * <br></br>calc
         * <br></br>vba
         * <br></br>externalLink
         *
         * @param c String[] {type, filename, rid}
         */
        @Throws(IOException::class)
        protected fun handlePassThroughs(zip: ZipFile, bk: WorkBookHandle, parentDir: String, externalDir: String, c: Array<String>) {
            passThrough(zip, parentDir + c[1], externalDir + c[1]) // save the original target file for later re-packaging
            val target = OOXMLAdapter.getEntry(zip, parentDir + "_rels/" + c[1].substring(c[1].lastIndexOf("/") + 1) + ".rels") // is there an associated .rels file??
            if (target == null)
            // no .rels, just link to original OOXML element/file
                bk.workBook!!.addOOXMLObject(arrayOf(c[0], parentDir, externalDir + c[1]))
            else
            // handle embedded objects in \book-level objects (theme embeds, externalLinks
                bk.workBook!!.addOOXMLObject(arrayOf<String>(c[0], parentDir, externalDir + c[1], c[2], null, Arrays.asList<String>(*storeEmbeds(zip, target, parentDir, externalDir)).toString()/* 1.6 only Arrays.toString(storeEmbeds(zip, target, p))*/))
        }

        /**
         * pass-through current sheet-level OOXML element/file - i.e. save file to external directory on disk
         * because it cannot be processed into our normal BIFF8 machinery
         * also, if current OOXML element has an associated .rels file containing links to other files known as "embeds",
         * store embeds on disk and link information to for later retrieval
         *
         * @param c String[] {type, filename, rid}
         */
        @Throws(IOException::class)
        fun handleSheetPassThroughs(zip: ZipFile, bk: WorkBookHandle, sht: Boundsheet, parentDir: String, externalDir: String, c: Array<String>, attrs: String) {
            passThrough(zip, parentDir + c[1], externalDir + c[1]) // save the original target file for later re-packaging
            val target = OOXMLAdapter.getEntry(zip, parentDir + "_rels/" + c[1].substring(c[1].lastIndexOf("/") + 1) + ".rels") // is there an associated .rels file??
            if (target == null)
            // no .rels, just link to original OOXML element/file
                sht.addOOXMLObject(arrayOf(c[0], parentDir, externalDir + c[1], c[2], attrs))
            else
            // handle embedded objects in sheet-level objects (activeX binaries ....)
                sht.addOOXMLObject(arrayOf(c[0], parentDir, externalDir + c[1], c[2], attrs, Arrays.asList<String>(*storeEmbeds(zip, target, parentDir, externalDir)).toString() /* 1.6 only Arrays.toString(storeEmbeds(zip, target, p))*/))
        }


        /**
         * handle OOXML files that we do not process at this time.
         * <br></br>Writes the file in question from zip file fin to file directory file fout
         *
         * @param zip
         * @param fin
         * @param fout
         * @throws IOException
         */
        @Throws(IOException::class)
        fun passThrough(zip: ZipFile, fin: String, fout: String) {
            try {
                val outfile = java.io.File(fout)
                // clean it up
                outfile.deleteOnExit()

                val dirs = outfile.parentFile
                if (dirs != null && !dirs.exists()) {
                    dirs.mkdirs()
                    dirs.deleteOnExit()
                }
                val fos = BufferedOutputStream(FileOutputStream(outfile))
                val fis = OOXMLReader.wrapInputStream(zip.getInputStream(OOXMLReader.getEntry(zip, fin)))
                var i = fis.read()
                while (i != -1) {
                    fos.write(i)
                    i = fis.read()
                }
                fos.flush()
                fos.close()
                dirs!!.delete()
            } catch (e: Exception) {
                // OK for external links for FNFE
            }

        }

        /**
         * get correct path for zip access based on path p and parent directory parentDir
         *
         * @param p
         * @param parentDir
         */
        fun parsePathForZip(p: String, parentDir: String): String {
            var p = p
            var parentDir = parentDir
            if (p.indexOf("/") != 0 || p.indexOf("\\") == 0) {
                while (p.indexOf("..") == 0) {
                    p = p.substring(3)
                    if (parentDir != "" && (parentDir[parentDir.length - 1] == '/' || parentDir[parentDir.length - 1] == '\\'))
                        parentDir = parentDir.substring(0, parentDir.length - 2)
                    var z = parentDir.lastIndexOf("/")
                    if (z == -1)
                        z = parentDir.lastIndexOf("\\")
                    parentDir = parentDir.substring(0, z + 1)
                }

                p = parentDir + p

                //if(DEBUG)
                //  Logger.logInfo("parsePathForZip:"+p);

            } else if (p != "")
                p = p.substring(1)

            return p
        }

        /**
         * retrieves the entire element at the current position in the xpp pullparser,
         * as a string, and advances the pullparser position to the next element
         *
         * @param xpp
         * @return
         */
        fun getCurrentElement(xpp: XmlPullParser): String {
            val el = StringBuffer()
            try {
                var eventType = xpp.eventType
                val elname = xpp.name
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG || eventType == XmlPullParser.TEXT) {
                        el.append(xpp.text)
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val t = xpp.text
                        if (t.indexOf("</") == 0)
                            el.append(t)
                        if (xpp.name == elname)
                            break

                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("OOXMLAdapter.getCurrentElement: $e")
            }

            return el.toString()
        }


        /**
         * XmlPullParser positioned on <is> child of the <c> (cell) element in sheetXXX.xml
         *
         * @param xpp XmlPullParser
         * @return String inline text
         * @throws XmlPullParserException
         * @throws IOException
        </c></is> */
        @Throws(XmlPullParserException::class, IOException::class)
        fun getInlineString(xpp: XmlPullParser): String? {
            var eventType = xpp.next()
            var ret = ""
            while (eventType != XmlPullParser.END_DOCUMENT &&
                    eventType != XmlPullParser.END_TAG &&
                    eventType != XmlPullParser.TEXT) {
                eventType = xpp.next()
            }
            if (eventType == XmlPullParser.TEXT)
                ret = xpp.text

            try {
                return String(ret.toByteArray(), OOXMLAdapter.inputEncoding)
            } catch (e: Exception) {
            }
            // inputEncoding can be null
            return ret
        }

        /**
         * intercept Sheet adds and hand off to parse event listener as needed
         */
        internal fun sheetAdd(sheet: WorkSheetHandle, `val`: Any, r: Int, c: Int, fmtid: Int): CellHandle? {
            return sheet.add(`val`, r, c, fmtid)
        }

        /**
         * take a passthrough element such as vmldrawing or theme which contains embedded objects (images), retrieve and store
         * for later re-writing to zip
         *
         * @param zip    open ZipFile
         * @param target ZipEntry pointing to .rels
         * @param p      path
         * @return String[] array of embeds
         */
        @Throws(IOException::class)
        fun storeEmbeds(zip: ZipFile, target: ZipEntry, p: String, externalDir: String): Array<String> {
            //if(DEBUG) Logger.logInfo("storeEmbeds about to call parseRels on: " + target.toString());

            val embeds = OOXMLAdapter.parseRels(OOXMLAdapter.wrapInputStream(OOXMLAdapter.wrapInputStream(zip.getInputStream(target)))) // obtain a list of image file references for use in later parsing
            Collections.sort(embeds, object : Comparator {
                override fun compare(o1: Any, o2: Any): Int {
                    val a = Integer.valueOf((o1 as Array<String>)[2].substring(3))
                    val b = Integer.valueOf((o2 as Array<String>)[2].substring(3))
                    return a.compareTo(b)
                }
            })
            val strEmbeds = arrayOfNulls<String>(embeds.size)
            for (j in embeds.indices) {
                val v = embeds[j] as Array<String>
                var path = StringTool.getPath(v[1])
                path = parsePathForZip(path, p)
                v[1] = StringTool.stripPath(v[1])
                if (!v[0].equals("externalLinkPath", ignoreCase = true))
                // it's OK for externally referenced book not to be present
                    try {
                        passThrough(zip, path + v[1], externalDir + v[1])    // save the original target file for later re-packaging
                    } catch (e: NullPointerException) {
                        //   if (!v[0].equalsIgnoreCase("externalLinkPath"))  // it's OK for externally referenced book not to be present
                        throw NullPointerException()
                        //}
                    }

                strEmbeds[j] = v[0] + "/" + path + v[1]
            }
            return strEmbeds
        }

        /**
         * retrieve pass-through files (Files not processed by normal WBH channels) for later writing
         *
         * @param zipIn
         * @param externalDir
         */
        fun refreshExternalFiles(zipIn: ZipFile, externalDir: String) {
            val ee = zipIn.entries()
            while (ee.hasMoreElements()) {
                val ze = ee.nextElement()
                val zename = ze.name
                // these elements are handled, all else is not
                if (!//zename.startsWith("xl/drawings") || may be am embed for a chart ...
                        (zename == "xl/workbook.xml" ||
                                zename == "xl/styles.xml" ||
                                zename == "xl/sharedStrings.xml" ||
                                zename == "[Content_Types].xml" ||
                                zename == "_rels/.rels" ||
                                zename == "xl/workbook.xml.rels" ||
                                zename.startsWith("xl/charts") ||
                                zename.startsWith("xl/worksheets"))) {
                    try {
                        val z = zename.lastIndexOf("/")
                        OOXMLReader.passThrough(zipIn, zename, externalDir + zename.substring(z)) // save the original target file for later re-packaging
                    } catch (e: Exception) {
                        Logger.logErr("OOXMLReader.refreshExternalFiles: error retrieving zip entries: $e")
                    }

                }
            }

            // docProps
            // xl/media
            // xl/printerSettings
            // xl/theme
            // xl/activeX
            // NOT: xl/charts, xl/drawings, xl/worksheets, xl/_rels, xl/workbook.xml, xl/styles.xml, comments, sharedStrings
            // ?? drawngs/vmlDrawingX.xml
            // xl/

        }
    }
}
