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
import io.starter.formats.XLS.charts.Chart
import io.starter.formats.XLS.charts.OOXMLChart
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool

import java.io.*
import java.util.*
import java.util.zip.ZipOutputStream

/**
 * Breaking out functionality for writing out of OOXMLAdapter
 */
class OOXMLWriter : OOXMLAdapter(), OOXMLConstants {

    /**
     * generates OOXML for a workbook (see specification above)
     * Creates the ZIP file and writes all files into proper directory structure re: OPC
     *
     * @param workbook
     * @param out      outputStream used by ZipOutputStream
     * @throws IOException
     */
    @Throws(IOException::class)
    fun getOOXML(bk: WorkBookHandle, out: OutputStream) {
        // clear out  ArrayLists ContentLists
        mainContentList = ArrayList()       // main .rels
        wbContentList = ArrayList()         // workbook.xml.rels
        drContentList = ArrayList()         // drawingX.xml.rels
        shContentList = ArrayList()         // sheetX.xml.rels
        sheetsContentList = ArrayList()     // total contents of each sheet
        vmlId = 0                               // reset ordinal id's for external refs
        drawingId = 0
        commentsId = 0
        activeXId = 0
        activeXBinaryId = 0
        printerSettingsId = 0
        oleObjectsId = 0
        chartId = 0
        imgId = 0

        // create XLSX zip file from OutputStream
        zip = ZipOutputStream(out)

        // Wrap the ZipOutputStream in a Writer to handle character encoding
        // setting encoding is important when input encoding is not utf8; writing to utf8 will convert (for example, format strings in styles.xml ...)
        writer = OutputStreamWriter(zip!!, OOXMLAdapter.inputEncoding)   //"UTF-8" );

        // retrive external directory used to store passthrough files
        externalDir = OOXMLAdapter.getTempDir(bk.workBook!!.factory!!.fileName)
        // writeOOXML files to zip
        writeOOXML(bk)
        // write main .rels file
        writeRels(mainContentList, "_rels/.rels")      // TODO: if have doc properties, must add to .rels
        // write [Content_Types].xml
        writeContentType()

        // write out the defferred files to the zip
        writeDeferredFiles()
        writer!!.flush()
        writer!!.close()
        if (zip != null) {
            zip!!.flush()
            zip!!.close()
        }
        if (!bk.workBook!!.factory!!.fileName.endsWith(".tmp"))
            OOXMLAdapter.deleteDir(File(externalDir))
        zip = null
    }


    /**
     * write the deferred files to the zipfile
     *
     * @throws IOException
     */
    @Throws(IOException::class)
    protected fun writeDeferredFiles() {
        val its = deferredFiles.keys.iterator()
        while (its.hasNext()) {
            val k = its.next().toString()
            val fx = deferredFiles[k] as String
            writeFileToZIP(fx, k)
            val fdel = File(fx)
            fdel.deleteOnExit()
            fdel.delete()
        }
    }


    /**
     * handle Sheet-level External References that are pass-throughs and NOT
     * recreated on output: control (activeX), printerSettings, oleObjects
     * Fairly complicated and klugdy ((:
     *
     * @param type
     * @param externalOOXML List is in format of: [0]= type, [1]= pass-through filename,
     * [2] original filename, [3] original rid [, [4]= Extra info, if
     * any, [5]= Embedded files, if any]]
     * @return
     */
    private fun writeSheetLevelExternalReferenceOOXML(out: Writer,
                                                      type: String, externalOOXML: List<*>) {

        val refs = getExternalRefType(type, externalOOXML)   // do we have any of the specific type of references?
        /*
         * because the following methods all write files to the zip it causes
         * problems with the current zipentry.
         *
         * for this reason, we must return the file for later writing, after the
         * current zipentry is closed (aka: sheet1.xml)
         */

        if (refs.size > 0) { // got something
            val ooxml = StringBuffer()
            val rId = -1
            try {
                if (type == "oleObject") {
                    ooxml.append(writeExOOMXLElement("oleObject", refs, false))
                } else if (type == "activeX") {
                    ooxml.append(writeExOOMXLElement("control", refs, false))

                } else if (type.equals("printerSettings", ignoreCase = true)) { // TODO: also: orientation, horizontalDPI
                    ooxml.append(writeExOOMXLElement("pageSetup", refs, true))
                } else { // TESTING - remove when done!
                    // *********************************************************************
                    Logger.logWarn("Unknown External Type $type")
                }
            } catch (e: IOException) {
                Logger.logErr("OOXMLWriter.writeSheetLevelExternalReferenceOOXML: $e")
            }

            try {
                out.write(ooxml.toString())
            } catch (e: Exception) {
            }

        }
    }


    /**
     * search through list of external objects to retrieve those associated with the desired type
     *
     * @param type
     * @param externalOOXML List of previously saved external Objects
     * @return
     */
    private fun getExternalRefType(type: String, externalOOXML: List<*>): ArrayList<*> {
        val refs = ArrayList()
        for (i in externalOOXML.indices) {
            val s = externalOOXML[i] as Array<String>
            if (s != null && s!!.size >= 0) {   // id, dir, filename, rId [, extra info [, embedded file info]]
                if (s!![0].equals(type, ignoreCase = true)) { // got one
                    refs.add(s)
                }
            }
        }
        return refs
    }


    /**
     * most OOXML objects external to workbook are handled here
     * (NOTE: External objects linked to sheets are handled elsewhere)
     * Eventually many of these will (docprops, etc) be created by OpenXLS
     *
     * @param bk
     * @throws IOException
     */
    @Throws(IOException::class)
    private fun writeExternalOOXML(bk: WorkBookHandle) {
        val externalOOXML = bk.workBook!!.ooxmlObjects
        for (i in externalOOXML.indices) {
            val s = externalOOXML[i] as Array<String>
            if (s != null && s!!.size >= 3) {   // id, dir, filename, rid, [extra info], [embedded information]
                val type = s!![EX_TYPE]
                if (type.equals("props", ignoreCase = true) ||
                        type == "exprops" ||
                        type == "custprops" ||
                        type == "connections" ||
                        /* type.equals("calc") || 20081122 KSC: Skip calcChain for now as will error if problems with formulas*/
                        type == "externalLink" ||
                        type == "theme" ||
                        type == "vba") {
                    if (type == "props" || type == "exprops")
                        writeExOOXMLFile(s!!, mainContentList)
                    else
                        writeExOOXMLFile(s!!, wbContentList)
                }
            }
        }
    }

    /**
     * given an external reference String array, parse, obtaining embedded files if present, then writing correct .rels + master file to ZIP
     * also adds master file to ContentList for later inclusion in corresponding .rels
     *
     * @param String[]  external reference String[] { EX_TYPE, EX_PATH, EX_FNAME, (rid)[ , EX_EXTRAINFO [, EX_EMBEDINFO]]
     * @param ArrayList contentList
     * @return rId  int one-based position in contentList for this reference
     * @throws IOException
     */
    @Throws(IOException::class)
    private fun writeExOOXMLFile(s: Array<String>, contentList: ArrayList<*>): Int {
        val p = s[EX_PATH]
        val f = s[EX_FNAME]
        var rId = -1
        val finx = File(f)
        if (finx.exists()) { // external object hasn't already been input into zip
            val fname = getExOOXMLFileName(f)
            if (s.size > EX_EMBEDINFO) { // then linked to external files must copy and account for
                val cl = ArrayList()
                val embeds = StringTool.splitString(s[EX_EMBEDINFO].substring(1, s[EX_EMBEDINFO].length - 1), ",")
                if (embeds != null) {  // EMBEDINFO as: type/path/filename these are usually activeXBinary
                    for (i in embeds.indices) {
                        var pp = embeds[i].trim { it <= ' ' }    // original path + filename
                        val typ = pp.substring(0, pp.indexOf("/"))
                        pp = pp.substring(pp.indexOf("/") + 1)
                        val pth = pp.substring(0, pp.lastIndexOf("/") + 1)
                        pp = pp.substring(pp.lastIndexOf("/") + 1)
                        var ff = pp
                        if (typ != "externalLinkPath") {  // retrieve embeds (previously stored in externaldir)
                            // ensure correct filename and write out to zip,
                            // storing embed info in content list for for .rels
                            ff = getExOOXMLFileName(pp) // ensure proper ordinal number for filename, if necessary
                            deferredFiles[pth + ff] = externalDir + pp // file on disk= externalDir + pp, desired filename in zip= pth+ff
                            cl.add(arrayOf("/$pth$ff", typ))
                            sheetsContentList.addAll(cl)   // most embeds need to be written to main content list
                        } // TODO: externalBooks do not write embeds; other types of external links??? dde, ole ...?
                        else {  // externalLinkPath - exception to the rule - should only be added to wb content list
                            cl.add(arrayOf("/$pth$ff", typ))
                        }
                    }
                }
                writeRels(cl, p + "_rels/" + fname + ".rels")
            }
            deferredFiles[p + fname] = f    // file in zip: p+ fname, file on disk= f (in externalDir)

            contentList.add(arrayOf("/$p$fname", s[EX_TYPE]))
            // remove original external filename from disk and map new name on zip to avoid dups
            rId = contentList.size
        }
        return rId
    }

    /**
     * given an ArrayList containing external reference String arrays, parse, obtaining embedded files if present, then writing correct .rels + master file to ZIP
     * also adds master file to shContentList for later inclusion in sheetX.xml.rels
     * Also generates proper OOXML for element, linking r:id in SheetX.xml to r:id in Sheet.X.xml.rels
     *
     * @param xmlElement String root XMLElement, if !onlyOne, adds a root + "s" e.g. controls or oleObjects
     * @param refs       ArrayList containing list of external reference String[] { EX_TYPE, EX_PATH, EX_FNAME, (rid)[ , EX_EXTRAINFO [, EX_EMBEDINFO]]
     * @param onlyOne    if true, doesn't add root "s" element
     * @return OOXML defining this element
     * @throws IOException
     */
    @Throws(IOException::class)
    private fun writeExOOMXLElement(xmlElement: String, refs: ArrayList<*>, onlyOne: Boolean): String {
        val ooxml = StringBuffer()
        if (!onlyOne) {
            ooxml.append("<" + xmlElement + "s>")
            ooxml.append("\r\n")
        }
        for (i in refs.indices) {
            val s = refs[i] as Array<String>
            if (s.size > EX_EXTRAINFO && s[EX_EXTRAINFO] != null) {   // add associated info, if any
                ooxml.append("<" + xmlElement + " " + s[EX_EXTRAINFO] + " r:id=\"rId" + (shContentList.size + 1) + "\"/>")
                ooxml.append("\r\n")
            } else {
                ooxml.append("<" + xmlElement + " r:id=\"rId" + (shContentList.size + 1) + "\"/>")
                ooxml.append("\r\n")
            }
            writeExOOXMLFile(s, shContentList)
        }
        if (!onlyOne) {
            ooxml.append("</" + xmlElement + "s>")
            ooxml.append("\r\n")
        }
        return ooxml.toString()
    }

    /**
     * generic method for creating a .rels file from a content list array cl
     *
     * @param cl        ArrayList contentList (type, filename)
     * @param relsfname
     * @throws IOException
     */
    @Throws(IOException::class)
    protected fun writeRels(cl: ArrayList<*>?, relsfname: String) {
        if (cl == null || cl.size == 0) return    // don't write a .rels if there are no relationships to track
        val rels = StringBuffer()
        rels.append(OOXMLConstants.xmlHeader)
        rels.append("\r\n")
        rels.append("<Relationships xmlns=\"" + OOXMLConstants.pkgrelns + "\">")
        rels.append("\r\n")

        for (i in cl.indices) {
            rels.append("<Relationship Id=\"rId" + (i + 1))
            // TODO: only external types are hyperlink externalLink?????
            val type = (cl[i] as Array<String>)[1]
            if (type != "hyperlink" && type != "externalLinkPath")
                rels.append("\" Type=\"" + getRelationshipType(type) + "\" Target=\"" + (cl[i] as Array<String>)[0] + "\"/>")
            else
                rels.append("\" Type=\"" + getRelationshipType(type) + "\" Target=\"" + (cl[i] as Array<String>)[0] + "\" TargetMode=\"External\"/>")
            rels.append("\r\n")
        }
        rels.append("</Relationships>")
        rels.append("\r\n")

        // write to tmp
        addDeferredFile(rels, relsfname)
    }

    /**
     * writes all package contents to [Content_Types].xml
     * Package contents are contained in global array lists
     * mainContentList, wbContentList, sheetsContentList, drContentList
     *
     * @throws IOException
     */
    @Throws(IOException::class)
    protected fun writeContentType() {
        val ct = StringBuffer()
        ct.append(OOXMLConstants.xmlHeader)
        ct.append("\r\n")
        ct.append("<Types xmlns=\"" + OOXMLConstants.typens + "\">")
        ct.append("\r\n")
        ct.append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>")
        ct.append("\r\n")
        ct.append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>")
        ct.append("\r\n")
        ct.append("<Default Extension=\"png\" ContentType=\"image/png\"/>")
        ct.append("\r\n")
        ct.append("<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>")
        ct.append("\r\n")
        ct.append("<Default Extension=\"emf\" ContentType=\"image/x-emf\"/>")
        ct.append("\r\n")
        ct.append("<Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/>")
        ct.append("\r\n")
        ct.append("<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>")
        ct.append("\r\n")
        // write ALL content lists here
        for (i in mainContentList.indices) {
            ct.append("<Override PartName=\"" + (mainContentList[i] as Array<String>)[0] + "\" ContentType=\"" + getContentType((mainContentList[i] as Array<String>)[1]) + "\"/>")
            ct.append("\r\n")
        }
        for (i in wbContentList.indices) {
            ct.append("<Override PartName=\"" + (wbContentList[i] as Array<String>)[0] + "\" ContentType=\"" + getContentType((wbContentList[i] as Array<String>)[1]) + "\"/>")
            ct.append("\r\n")
        }
        /* printerSettings and vmlDrawing files are not included in Content_Types.xml - rather, they are handled via <Default Extension> element*/
        /* same goes for images */
        for (i in sheetsContentList.indices) {
            if (!((sheetsContentList[i] as Array<String>)[1] == "printerSettings" ||
                            (sheetsContentList[i] as Array<String>)[1] == "vmldrawing" ||
                            (sheetsContentList[i] as Array<String>)[1] == "hyperlink" ||
                            (sheetsContentList[i] as Array<String>)[1] == "image")) {
                ct.append("<Override PartName=\"" + (sheetsContentList[i] as Array<String>)[0] + "\" ContentType=\"" + getContentType((sheetsContentList[i] as Array<String>)[1]) + "\"/>")
                ct.append("\r\n")
            }
        }
        for (i in drContentList.indices) {
            if ((drContentList[i] as Array<String>)[1] != "image")
            /* image files not included in Content_Type - rather, they are handled via <Default Extension> element*/
                ct.append("<Override PartName=\"" + (drContentList[i] as Array<String>)[0] + "\" ContentType=\"" + getContentType((drContentList[i] as Array<String>)[1]) + "\"/>")
            ct.append("\r\n")
        }

        ct.append("</Types>")
        addDeferredFile(ct, "[Content_Types].xml")
    }

    /**
     * generates OOXML for a workbook
     * Creates the ZIP file and writes all files into proper directory structure
     * Will create either an .xlsx or an .xlsm output, depending upon
     * whether WorkBookHandle bk contains macros
     *
     * @param bk   workbookhandle
     * @param path output filename and path
     */
    @Throws(Exception::class)
    fun getOOXML(bk: WorkBookHandle, path: String) {
        var path = path
        if (!OOXMLAdapter.hasMacros(bk)) {
            path = StringTool.replaceExtension(path, ".xlsx")
            format = WorkBookHandle.FORMAT_XLSX
        } else {    // it's a macro-enabled workbook
            path = StringTool.replaceExtension(path, ".xlsm")
            format = WorkBookHandle.FORMAT_XLSM
        }

        val fout = java.io.File(path)
        val dirs = fout.parentFile
        if (dirs != null && !dirs.exists())
            dirs.mkdirs()
        getOOXML(bk, FileOutputStream(path))
    }

    /**
     * creates sharedStrings.xml if there are entries in the SST
     * and writes it to the root of the OPC ZIP
     *
     * @param bk
     */
    @Throws(IOException::class)
    private fun writeSSTOOXML(bk: WorkBookHandle) {
        // SHAREDSTRINGS.XML
        nextZipEntry("xl/sharedStrings.xml")
        bk.workBook!!.sharedStringTable!!.writeOOXML(writer!!)

        wbContentList.add(arrayOf("/xl/sharedStrings.xml", "sst"))
    }


    /**
     * Calls all the necessary methods to create style OOXML, workbook OOXML, sheet(s) OOXML ...
     * Expects global  zip var to be set to correct zip file
     *
     * @param bk
     * @param output
     * @throws IOException
     * @see getOOXML
     */
    @Throws(IOException::class)
    protected fun writeOOXML(bk: WorkBookHandle) {
        val origcalcmode = bk.workBook!!.calcMode
        bk.workBook!!.calcMode = WorkBook.CALCULATE_EXPLICIT  // don't recalculate
        writeExternalOOXML(bk)     // all workbook-level objects we cannot handle at this point such as docProps, themes, or are pass-throughs, such as vbaprojects ...
        writeSSTOOXML(bk)
        // writeStylesOOXML AFTER sheet OOXML in order to capture any dxf's  (differential xf's used in conditional formatting and others)
        // but ensure that workbook.xml.rels knows about the file styles.xml ***
        wbContentList.add(arrayOf("/xl/styles.xml", "styles"))
        writeWorkBookOOXML(bk)
        val wsh = bk.workSheets
        bk.workBook!!.dxfs = null // rebuild
        for (i in wsh!!.indices) {
            writeSheetOOXML(bk, wsh[i], i)
        }
        writeStylesOOXML(bk)   // must do AFTER sheet OOXML to capture any dxf's (differential xf's used in conditional formatting and others)
        bk.workBook!!.calcMode = origcalcmode // reset

    }


    /**
     * Creates Styles.xml with font and xf information, and writes it to the root directory of the OPC ZIP
     *
     *
     * A Style is a named collection of formatting elements.
     * A cell style can specify number format, cell alignment, font information, cell border specifications, colors, and background / foreground fills.
     * Table styles specify formatting elements for the regions of a table (e.g. make the header row & totals bold face, and apply light
     * gray fill to alternating rows in the data portion of the table to achieve striped or banded rows).
     * PivotTable styles specify formatting elements for the regions of a PivotTable (e.g. 1st & 2nd level subtotals, row axis,
     * column axis, and page fields).
     *
     *
     * A Style can specify color, fonts, and shape effects directly, or these elements can be referenced indirectly by
     * referring to a Theme definition. Using styles allows for quicker application of formatting and more consistently
     * stylized documents.
     * Themes define a set of colors, font information, and effects on shapes (including Charts). If a style or
     * formatting element defines its color, font, or effect by referencing a theme, then picking a new theme switches
     * all the colors, fonts, and effects for that formatting element.
     * Applying Direct Formatting means that particular elements of formatting (e.g. a bold font face or a number
     * format) have been applied, but the elements of formatting have been chosen individually instead of
     * collectively by choosing a named Style. Note that when applying direct formatting, themes can still be
     * referenced, causing those elements to change as the theme is changed.
     *
     *
     *
     *
     * Styles.xml may contain:
     *
     *
     * <styleSheet>          ROOT
    </styleSheet> *
     *
     * borders (Borders) 3.8.5
     * cellStyles (Cell Styles) for a named cell style ... built-in styles are referenced here by name e.g. Heading 1, Normal (= Default) ...
     * [name= "" buildinId="xfid"]
     * cellStyleXfs (Formatting Records) Master formatting xf's, can override others ...
     * a cell can have both direct formatting e.g. bold, and a cell style,
     * therefore, both cellStyleXf + cell xf records must be read
     * cellXfs (Cell Formats)    Cell xf's; cells in the <sheet> section of Workbook.xml reference the 0-based xf index here
     * colors (Colors)
     * dxfs (Formats)            Differential formatting, for all non-cell elements
     * extLst (Future Feature Data Storage Area)
     * fills (Fills)
     * fonts (Fonts)
     * numFmts (Number Formats)
     * tableStyles (Table Styles)
     *
     * @param bk
    </sheet> */
    @Throws(IOException::class)
    private fun writeStylesOOXML(bk: WorkBookHandle) {

        val stylesooxml = StringBuffer()
        stylesooxml.append(OOXMLConstants.xmlHeader)
        stylesooxml.append("<styleSheet xmlns=\"" + OOXMLConstants.xmlns + "\">")
        stylesooxml.append("\r\n")

        // Now create nodes for various XF elements
        val xfs = bk.workBook!!.xfrecs

        val cellxfs = ArrayList()   // references various style source elements for ea xf
        val fills = ArrayList()
        val borders = ArrayList()
        val numfmts = ArrayList()
        val fonts = ArrayList()

        // input default fills -- both appear to be required
        fills.add(Fill.getOOXML(0, -1, -1)) // none
        fills.add(Fill.getOOXML(17, -1, -1)) // gray125
        // input default borders element (= no borders)
        borders.add(Border.getOOXML(intArrayOf(-1, -1, -1, -1, -1), intArrayOf(0, 0, 0, 0, 0)))

        // Iterate the xf's and populate values
        for (i in xfs.indices) {
            val xf = xfs[i] as Xf
            addXFToStyle(xf, cellxfs, fills, borders, numfmts, fonts)

        }

        //** stylesheet element contains an ordered SEQUENCE of elements **//

        // Number formats
        if (numfmts.size > 0) {
            stylesooxml.append("<numFmts count=\"" + numfmts.size + "\">")
            stylesooxml.append("\r\n")
            for (i in numfmts.indices) {
                stylesooxml.append(numfmts.get(i) as String)
                stylesooxml.append("\r\n")
            }
            stylesooxml.append("</numFmts>")
            stylesooxml.append("\r\n")
        }

        // fonts element
        stylesooxml.append("<fonts count=\"" + fonts.size + "\">")
        stylesooxml.append("\r\n")
        for (i in fonts.indices) {
            stylesooxml.append(fonts.get(i) as String)
            stylesooxml.append("\r\n")
        }
        stylesooxml.append("</fonts>")
        stylesooxml.append("\r\n")

        // fill patterns element - always has two defaults
        stylesooxml.append("<fills count=\"" + fills.size + "\">")
        stylesooxml.append("\r\n")
        for (i in fills.indices) {
            stylesooxml.append(fills.get(i) as String)
            stylesooxml.append("\r\n")
        }
        stylesooxml.append("</fills>")
        stylesooxml.append("\r\n")

        //borders element - has one default
        stylesooxml.append("<borders count=\"" + borders.size + "\">")
        stylesooxml.append("\r\n")
        for (i in borders.indices) {
            stylesooxml.append(borders.get(i) as String)
            stylesooxml.append("\r\n")
        }
        stylesooxml.append("</borders>")
        stylesooxml.append("\r\n")


        // cellXfs
        stylesooxml.append("<cellXfs count=\"" + cellxfs.size + "\">")
        stylesooxml.append("\r\n")
        for (i in cellxfs.indices) {
            // xfId= 0 based index of an xf record contained in cellStyleXfs corresponding to the
            // cell style applied to the cell (only for celLXfs, not cellStyleXfs)
            stylesooxml.append("<xf ")
            val refs = cellxfs.get(i) as IntArray
            // all id refs are 0-based
            val ftId = refs[0]    // font ref
            val fId = refs[1]     // fill ref
            val bId = refs[2]     ///border ref
            val nId = refs[3]     // number format ref
            val ha = refs[4]
            val va = refs[5]
            val wr = refs[6]
            val ind = refs[7]
            val rot = refs[8]
            val hidden = refs[9]
            val locked = refs[10]
            val shrink = refs[11]
            val rtoleft = refs[12]

            if (nId > -1)
                stylesooxml.append(" numFmtId=\"$nId\"")
            if (ftId > -1)
                stylesooxml.append(" fontId=\"$ftId\"")
            if (fId > -1)
                stylesooxml.append(" fillId=\"$fId\"")
            if (fId > 0)
                stylesooxml.append(" applyFill=\"1\"")
            if (bId > -1)
                stylesooxml.append(" borderId=\"$bId\"")

            // TODO: shrinkToFit ...
            val alignblock = ha != 0 || va != 2 || wr != 0 || ind != 0 || rot != 0
            val protectblock = hidden == 1 || locked == 0
            if (alignblock || protectblock)
                stylesooxml.append(">")
            if (ha != 0 || va != 2 || wr != 0 || ind != 0 || rot != 0) {
                stylesooxml.append("<alignment")
                if (ha != 0)
                //default=general
                    stylesooxml.append(" horizontal=\"" + OOXMLConstants.horizontalAlignment[ha] + "\"")
                if (va != 2)
                //default= bottom
                    stylesooxml.append(" vertical=\"" + OOXMLConstants.verticalAlignment[va] + "\"")
                if (wr == 1)
                    stylesooxml.append(" wrapText=\"1\"")
                if (ind > 0)
                    stylesooxml.append(" indent=\"$ind\"")
                if (rot != 0)
                    stylesooxml.append(" textRotation=\"$rot\"")
                if (shrink == 1)
                    stylesooxml.append(" shrinkToFit=\"1\"")
                if (rtoleft != 0)
                    stylesooxml.append(" readingOrder=\"$rtoleft\"")
                stylesooxml.append("/>\r\n")
            }
            if (hidden == 1 || locked == 0) { // if not the default protection settings, add protection element
                stylesooxml.append("<protection")
                if (hidden == 1)
                    stylesooxml.append(" hidden=\"1\"")
                if (locked == 0)
                    stylesooxml.append(" locked=\"0\"")
                stylesooxml.append("/>\r\n")
            }
            if (alignblock || protectblock)
                stylesooxml.append("</xf>\r\n")
            else
                stylesooxml.append("/>\r\n")
        }
        stylesooxml.append("</cellXfs>")
        stylesooxml.append("\r\n")

        // cellStyles -- for named styles -- NEEDED??

        // dxf's -- incremental style info
        if (bk.workBook!!.dxfs != null) {
            val dxfs = bk.workBook!!.dxfs
            if (dxfs!!.size > 0) {
                stylesooxml.append("<dxfs count=\"" + dxfs.size + "\">")
                for (i in dxfs.indices) {
                    stylesooxml.append((dxfs[i] as Dxf).ooxml)
                }
                stylesooxml.append("</dxfs>")
            }
        }
        // NOTE: indexed colors are depreciated and represent the hard-coded default palate but necessary for proper color translation
        // Indexed Colors:
        stylesooxml.append("<colors>")
        stylesooxml.append("\r\n")
        stylesooxml.append("<indexedColors>")
        stylesooxml.append("\r\n")
        for (i in 0 until bk.workBook!!.colorTable.size) {
            stylesooxml.append("<rgbColor rgb=\"" + "00" + FormatHandle.colorToHexString(bk.workBook!!.colorTable[i]).substring(1) + "\"/>")
            stylesooxml.append("\r\n")
        }
        stylesooxml.append("</indexedColors>")
        stylesooxml.append("\r\n")
        stylesooxml.append("</colors>")
        stylesooxml.append("\r\n")
        stylesooxml.append("</styleSheet>")
        addDeferredFile(stylesooxml, "xl/styles.xml")
    }


    /**
     * Using the xf record passed in, populate the required variables for writing out to
     * styles.xml
     *
     * @param xfs
     * @param cellxfs
     * @param fills
     * @param borders
     * @param numfmts
     * @param fonts
     */
    private fun addXFToStyle(xf: Xf, cellxfs: ArrayList<*>,
                             fills: ArrayList<*>, borders: ArrayList<*>, numfmts: ArrayList<*>,
                             fonts: ArrayList<*>) {

        val refs = IntArray(13)

        // fonts
        var s = xf.font!!.ooxml
        var id = fonts.indexOf(s)
        if (id == -1) {
            fonts.add(s)
            refs[0] = fonts.size - 1
        } else {
            refs[0] = id
        }

        // fills
        id = 0
        if (xf.getFill() != null) {
            s = xf.getFill()!!.ooxml
            id = fills.indexOf(s)
        } else if (xf.fillPattern > 0) {
            s = Fill.getOOXML(xf)
            id = fills.indexOf(s)
        }
        if (id == -1) {
            fills.add(s)
            refs[1] = fills.size - 1
        } else
            refs[1] = id


        // borders
        s = Border.getOOXML(xf/*fh[i])*/)
        id = borders.indexOf(s)
        if (id == -1) {
            borders.add(s)
            refs[2] = borders.size - 1
        } else
            refs[2] = id

        if (xf.ifmt > FormatConstants.BUILTIN_FORMATS.size) {    // only input user defined formats ...
            s = NumFmt.getOOXML(xf)
            id = numfmts.indexOf(s)
            if (id == -1)
                numfmts.add(s)
        }

        refs[3] = xf.ifmt.toInt()
        refs[4] = xf.horizontalAlignment
        refs[5] = xf.verticalAlignment
        refs[6] = if (xf.wrapText) 1 else 0
        refs[7] = xf.indent
        refs[8] = xf.rotation
        refs[9] = if (xf.isFormulaHidden) 1 else 0
        refs[10] = if (xf.isLocked) 1 else 0
        refs[11] = if (xf.isShrinkToFit) 1 else 0
        refs[12] = xf.rightToLeftReadingOrder

        cellxfs.add(refs)        // link formatHandles to referenced formats

    }


    /**
     * Creates workbook.xml containing worksheet information, and writes it out to
     * the root of the OPC ZIP
     *
     *
     * Format of the workbook.xml:
     * (all elements are OPTIONAL except for <sheets>)
     *
     * <bookViews>       Window Position and height/width, Filter Options .. No limit to how many are defined
     * <calcPr>          Stores calculation status and details
     * <customWorkbookViews>
     * <definedNames><definedName name="" [comment="" hidden="1/0" localSheetId="" for external refs]>RANGE or FORMULA</definedName>
     * <extLst>          Future Feature- Data Storage Area
     * <externalReferences>
     * <fileRecoveryPr>  File Recovery Properties
     * <fileSharing>     Specifies pwd + username
     * <fileVersion>     tracks versions
     * <functionGroups>
     * <oleSize>         Embedded Object Size
     * <pivotCaches>     Represents a cache of data for pivot tables and formulas
     * <sheets><sheet r:id="relationship id" name="unique sheet name" sheetId="#" [state="visible" ]></sheet></sheets>          -- required
     * <smartTagPr>
     * <webPublishing>   Attibutes related to publishing on the web
     * <workbookPr>      Workbook Properties: date1904,showObjects ...
     * <workbookProtection>
     *
     * @param bk
    </workbookProtection></workbookPr></webPublishing></smartTagPr></pivotCaches></oleSize></functionGroups></fileVersion></fileSharing></fileRecoveryPr></externalReferences></extLst></definedNames></customWorkbookViews></calcPr></bookViews></sheets> */
    @Throws(IOException::class)
    private fun writeWorkBookOOXML(bk: WorkBookHandle) {

        // create Zip Entry
        nextZipEntry("xl/workbook.xml")

        // WORKBOOK.XML
        writer!!.write(OOXMLConstants.xmlHeader)
        writer!!.write("\r\n")

        // namespace
        writer!!.write("<workbook xmlns=\"" + OOXMLConstants.xmlns + "\" xmlns:r=\"" + OOXMLConstants.relns + "\">")
        writer!!.write("\r\n")

        // IF MACRO-ENABLED, MUST HAVE CODENAME       // TODO: other attributes
        if (bk.workBook!!.codename != null) {
            writer!!.write("<workbookPr codeName=\"" + bk.workBook!!.codename + "\"/>")
            writer!!.write("\r\n")
        }

        // BOOKVIEW
        writer!!.write("<bookViews>")
        writer!!.write("\r\n")
        // TODO: Only 1?  TODO: Handle other workbookview options
        writer!!.write("<workbookView")
        writer!!.write(" firstSheet=\"" + bk.workBook!!.firstSheet + "\"")
        writer!!.write(" activeTab=\"" + bk.workBook!!.selectedSheetNum + "\"")
        if (!bk.showSheetTabs()) writer!!.write(" showSheetTabs=\"0\"")
        writer!!.write("/>\r\n")
        writer!!.write("</bookViews>")
        writer!!.write("\r\n")

        // IDENTIFY SHEETS
        writer!!.write("<sheets>")
        writer!!.write("\r\n")
        val wsh = bk.workSheets
        for (i in wsh!!.indices) {
            val s = "sheet" + (i + 1)        //Write SheetXML to SheetX.xml, 1-based
            writer!!.write("<sheet name=\"" + OOXMLAdapter.stripNonAscii(wsh[i].sheetName) + "\" sheetId=\"" + (i + 1) + "\" r:id=\"rId" + (i + 1) + "\"")
            if (wsh[i].veryHidden)
                writer!!.write(" state=\"veryHidden\"")
            else if (wsh[i].hidden)
                writer!!.write(" state=\"hidden\"")
            writer!!.write("/>")
            writer!!.write("\r\n")
            wbContentList.add(i, arrayOf("/xl/worksheets/$s.xml", "sheet")) // make sure rId in workbook.xml matches workbook.xml.rels
        }
        writer!!.write("</sheets>")
        writer!!.write("\r\n")

        var rId = wsh.size  // start counting after sheets
        // EXTERNAL LINKS AFTER SHEETS
        if (getExternalRefType("externalLink", bk.workBook!!.ooxmlObjects).size > 0) { // has external refs
            val refs = getExternalRefType("externalLink", bk.workBook!!.ooxmlObjects)
            writer!!.write("<externalReferences>")
            writer!!.write("\r\n")
            for (i in refs.indices) {
                val r = refs[i] as Array<String>
                r[3] = "rId" + (rId + 1)  // ensure rId is correct
                refs.removeAt(i)
                refs.add(i, r)
                writer!!.write("<externalReference r:id=\"rId" + ++rId + "\"/>")
                writer!!.write("\r\n")
            }
            writer!!.write("</externalReferences>")
            writer!!.write("\r\n")
        }

        // NAMES AFTER EXTERNAL REFS
        // TODO:  add handling for name parameters
        val names = bk.workBook!!.names
        if (names != null && names.size > 0) {
            writer!!.write("<definedNames>")
            writer!!.write("\r\n")
            for (i in names.indices) {
                val s = OOXMLAdapter.stripNonAsciiRetainQuote(names[i].expressionString.substring(1)).toString() //avoid "="
                if (s != null && s.length != 0 && !s.startsWith("#REF!")) {
                    if (!names[i].isBuiltIn) {
                        writer!!.write("<definedName name=\"" + OOXMLAdapter.stripNonAscii(names[i].toString()) + "\"")
                        if (names[i].getItab() > 0) {
                            writer!!.write(" localSheetId=\"" + (names[i].getItab() - 1) + "\">")
                        } else {
                            writer!!.write(">")
                        }
                    } else {
                        //if (names[i].getBuiltInType()==Name.PRINT_TITLES) { // must set localsheetid
                        writer!!.write("<definedName name=\"" + OOXMLConstants.builtInNames[names[i].builtInType] + "\"")
                        if (names[i].getItab() > 0)
                            writer!!.write(" localSheetId=\"" + (names[i].getItab() - 1) + "\">")
                        else
                            writer!!.write(">")
                    }
                    writer!!.write(s)
                    writer!!.write("</definedName>")
                    writer!!.write("\r\n")
                }
            }
            writer!!.write("</definedNames>")
            writer!!.write("\r\n")
        }

        writer!!.write("</workbook>")
        writer!!.write("\r\n")


        // add workbook.xml to content list (for [Content_Types].xml)
        if (format == WorkBookHandle.FORMAT_XLTM) { // macro-enabled template
            mainContentList.add(arrayOf("/xl/workbook.xml", "documentTemplateMacroEnabled"))
        } else if (format == WorkBookHandle.FORMAT_XLSM || OOXMLAdapter.hasMacros(bk)) {// format can be XLSM even though it does not contain macros ((:
            mainContentList.add(arrayOf("/xl/workbook.xml", "documentMacroEnabled"))
            format = WorkBookHandle.FORMAT_XLSM   // ensure flag is set properly - macro-enabled workbook
        } else if (format == WorkBookHandle.FORMAT_XLTX) {  // template
            mainContentList.add(arrayOf("/xl/workbook.xml", "documentTemplate"))
        } else
        // regular xlsx workbook
            mainContentList.add(arrayOf("/xl/workbook.xml", "document"))

        writeRels(wbContentList, "xl/_rels/workbook.xml.rels")   // write workbook.xml.rels
    }

    /***
     * Handles preliminary writing of worksheets, essentially everything before <row> elements start
     *
     * @param sheet
     * @param bk
     * @param id
     * @throws IOException
    </row> */
    @Throws(IOException::class)
    protected fun writeSheetPrefix(sheet: WorkSheetHandle, bk: WorkBookHandle, id: Int) {
        val slx = "xl/worksheets/sheet" + (id + 1) + ".xml"
        nextZipEntry(slx)


        writer!!.write(OOXMLConstants.xmlHeader)
        writer!!.write("\r\n")
        writer!!.write("<worksheet xmlns=\"" + OOXMLConstants.xmlns + "\" xmlns:r=\"" + OOXMLConstants.relns + "\">")

        if (sheet.mysheet!!.sheetPr != null)
        // sheet properties
            writer!!.write(sheet.mysheet!!.sheetPr!!.ooxml)

        // dimensions    // TODO: Deal with MAXROWS MAXCOLS in various Excel Versions
        var last = sheet.lastCol - 1
        if (last == WorkBook.MAXCOLS - 1)
            last = XLSConstants.MAXCOLS - 1    // 20081204 KSC: eventually WorkBook.MAXCOLS will == Excel-7 MAXCOLS, but until, convert
        if (last < 0)
            last = 0
        val d = ExcelTools.formatLocation(intArrayOf(sheet.firstRow, sheet.firstCol)) + ":" +
                ExcelTools.formatLocation(intArrayOf(sheet.lastRow, last))
        writer!!.write("<dimension ref=\"$d\"/>")
        writer!!.write("\r\n")


        // Sheet View Properties
        writer!!.write("<sheetViews>")
        writer!!.write("\r\n")    // TODO: it's possible to have multiple sheetViews
        if (sheet.mysheet!!.sheetView == null)
            sheet.mysheet!!.sheetView = SheetView()

        // TODO: finish options:  colorId, defaultGridColor, rightToLeft
        // showFormulas, showRuler, showWhiteSpace, view
        // zoomScaleNormal, zoomScalePageLayoutView, zoomScaleSheetLayoutView
        // showFormulas, default= false
        if (!sheet.showGridlines)
            sheet.mysheet!!.sheetView!!.setAttr("showGridLines", "0") // default= true
        if (!sheet.showSheetHeaders)
            sheet.mysheet!!.sheetView!!.setAttr("showRowColHeaders", "0") // default= true
        if (!sheet.showZeroValues)
            sheet.mysheet!!.sheetView!!.setAttr("showZeros", "0") // default= true
        // rightToLeft
        if (sheet.selected)
            sheet.mysheet!!.sheetView!!.setAttr("tabSelected", "1")           // default= false
        else
            sheet.mysheet!!.sheetView!!.removeSelection()  // in case previously selected, remove any seletions
        //       if (sheet.getTopLeftCell()!=null) { sheet.getMysheet().getSheetView().setAttr("topLeftCell", sheet.getTopLeftCell()); }

        // showRuler
        //       if (!sheet.getShowOutlineSymbols()) sheet.getMysheet().getSheetView().setAttr("showOutlineSymbols", "0");    // default= true
        // defaultGridColor, showWhiteSpace, view, topLeftCell, colorId
        if (sheet.zoom.toDouble() != 1.0)
            sheet.mysheet!!.sheetView!!.setAttr("zoomScale", (sheet.zoom * 100).toInt().toString())

        // zoomScalePageLayoutView, zoomScaleSheetLayoutView
        sheet.mysheet!!.sheetView!!.setAttr("workbookViewId", "0")        // TODO: may be other workbookviews, can't always assume 0

        writer!!.write(sheet.mysheet!!.sheetView!!.ooxml)
        writer!!.write("\r\n")
        writer!!.write("</sheetViews>")
        writer!!.write("\r\n")

        // Sheet Format Properties
        writer!!.write("<sheetFormatPr")
        if (sheet.mysheet!!.defaultColumnWidth > -1)
            writer!!.write(" defaultColWidth=\"" + sheet.mysheet!!.defaultColumnWidth + "\"")
        writer!!.write(" defaultRowHeight=\"" + sheet.mysheet!!.defaultRowHeight + "\"")    // required
        if (sheet.mysheet!!.hasCustomHeight())
            writer!!.write(" customHeight=\"1\"")
        if (sheet.mysheet!!.hasZeroHeight())
            writer!!.write(" zeroHeight=\"1\"")
        if (sheet.mysheet!!.hasThickTop())
            writer!!.write(" thickTop=\"1\"")
        if (sheet.mysheet!!.hasThickBottom())
            writer!!.write(" thickBottom=\"1\"")
        writer!!.write("/>")
        writer!!.write("\r\n")

        // Columns
        writer!!.write(getColOOXML(bk, sheet).toString())

        // Sheet Data - rows and cells
        writer!!.write("<sheetData>")
        writer!!.write("\r\n")

    }

    /**
     * writeSheetOOXML writes XML data in SheetML specificaiton to [worksheet].xml
     *
     *
     * Main portion is the <sheetData> section, containing all row and cell information
    </sheetData> *
     *
     * format of SheetML:
     * <worksheet>     ROOT
     * autoFilter (AutoFilter Settings)   Hides rows based upon criteria
     * cellWatches (Cell Watch Items)
     * colBreaks (Vertical Page Breaks)
     * cols (Column Information)
     * conditionalFormatting (Conditional Formatting)
     * controls (Embedded Controls)
     * customProperties (Custom Properties)
     * customSheetViews (Custom Sheet Views)
     * dataConsolidate (Data Consolidate)
     * dataValidations (Data Validations)
     * dimension (Worksheet Dimensions)
     * drawing (Drawing)
     * extLst (Future Feature Data Storage Area)
     * headerFooter (Header Footer Settings)
     * hyperlinks (Hyperlinks)
     * ignoredErrors (Ignored Errors)
     * legacyDrawing (Legacy Drawing Reference)
     * legacyDrawingHF (Legacy Drawing Reference in Header Footer)
     * mergeCells (Merge Cells)
     * oleObjects (Embedded Objects)
     * pageMargins (Page Margins)
     * pageSetup (Page Setup Settings)
     * phoneticPr (Phonetic Properties)
     * picture (Background Image)
     * printOptions (Print Options)
     * protectedRanges (Protected Ranges)
     * rowBreaks (Horizontal Page Breaks (Row))
     * scenarios (Scenarios)
     * sheetCalcPr (Sheet Calculation Properties)
     * sheetData (Sheet Data)
     * sheetFormatPr (Sheet Format Properties)
     * sheetPr (Sheet Properties)
     * sheetProtection (Sheet Protection Options)
     * sheetViews (Sheet Views)
     * smartTags (Smart Tags)
     * sortState (Sort State)
     * tableParts (Table Parts)
     * webPublishItems (Web Publishing Items)  *
    </worksheet> *
     *
     *
     *
     * ORDER OF THE ABOVE:
     * (sheet properties -- all optional)
     * <sheetPr filterMode="1"></sheetPr>       indicates that an autofilter has been applied
     * <dimension ref="RANGE"></dimension>        indicates the used range on the sheet (there should be no data or formulas outside this range)
     * <sheetViews>                    indicates which cell and sheet are active
     * <sheetView tabSelected="1" workbookViewId="0">
     * <selection activeCell="B3" sqref="B3"></selection>
    </sheetView> *
    </sheetViews> *
     * <sheetFormatPr defaultRowHeight="15"></sheetFormatPr>      default row height
     * <cols>
     * <col min="1" max="1" width="12.85546875" bestFit="1" customWidth="1"></col>
     * <col min="3" max="3" width="3.28515625" customWidth="1"></col>
     * <col min="4" max="4" width="11.140625" bestFit="1" customWidth="1"></col>
     * <col min="8" max="8" width="17.140625" style="1" customWidth="1"></col>
    </cols> *
     *
     *
     * **     <sheetData>     cell table - specifies rows and cells with types and values (see below) -- REQUIRED
    </sheetData> *
     *
     * ("Supporting Features" -- all optional)
     * <sheetProtection objects="0" scenarios="0"></sheetProtection>
     * <autoFilter ref="D5:H11">
     * <filterColumn colId="0">
     * <customFilters and="1">
     * <customFilter operator="greaterThan" val="0"></customFilter>
     * <mergeCells>
     * <phoneticPr>
     * <conditionalFormatting>
     * <printOptions></printOptions>
     * <dataValidations>
     * <hyperlinks>
     * <printOptions>
     * <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"></pageMargins>
     * <pageSetup orientation="portrait" horizontalDpi="300" verticalDpi="300"></pageSetup>
     * <headerFooter>
     * <rowBreaks>
     * <colBreaks>
     * <customProperties>
     * <cellWatches>
     * <ignoredErrors>
     * <smartTags>
     * <drawing>
     * <legacyDrawing>
     * <legacyDrawingHF>
     * <picture>
     * <oleObjects>
     * <controls>
     * <webPublishItems>
     * <tableParts>
     * <extLst>
     *
     * @param bk
     * @param sheet
    </extLst></tableParts></webPublishItems></controls></oleObjects></picture></legacyDrawingHF></legacyDrawing></drawing></smartTags></ignoredErrors></cellWatches></customProperties></colBreaks></rowBreaks></headerFooter></printOptions></hyperlinks></dataValidations></conditionalFormatting></phoneticPr></mergeCells></customFilters></filterColumn></autoFilter> */
    @Throws(IOException::class)
    protected fun writeSheetOOXML(bk: WorkBookHandle, sheet: WorkSheetHandle, id: Int) {
        // Sst sst= bk.getWorkBook().getSharedStringTable();
        val hyperlinks = ArrayList()
        // SHEETxx.XML
        this.writeSheetPrefix(sheet, bk, id)
        val rows = sheet.rows
        for (xd in rows.indices) {
            val row = rows[xd]
            try { // note: row #, col #'s are 1-based, sst and style index are 0-based
                this.writeRow(row, hyperlinks)
                //} catch (RowNotFoundException re) {
                // do nothing
            } catch (e: Exception) {
                Logger.logErr("OOXMLWriter.writeSheetOOXML writing rows: $e")
                e.printStackTrace()
            }

        }
        writer!!.write("</sheetData>")
        writer!!.write("\r\n")

        // after sheetData include "supporting features"
        // *******************************************************************************************************
        // In Order:
        // sheetCalcPr
        // sheetProtection
        if (sheet.protected) {
            val pwd = sheet.hashedProtectionPassword
            if (pwd != null)
                writer!!.write("<sheetProtection password=\"" + sheet.hashedProtectionPassword + "\" sheet=\"1\" objects=\"1\" scenarios=\"1\"/>")
            else
                writer!!.write("<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"/>")
            writer!!.write("\r\n")
        }


        // protectedRanges
        // scenarios
        // autoFilter
        if (sheet.mysheet!!.ooAutoFilter != null)
        // TODO: Merge with 2003 AutoFilter
            writer!!.write(sheet.mysheet!!.ooAutoFilter!!.ooxml)
        // sortState
        // dataConsolidation
        // customSheetViews
        // mergeCells
        this.writeMergedCellRecords(sheet)

        // phoneticPr
        // conditionalFormatting
        if (sheet.mysheet!!.conditionalFormats != null) {
            val condfmts = sheet.mysheet!!.conditionalFormats
            val priority = IntArray(1)
            priority[0] = 1
            for (i in condfmts!!.indices) {
                val cfmt = (condfmts[i] as Condfmt).getOOXML(bk, priority)
                writer!!.write(cfmt)
                writer!!.write("\r\n")
            }
        }
        // dataValidations
        if (sheet.hasDataValidations())
            writer!!.write(sheet.mysheet!!.dvalRec!!.ooxml)
        // hyperlinks
        if (hyperlinks.size > 0) {
            writer!!.write("<hyperlinks>")
            writer!!.write("\r\n")
            for (i in hyperlinks.indices) {
                val s = hyperlinks.get(i) as Array<String>
                if (s[2] != "")
                // has a description
                    writer!!.write("<hyperlink ref=\"" + s[0] + "\" r:id=\"rId"
                            + (shContentList.size + 1) + "\" display=\""
                            + s[2] + "\"/>")
                else
                    writer!!.write("<hyperlink ref=\"" + s[0] + "\" r:id=\"rId"
                            + (shContentList.size + 1) + "\"/>")
                writer!!.write("\r\n")
                shContentList.add(arrayOf(s[1], "hyperlink"))
            }
            hyperlinks.clear()
            writer!!.write("</hyperlinks>")
            writer!!.write("\r\n")
        }

        // now handle any external ooxml object references (in required order):
        // printerOptions
        // pageMargin
        // pageSetup

        // headerFooter
        // rowBreaks
        // colBreaks
        // customProperties
        // cellWatches
        // ignoredErrors
        // smartTags
        val externalOOXML = sheet.mysheet!!.ooxmlObjects
        // printerOptions
        writeSheetLevelExternalReferenceOOXML(writer, "printerSettings", externalOOXML)
        // drawing objects linked to this sheet - charts, shapes ...
        writeDrawingObjects(writer, sheet, bk)
        // legacy drawing object (vml)
        if (writeLegacyDrawingObjects(writer, sheet, bk)) {// legacyDrawing objects, linked to either oleObject embeds or control embeds
            writer!!.write("<legacyDrawing r:id=\"rId" + shContentList.size + "\"/>\r\n")
        }
        //  oleObjects
        writeSheetLevelExternalReferenceOOXML(writer, "oleObject", externalOOXML)
        // activeX objects
        writeSheetLevelExternalReferenceOOXML(writer, "activeX", externalOOXML)
        // comments (notes)
        writeComments(writer, sheet, bk)

        writer!!.write("</worksheet>") // 20081028 KSC: Sheet xml should
        // be named SheetX.xml instead
        // of sheetname (see
        // writeWbOOXML as well)
        // finished, now write to <sheet>.xml + <sheet>.xml.rels, if has
        // associated content

        // write rels if necessary (printer settings, drawings ...)
        writeRels(shContentList, "xl/worksheets/_rels/sheet" + (id + 1) + ".xml.rels")
        sheetsContentList.addAll(shContentList) // since shContentList will be cleared out for each sheet, make sure sheet contents are stored
        // clear out content list for the sheet [associated documents for the
        // sheet, will be written to <SheetX>.xml.rels]
        shContentList.clear()
    }


    /**
     * Write the merged cell records for a worksheet
     */
    @Throws(IOException::class)
    protected fun writeMergedCellRecords(sheet: WorkSheetHandle) {
        val mcs = sheet.mysheet!!.mergedCells // TODO: PROBLEM:
        // getMergedCellRecs
        // contains null values,
        // screwing up output
        // ... fix!!
        // Use getMergedCells method - which DOESN'T add new blank merged cell
        if (mcs != null && mcs.size > 0) {
            val mc = StringBuffer()
            var cnt = 0
            for (i in mcs.indices) {
                val cr = (mcs[i] as Mergedcells).mergedRanges
                if (cr != null) {
                    for (j in cr.indices) {
                        val rng = cr[j].range
                        if (rng != null) {
                            val z = rng.indexOf("!") // strip sheetname
                            mc.append("<mergeCell ref=\""
                                    + cr[j].range!!.substring(z + 1) + "\"/>")
                            mc.append("\r\n")
                            cnt++
                        }
                    }
                }
            }
            if (mc.length > 0) { // Only input element if actual valid merged
                // cell ranges exist,
                writer!!.write("<mergeCells count=\"$cnt\">")
                writer!!.write("\r\n")
                writer!!.write(mc.toString())
                writer!!.write("</mergeCells>")
                writer!!.write("\r\n")
            }
        }

    }


    /**
     * Writes a row and all contents to the zip output
     *
     * @param hyperlinks
     * @throws IOException
     * @throws FormulaNotFoundException
     */
    @Throws(IOException::class)
    fun writeRow(row: RowHandle, hyperlinks: ArrayList<*>) {

        // row = sheet.getRow(x);
        // TODO: need spans?
        // <row element> -- eventually will be refactored as Object
        var h = ""
        if (row.height != 255)
        // if it's not default
            h = " ht=\"" + row.height / OOXMLAdapter.rowHtFactor + "\" customHeight=\"1\""
        writer!!.write("<row r=\"" + (row.rowNumber + 1) + "\"" + h)
        if (row.formatId > 0 && row.formatId > row.workBook!!.workBook.defaultIxfe)
        // row-level formatting specified
            writer!!.write(" s=\"" + row.formatId + "\" customFormat=\"1\"")
        if (row.hasAnyThickTopBorder)
            writer!!.write(" thickTop=\"1\"")
        if (row.hasAnyBottomBorder)
            writer!!.write(" thickBot=\"1\"")

        if (row.isHidden)
            writer!!.write(" hidden=\"1\"")
        if (row.isCollapsed)
            writer!!.write(" collapsed=\"1\"")                // 20090513 KSC: Added collapsed, outlineLevel [BUGTRACKER 2371]
        if (row.outlineLevel != 0)
            writer!!.write(" outlineLevel=\"" + row.outlineLevel + "\"")
        writer!!.write(">")
        writer!!.write("\r\n")
        // Cell element <c
        val ch = row.cells
        // iterate cells and output xml
        for (j in ch.indices) {
            val styleId = ch[j].cell!!.ixfe
            val dataType = ch[j].cellType
            if (ch[j].hasHyperlink()) {   // save; hyperlinks go after sheetData
                hyperlinks.add(arrayOf(ch[j].cellAddress, ch[j].url, ch[j].urlDescription))
            }
            writer!!.write("<c r=\"" + ch[j].cellAddress + "\"")
            if (styleId > 0)
                writer!!.write(" s=\"$styleId\"")
            when (dataType) {
                CellHandle.TYPE_STRING -> {
                    val s = ch[j].stringVal
                    var isErrVal = false
                    if (s!!.indexOf("#") == 0) {   // 20090521 KSC: must test if it's an error string value
                        isErrVal = Collections.binarySearch(Arrays.asList("#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"), s.trim { it <= ' ' }) > -1
                    }
                    if (!isErrVal) {
                        writer!!.write(" t=\"s\">")
                        // 20090520 KSC: can't use only string value as intra-cell formatting is possible zip.write("<v>" + ssts.indexOf(stripNonAscii(ch[j].getStringVal())) + "</v>");
                        val v = (ch[j].cell as Labelsst).isst // use isst instead of a lookup -- MUCHMUCHMUCH faster!
                        writer!!.write("<v>$v</v>")
                    } else {// it's an error value, must have type of "e"
                        writer!!.write(" t=\"e\">")
                        writer!!.write("<v>$s</v>")
                    }
                }
                CellHandle.TYPE_DOUBLE, CellHandle.TYPE_FP, CellHandle.TYPE_INT -> writer!!.write(" t=\"n\"><v>" + ch[j].`val` + "</v>")
                CellHandle.TYPE_FORMULA -> {
                    val fh: FormulaHandle
                    try {
                        fh = ch[j].formulaHandle
                        writer!!.write(fh.ooxml)
                    } catch (e: FormulaNotFoundException) {
                        Logger.logErr("Error getting formula handle in OOXML Writer")
                    }

                }
                CellHandle.TYPE_BOOLEAN -> writer!!.write(" t=\"b\"><v>" + ch[j].intVal + "</v>")
                CellHandle.TYPE_BLANK -> writer!!.write(">")
            }
            writer!!.write("</c>")
            writer!!.write("\r\n")
        }
        writer!!.write("</row>")
        writer!!.write("\r\n")
    }

    /**
     * retrieves the column ooxml possible attributes: bestFit, collapsed,
     * customWidth, hidden, max, min, style, width
     *
     * @param bk
     * @param sheet
     * @return
     */
    private fun getColOOXML(bk: WorkBookHandle, sheet: WorkSheetHandle): StringBuffer {
        val colooxml = StringBuffer()
        // ColHandle cols[]= sheet.getColumns();
        val iter = sheet.mysheet!!.colinfos.iterator()
        if (iter.hasNext()) {
            colooxml.append("<cols>")
            colooxml.append("\r\n")
            while (iter.hasNext()) {
                try {
                    // col width = width/256???
                    val c = iter.next()
                    val collast = c.colLast + 1
                    val w = c.colWidth / OOXMLAdapter.colWFactor
                    colooxml.append("<col min=\"" + (c.colFirst + 1)
                            + "\" max=\"" + collast + "\" width=\""
                            + c.colWidth / OOXMLAdapter.colWFactor
                            + "\" customWidth=\"1\"")
                    if (c.isHidden)
                        colooxml.append(" hidden=\"1\"")
                    if (c.ixfe > 0)
                    // column-level formatting specified
                        colooxml.append(" style=\"" + c.ixfe + "\"")

                    colooxml.append("/>")
                    colooxml.append("\r\n")
                } catch (e: Exception) {
                }

            }
            colooxml.append("</cols>")
            colooxml.append("\r\n")
        }
        return colooxml
    }

    /**
     * write all necessary information describing ImageHandle im
     * including .rels and drawing xml
     *
     * @param im
     */
    private fun getImageOOXML(im: ImageHandle): String {
        var ret = ""
        try {
            ret = im.getOOXML(drContentList.size + 1)   // obtain image OOXML; if errors don't write out to zip or contents
            // write image bytes to file
            var ext = im.type                   // 20090117 KSC: may come back as undefined, usually because it's an EMF and we can't process right now ...
            if (ext == "undefined") ext = "emf"
            val imageName = "image" + ++imgId + "." + ext
            // TODO: HANDLE REUSING IMAGES!!!!!!!!!!!!!!!!!!!!!
            addDeferredFile(im.imageBytes, OOXMLConstants.mediaDir + "/" + imageName)

            // trap package contents for drawing.xml
            drContentList.add(arrayOf("/" + OOXMLConstants.mediaDir + "/" + imageName, "image"))
        } catch (e: Exception) {
            Logger.logErr("OOXMLWriter.getImageOOXML: $e")
            ret = ""
        }

        return ret
    }

    /**
     * write out the drawing objects - images and charts + drawingml
     *
     * @param out
     * @param sheet
     * @param bk
     * @throws IOException
     */
    @Throws(IOException::class)
    private fun writeDrawingObjects(out: Writer, sheet: WorkSheetHandle, bk: WorkBookHandle) {
        // Drawing Objects  - Images, Charts & Shapes (OOXML-specific) = drawing, legacyDrawing, legacyDrawingHF, picture, oleObjects, controls
        val drawing = StringBuffer()
        val imz = sheet.images
        if (imz!!.size > 0) {
            // For each image, create a Drawing reference in sheet xml + write imageOOXML to drawingX.xml
            for (i in imz.indices) {
                // obtain image OOXML + write image file to ZIP
                drawing.append(getImageOOXML(imz[i]))
                drawing.append("\r\n")
            }
        }

        val charts = sheet.mysheet!!.charts
        var nUserShapes = 0 // if charts link to any drawing ml user shapes, see below for handling
        if (charts.size > 0) {
            // for each chart, create a chart.xml + trap references for drawingX.xml.rels
            for (i in charts.indices) {
                try {   // obtain image OOXML + write image file to ZIP
                    val c = charts[i] as Chart
                    drawing.append(getChartDrawingOOXML(ChartHandle(c, bk)))
                    drawing.append("\r\n")
                    if (c is OOXMLChart) {
                        val chartEmbeds = (charts[i] as OOXMLChart).chartEmbeds
                        if (chartEmbeds != null) {
                            val origDrawingId = drawingId            // id for THIS CURRENT DRAWING ML describing this chart(s), etc.
                            val chContentList = ArrayList()
                            for (j in chartEmbeds.indices) {
                                // obtain external drawingml file(s) which define shape and write to zip
                                val embed = chartEmbeds[j] as Array<String>
                                if (embed[0] == "userShape") {
                                    drawingId += nUserShapes + 1    // id for USER SHAPES drawingml
                                    val f = embed[1]
                                    nUserShapes++    // keep track of increment
                                    writeExOOXMLFile(arrayOf(embed[0], OOXMLConstants.drawingDir + "/", f), chContentList)
                                } else if (embed[0] == "image") {
                                    val f = embed[1]
                                    writeExOOXMLFile(arrayOf(embed[0], OOXMLConstants.mediaDir + "/", f), chContentList)
                                } else if (embed[0] == "themeOverride") {
                                    val f = embed[1]
                                    writeExOOXMLFile(arrayOf(embed[0], OOXMLConstants.themeDir + "/", f), chContentList)
                                }
                            }
                            writeRels(chContentList, OOXMLConstants.chartDir + "/_rels/chart" + chartId + ".xml.rels")
                            sheetsContentList.addAll(chContentList) // ensure embeds are written to main [Content_Types].xml
                            drawingId = origDrawingId            // reset to id for THIS CURRENT DRAWING ML
                        }
                    }
                } catch (e: Exception/*ChartNotFoundException c*/) {
                    Logger.logErr("OOXMLWriter.writeDrawingObjects failed getting Chart: $e")
                }

            }
        }

        // OOXML shapes OTHER THAN vmldrawings, which are handled elsewhere (in writelegacyDrawingObjects)
        if (sheet.mysheet!!.ooxmlShapes != null) {
            val shapes = sheet.mysheet!!.ooxmlShapes
            val i = shapes!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next() as String
                if (key == "vml") continue // handled in writelegacyDrawingObjects
                val o = shapes[key]
                if (o is TwoCellAnchor) {
                    drawing.append(o.ooxml)
                    if (o.embedFilename != null) {    // shape has embedded images we must also store
                        val f = o.embedFilename
                        val rId = writeExOOXMLFile(arrayOf("image", OOXMLConstants.mediaDir + "/", externalDir + f!!), drContentList)
                        o.embed = "rId$rId"
                    }
                } else if (o is OneCellAnchor) {
                    drawing.append(o.ooxml)
                    // TODO: trap embedded images as in twoCellAnchor
                    //drContentList.add(new String[] { "/" + mediaDir + "/"+ imageName, "image"});
                }
                drawing.append("\r\n")
            }
        }

        //
        if (drawing.length > 0) {   // then have drawing objects to write out
            // one drawingX.xml per sheet, write reference in sheet
            out.write("<drawing r:id=\"rId" + (shContentList.size + 1) + "\"/>")
            out.write("\r\n")    // link drawing.xml to specific image
            // write out drawingml
            val drawingml = StringBuffer()
            drawingml.append(OOXMLConstants.xmlHeader)
            drawingml.append("\r\n")
            drawingml.append("<xdr:wsDr xmlns:xdr=\"" + OOXMLConstants.drawingns + "\" xmlns:a=\"" + OOXMLConstants.drawingmlns + "\">")
            drawingml.append("\r\n")
            drawingml.append(drawing)
            drawingml.append("</xdr:wsDr>")
            drawingml.append("\r\n")
            // write drawingX.xml to Zip

            // FIX
            addDeferredFile(drawingml, OOXMLConstants.drawingDir + "/drawing" + ++drawingId + ".xml")

            shContentList.add(arrayOf("/" + OOXMLConstants.drawingDir + "/drawing" + drawingId + ".xml", "drawing"))

            // write drawingX.xml.rels to Zip
            writeRels(drContentList, OOXMLConstants.drawingDir + "/_rels/drawing" + drawingId + ".xml.rels")

            sheetsContentList.addAll(drContentList)
            drContentList.clear()
            drawingId += nUserShapes  // if wrote drawingml for usershapes, must increment next drawingId to AFTER these files
        }
    }


    /**
     * write all necessary information describing ChartHandle ch
     * including .rels and drawing xml
     *
     * @param ch
     * @return
     */
    private fun getChartDrawingOOXML(ch: ChartHandle): String {
        var ret = ""
        try {
            // trap package contents for drawing.xml
            val chartml = StringBuffer()
            chartml.append(OOXMLConstants.xmlHeader)
            chartml.append("\r\n")
            chartml.append("<c:chartSpace xmlns:c=\"" + OOXMLConstants.chartns + "\" xmlns:a=\"" + OOXMLConstants.drawingmlns + "\" xmlns:r=\"" + OOXMLConstants.relns + "\">")
            chartml.append("\r\n")
            chartml.append(ch.getOOXML(drContentList.size + 1))
            chartml.append("</c:chartSpace>")
            chartml.append("\r\n")
            ret = ch.getChartDrawingOOXML(drContentList.size + 1)   // if errors out don't write/add to zip
            addDeferredFile(chartml, OOXMLConstants.chartDir + "/chart" + ++chartId + ".xml")
            drContentList.add(arrayOf("/" + OOXMLConstants.chartDir + "/chart" + chartId + ".xml", "chart"))
        } catch (e: Exception) {
            Logger.logErr("er: $e")
            ret = ""
        }

        return ret
    }


    /**
     * write notes or comments for the specific sheet
     *
     * @param zip
     * @param sheet
     * @param bk
     * @throws IOException
     */
    @Throws(IOException::class)
    private fun writeComments(out: Writer, sheet: WorkSheetHandle,
                              bk: WorkBookHandle) {
        val nh = sheet.commentHandles
        if (nh == null || nh.size == 0)
            return
        val comments = StringBuffer()
        comments.append(OOXMLConstants.xmlHeader)
        comments.append("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n<authors>")
        // run thru 1x to get authors
        val authors = ArrayList()
        for (i in nh.indices) {
            if (!authors.contains(nh[i].author)) {
                comments.append("\r\n<author>" + OOXMLAdapter.stripNonAscii(nh[i].author) + "</author>")
                authors.add(nh[i].author)
            }
        }
        comments.append("\r\n</authors>\r\n<commentList>")
        for (i in nh.indices) {
            comments.append("\r\n" + nh[i].getOOXML(authors.indexOf(nh[i].author)))
        }
        comments.append("\r\n</commentList>\r\n</comments>")
        addDeferredFile(comments, "xl/comments" + ++commentsId + ".xml")

        shContentList.add(arrayOf("/xl/comments$commentsId.xml", "comments"))

    }

    /**
     * writes the legacy drawing xml in drawing directory/vmlDrawingX.vml plus the rels for any embedded objects
     * <br></br>most of the vml is saved from the original (if any); the vml used to define notes is genearated here
     *
     * @param zip
     * @param sheet
     * @param bk
     * @return true if there is vml (legacy drawing info and/or notes) for this sheet
     * @throws IOException
     */
    @Throws(IOException::class)
    private fun writeLegacyDrawingObjects(out: Writer, sheet: WorkSheetHandle,
                                          bk: WorkBookHandle): Boolean {
        // TODO: handle embeds + rels -- correct???? ****************************
        // TODO: data=1 --> is that number of shapes??  <o:idmap v:ext="edit" data=""/>
        // number of shape id's??????
        var embeds: Array<String>? = null
        var vml: StringBuffer? = null
        if (sheet.mysheet!!.ooxmlShapes != null) {
            if (sheet.mysheet!!.ooxmlShapes!!["vml"] is Array<Any>) { // original vml contains embeds
                val o = sheet.mysheet!!.ooxmlShapes!!["vml"] as Array<Any>
                vml = o[0]
                embeds = o[1] as Array<String>
            } else
                vml = sheet.mysheet!!.ooxmlShapes!!["vml"] as StringBuffer
        }
        // add Notes to the vml, if any
        val nh = sheet.commentHandles
        if (nh != null && nh.size > 0) {
            if (vml == null) {
                vml = StringBuffer()
                // add apparently required shapelayout element
                vml.append("<o:shapelayout v:ext=\"edit\">" +
                        "<o:idmap v:ext=\"edit\" data=\"1\"/>" + /*        --> data="1" --> number of elements in the vml?????*/
                        "</o:shapelayout>")
            }
            // add shapetype which defines textbox (==202)
            vml.append("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">" +
                    "<v:stroke joinstyle=\"miter\"/>" +
                    "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>" +
                    "</v:shapetype>")


            for (i in nh.indices) {
                val hidden = nh[i].isHidden
                val row = nh[i].rowNum
                val col = nh[i].colNum
                val bounds = nh[i].textBoxBounds
                val spid = nh[i].internalNoteRec!!.spid
                vml.append("<v:shape id=\"_x0000_s" + spid + "\"" +  /* id of text box = id of mso */
                        " type=\"#_x0000_t202\"" +                           /* type of text box */
                        " style=\"position:absolute;" +
                        "margin-left:203.25pt;" +
                        "margin-top:37.5pt;" +
                        "width:96pt;" +
                        "height:55.5pt;" +
                        "z-index:1;" +
                        "visibility:" + (if (hidden) "hidden" else "visible") + "\"" +  /* shown or hidden */
                        " fillcolor=\"#ffffe1\"" +
                        " o:insetmode=\"auto\">")
                vml.append("<v:fill color2=\"#ffffe1\"/>" +                /* general textbox characteristics */
                        " <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>" +
                        " <v:path o:connecttype=\"none\"/>" +
                        "<v:textbox style=\"mso-direction-alt:auto\">" +
                        "<div style=\"text-align:left\"/>" +
                        " </v:textbox>")
                vml.append("<x:ClientData ObjectType=\"Note\">" +      /* note object */
                        "<x:MoveWithCells/>" +
                        "<x:SizeWithCells/>")
                vml.append("<x:Anchor>")                              /* bounds of text box */
                for (j in bounds!!.indices) {
                    vml.append(bounds[j].toString() + if (j < bounds.size - 1) "," else "")
                }
                vml.append("</x:Anchor>")

                vml.append("<x:AutoFill>False</x:AutoFill>" +          /* row/col where note is attached to */
                        "<x:Row>" + row + "</x:Row>" +
                        "<x:Column>" + col + "</x:Column>")
                if (!hidden)
                    vml.append("<x:Visible/>")
                vml.append("</x:ClientData>" + "</v:shape>")
            }
        }
        if (vml != null && vml.length > 0) {
            // start with ns info
            vml.insert(0, "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"" +
                    " xmlns:o=\"urn:schemas-microsoft-com:office:office\"" +
                    " xmlns:x=\"urn:schemas-microsoft-com:office:excel\">")
            vml.append("</xml>")
            addDeferredFile(vml, OOXMLConstants.drawingDir + "/vmlDrawing" + ++vmlId + ".vml")

            shContentList.add(arrayOf("/" + OOXMLConstants.drawingDir + "/vmlDrawing" + vmlId + ".vml", "vmldrawing"))

            if (embeds != null) {
                val vmlContentList = ArrayList()
                for (i in embeds!!.indices) {
                    var pp = embeds!![i].trim({ it <= ' ' })   // file on disk or saved filename
                    val typ = pp.substring(0, pp.indexOf("/"))
                    pp = pp.substring(pp.indexOf("/") + 1)
                    val z = pp.lastIndexOf("/") + 1
                    val pth = pp.substring(0, z)
                    pp = pp.substring(z)
                    val ff = getExOOXMLFileName(pp) // desired or outut filename
                    deferredFiles[pth + ff] = externalDir + pp
                    vmlContentList.add(arrayOf("/$pth$ff", typ))
                }
                writeRels(vmlContentList, OOXMLConstants.drawingDir + "/_rels/vmlDrawing" + vmlId + ".vml.rels")
                // NOTE: vmlfiles do not get listed in main [Content_Types].xml
            }
            return true
        }
        return false
    }

}
