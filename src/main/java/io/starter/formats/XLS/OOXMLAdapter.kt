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

import io.starter.OpenXLS.ChartHandle
import io.starter.OpenXLS.Document
import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.OOXML.OOXMLConstants
import io.starter.formats.XML.UnicodeInputStream
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import io.starter.toolkit.TempFileManager
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserException
import org.xmlpull.v1.XmlPullParserFactory

import java.io.*
import java.util.*
import java.util.zip.ZipEntry
import java.util.zip.ZipFile
import java.util.zip.ZipOutputStream

// TODO: finish Style options, Chart options, Japanese Chars ...,
// TODO: dataTables
// TODO: external names
// TODO: drawing shapes
// TODO: Pivot Tables, Macro Sheets ...
// TODO: handle generation of: printer settings, doc properties, external refs ... on xls --> xlsx

/* ******************************************************************************************************
 * Dissertation on one of the main complications in the below code:
 *
 * "External Objects" are XMLs and associated files that are external to the main workbook and worksheets,
 * and essentially are treated as pass-throughs (except for charts, sharedString table + styles, which are generated)
 *
 * Because these are not created upon output, pass-through external objects (xml files, bins, etc) are written to a directory upon input,
 * and pulled from that directory upon output.  Data associated with the external object is stored in either WorkBook.getOOMXLObjects or BoundSheet.getOOXMLObjects.
 * Linked data associated with the pass-through object is stored and .rels are re-created
 *
 * Here is a list of the External Objects:
 *
 * (Workbook Level)
 * vbaProject
 * externalLink**
 * calcChain
 * sharedStrings
 * styles
 * theme**
 *
 * (Sheet Level)
 * printerSettings
 * drawing**
 * control (=activeX)
 * oleObject
 *
 * **external items may also have linked items to them e.g. themes may have object associated with them
 *
 * All associated objects are listed in the associated .rels file
 */
open class OOXMLAdapter : OOXMLConstants {
    protected var DEBUG = !true

    internal var zip: ZipOutputStream? = null
    internal var writer: Writer? = null
    internal var deferredFiles: MutableMap<*, *> = HashMap()
    // map of original external filename and new filename on disk (may be different
    // as images, etc. must have consecutive indexes ...
    // HashMap externalFiles= new HashMap(); not used anylonger

    // content lists used to create all .rels + [Content_Types].xml
    internal var mainContentList = ArrayList() // [Content_Types].xml
    internal var wbContentList = ArrayList() // workbook.xml.rels - written to [Content_Types].xml
    internal var drContentList = ArrayList() // drawingX.xml.rels - written to [Content_Types].xml
    internal var shContentList = ArrayList() // sheetX.xml.rels
    internal var sheetsContentList = ArrayList() // total contents of each sheet - written to [Content_Types].xml
    // also have vmContentList and chContentList for fairly rare occurrences of
    // vmldrawings and charts containing embeds

    // External OOXML Object such as Vba projects, Ole Objects, Printer Settings,
    // etc.
    // links external ref "extra info" with the external reference id
    // Map shExternalLinkInfo= new HashMap();

    internal var externalDir = "" // store "pass-through" files i.e. files we cannot process into our BIFF8 rec
    // structure (vbaProject.bin, for example)
    // ordinal numbers for sheet-level objects (rid is stored in sheetX.xml and file
    // stored in appropriate directory, liked via sheetX.xml.rels
    // Each item (images, charts, etc) has a very specific and ordered name format
    // e.g. image1.jpg, printerSettings2.bin ...
    internal var drawingId = 0
    internal var vmlId = 0 // vmlDrawing.vml
    internal var commentsId = 0
    internal var activeXId = 0
    internal var activeXBinaryId = 0 // activeX.bin
    internal var printerSettingsId = 0
    internal var oleObjectsId = 0
    internal var chartId = 0
    internal var imgId = 0

    // TODO: finish Styles OOXML --> cellStyleXfs, MANY options not handled ...
    // TODO: finish charts and images (many options not handled)
    // TODO: handle tableData
    // TODO: handle themes, doc properties (create? alter?)
    // TODO: handle shapes

    // ***************************************************************************************
    // contents of External OOXMLObject arraylist
    internal var EX_TYPE = 0 // type of External Object - must be listed in OOXMLConstants
    internal var EX_PATH = 1 // path in ZIP
    internal var EX_FNAME = 2 // file name
    // 3= rid
    internal var EX_EXTRAINFO = 4 // any extra information associated - object specific
    internal var EX_EMBEDINFO = 5 // string of embedd
    // ***************************************************************************************
    /**
     * set the XLSX format for this WorkBook <br></br>
     * either FORMAT_XLSX, FORMAT_XLSM (Macro-enabled), FORMAT_XLTS (template) or
     * FORMAT_XLTM (Macro-enabled template) <br></br>
     * NOTE: If file extension is .XLSM format FORMAT_XLSM must be set <br></br>
     * either because there are macros present or because the filename <br></br>
     * is unconditionally set to .XLSM
     *
     * @param format
     */
    var format = WorkBookHandle.FORMAT_XLSX // default format is non-macro-enabled workbook

    /**
     * generic utility to take the original external object filename and parse it
     * using a new ordinal id e.g. image10.emf may become image3.emf NOTE:
     * Increments global ordinal ids
     *
     * @param f Original filename with path
     * @return New filename
     */
    protected fun getExOOXMLFileName(f: String): String {
        var fname = f.substring(f.lastIndexOf("/") + 1)
        val ext = fname.substring(fname.lastIndexOf("."))
        var root = fname.substring(0, fname.indexOf("."))
        var z = root.length - 1 // now skip # appended to end of root (number
        // will be regenerated from id)
        while (Character.isDigit(root.get(z)))
            z--
        root = root.substring(0, z + 1)
        if (root.equals("image", ignoreCase = true)) {
            fname = root + ++imgId + ext
        } else if (root.equals("oleObject", ignoreCase = true)) {
            fname = root + ++oleObjectsId + ext
        } else if (root.equals("activeX", ignoreCase = true)) {
            if (ext.toLowerCase() == ".xml")
                fname = root + ++activeXId + ext
            else if (ext.toLowerCase() == ".bin")
                fname = root + ++activeXBinaryId + ext
        } else if (root.equals("printerSettings", ignoreCase = true)) {
            fname = root + ++printerSettingsId + ext
        } else if (root.equals("drawing", ignoreCase = true)) {
            fname = root + ++drawingId + ext
            // if workbook-level file, no incrementing id, just use original
            // filename
        } else if (root.equals("app", ignoreCase = true) || root.equals("core", ignoreCase = true) || root.equals("theme", ignoreCase = true)
                || root.equals("themeOverride", ignoreCase = true) || root.equals("app", ignoreCase = true)
                || root.equals("custom", ignoreCase = true) || root.equals("connections", ignoreCase = true)
                || root.equals("externalLink", ignoreCase = true) || root.equals("calcChain", ignoreCase = true)
                || root.equals("styles", ignoreCase = true) || root.equals("sharedStrings", ignoreCase = true)
                || root.equals("vbaProject", ignoreCase = true)) {
            // do nothing, use original fname *** these do not have ordinal id's
        } else
        // TESTING - remove when done!
        // *********************************************************************
            Logger.logErr("Unknown External Type: $root")
        return fname
    }

    /**
     * write a temp file for later inclusion in the zip file
     *
     * @param sb
     * @param fn
     */
    protected fun addDeferredFile(sb: StringBuffer, fn: String) {
        try {
            val fx = addDeferredFile(fn)
            val fos = FileOutputStream(fx)
            /*
             * new way OOXMLAdapter.writeSBToStreamEfficiently(sb,fos); fos.flush();
             * fos.close();
             */ // have to use writer to set encoding - vital to non-utf8 input files
            val out = OutputStreamWriter(fos, inputEncoding) // "UTF-8" );
            OOXMLAdapter.writeSBToStreamEfficiently(sb, out)
            out.close()
            /*
             * old way Writer out = new OutputStreamWriter( fos, inputEncoding); //"UTF-8"
             * ); out.write( sb.toString() ); out.close();
             */
        } catch (e: Exception) {
            Logger.logErr("OOXMLAdapter addDeferredFile failed.", e)
        }

    }

    @Throws(IOException::class)
    protected fun addDeferredFile(name: String): File {
        val file = TempFileManager.createTempFile("OOXMLOutput_", ".tmp")
        deferredFiles[name] = file.absolutePath
        return file
    }

    /**
     * write a temp file for later inclusion in the zip file
     *
     * @param sb
     * @param fn
     * @throws IOException
     */
    @Throws(IOException::class)
    protected fun addDeferredFile(b: ByteArray, fn: String) {
        val fx = addDeferredFile(fn)
        // write to temp file
        val fos = FileOutputStream(fx)
        val bos = BufferedOutputStream(fos)
        bos.write(b)
        bos.close()
    }

    /**
     * returns truth of "Book Contains External OOXML Object named type" *
     *
     * @param bk
     * @param type e.g. "vba", "custprops"
     * @return
     */
    private fun hasObject(bk: WorkBookHandle, type: String): Boolean {
        val externalOOXML = bk.workBook!!.ooxmlObjects
        for (i in externalOOXML.indices) {
            val s = externalOOXML[i] as Array<String>
            if (s != null && s!!.size == 3) { // id, dir, filename
                if (s!![0].equals(type, ignoreCase = true))
                    return true
            }
        }
        return false
    }

    /**
     * utility which, given an an exisitng file f2write, creates zipEntry and writes
     * to zip var
     *
     *
     * NOTE: global zip ZipOutputStream must be open
     *
     * @param f2write name and path of exisitng file
     * @param fname   desired name in zip
     * @throws IOException
     */
    @Throws(IOException::class)
    protected fun writeFileToZIP(f2write: String, fname: String) {
        nextZipEntry(fname)

        val fis = FileInputStream(f2write)
        val bis = BufferedInputStream(fis)

        var i = bis.read()
        while (i != -1) {
            zip!!.write(i)
            i = bis.read()
        }

        bis.close()
    }

    @Throws(IOException::class)
    protected fun nextZipEntry(name: String) {
        // Flush the writer to ensure data ends up in the right entry
        try {
            writer!!.flush()
        } catch (e: Exception) {
            Logger.logErr("Flush failing on zip entry, likely due to streaming first sheet $e")
        }

        // Start the new entry in the ZIP file
        zip!!.putNextEntry(ZipEntry(name))
    }

    /**
     * utility to look up Content Type string for type abbreviation
     *
     * @param type
     * @return
     */
    protected fun getContentType(type: String): String {
        for (i in OOXMLConstants.contentTypes.indices) {
            if (OOXMLConstants.contentTypes[i][0].equals(type, ignoreCase = true))
                return OOXMLConstants.contentTypes[i][1]
        }
        return "UNKNOWN TYPE $type"
    }

    /**
     * utility to look up Relationship Type string for type abbreviation
     *
     * @param type
     * @return
     * @see OOXMLConstnats
     */
    protected fun getRelationshipType(type: String): String {
        for (i in OOXMLConstants.relsContentTypes.indices) {
            if (OOXMLConstants.relsContentTypes[i][0].equals(type, ignoreCase = true))
                return OOXMLConstants.relsContentTypes[i][1]
        }
        return "UNKNOWN TYPE $type"
    }

    // this is for testing purposes only, not used so no need to comment out logger
    // msg :)
    private fun getZipEntries(zf: ZipFile) {
        // testing!!
        try {
            val e = zf.entries()
            while (e.hasMoreElements()) {
                val ze = e.nextElement() as ZipEntry
                Logger.logInfo(ze.name)
            }
        } catch (e: Exception) {
            Logger.logErr("getZipEntries: $e")
        }

    }

    /**
     * simple utility that ensures that sheets are last in the workbookcontent list
     * also ensure that theme(s) are parsed first, as are used in styles etc. in
     * order to create all dependent objects first
     *
     * @param wbContentList
     */
    protected fun reorderWbContentList(wbContentList: ArrayList<*>) {
        for (j in wbContentList.indices) {
            val wb = wbContentList[j] as Array<String>
            if (wb[0] != "sheet") { // sheets come last
                wbContentList.removeAt(j)
                wbContentList.add(0, wb)
            }
        }
        for (j in wbContentList.indices) {
            val wb = wbContentList[j] as Array<String>
            if (wb[0] == "theme") { // goes before styles
                wbContentList.removeAt(j)
                wbContentList.add(0, wb)
                break
            }
        }

    }

    /**
     * creates a copy of a stack so changes won't affect original
     *
     * @param origStack
     * @return
     */
    protected fun cloneStack(origStack: Stack<*>): Stack<*> {
        val s = Stack()
        for (i in origStack.indices) {
            s.push(origStack.elementAt(i))
        }
        return s
    }

    companion object {

        internal val rowHtFactor = 20.0
        internal val colWFactor = 256.0
        internal var inputEncoding = "UTF-8" // default

        /**
         * Parses an xsd:boolean value.
         *
         * @param value the string to parse
         * @return the boolean value of the given string
         * @throws IllegalArgumentException if the given string is not a valid boolean value
         */
        fun parseBoolean(value: String): Boolean {
            val trimmed = value.trim { it <= ' ' }
            if (trimmed == "true" || trimmed == "1")
                return true
            if (trimmed == "false" || trimmed == "0")
                return false
            throw IllegalArgumentException("'$value' is not a valid boolean value")
        }

        /**
         * get a standalone ChartML document
         *
         * @param ch
         * @return
         */
        fun getStandaloneChartDrawingOOXML(ch: ChartHandle): String {
            var ret = ""
            try {
                // trap package contents for drawing.xml
                val chartml = StringBuffer()
                chartml.append(OOXMLConstants.xmlHeader)
                chartml.append("\r\n")
                chartml.append("<c:chartSpace xmlns:c=\"" + OOXMLConstants.chartns + "\" xmlns:a=\"" + OOXMLConstants.drawingmlns + "\" xmlns:r=\""
                        + OOXMLConstants.relns + "\">")
                chartml.append("\r\n")
                chartml.append(ch.getOOXML(1))
                chartml.append("</c:chartSpace>")
                chartml.append("\r\n")
                return chartml.toString()
            } catch (e: Exception) {
                Logger.logErr("OOXMLAdapter.getStandaloneChartDrawingOOXML: $e")
                ret = ""
            }

            return ret
        }

        /**
         * A memory efficient way to write a StringBuffer to an OutputStream without
         * creating Strings and other Objects.
         *
         *
         * Mar 9, 2012
         *
         * @param aSB
         * @param ous
         * @throws IOException
         */
        // public static void writeSBToStreamEfficiently(StringBuffer aSB, OutputStream
        // ous) throws IOException{
        @Throws(IOException::class)
        fun writeSBToStreamEfficiently(aSB: StringBuffer, ous: Writer) {

            val aLength = aSB.length
            val aChunk = 1024
            val aChars = CharArray(aChunk) // aPosEnd-aPosStart]; //aChunk];

            var aPosStart = 0
            while (aPosStart < aLength) {
                val aPosEnd = Math.min(aPosStart + aChunk, aLength)
                aSB.getChars(aPosStart, aPosEnd, aChars, 0) // Create no new buffer
                val aCARead = CharArrayReader(aChars) // Create no new buffer

                // This may be slow but it will not create any more buffer (for bytes)
                var aByte: Int
                var i = 0
                while ((aByte = aCARead.read()) != -1 && i++ < aPosEnd - aPosStart)
                    ous.write(aByte)
                aPosStart += aChunk
            }
        }

        /**
         * return true of workbook bk contains macros
         *
         * @param bk
         * @return
         */
        fun hasMacros(bk: Document): Boolean {
            if (bk is WorkBookHandle) {
                val externalOOXML = bk.workBook!!.ooxmlObjects
                for (i in externalOOXML.indices) {
                    val s = externalOOXML[i] as Array<String>
                    if (s != null && s!!.size == 3) { // id, dir, filename
                        if (s!![0].equals("vba", ignoreCase = true) || s!![0].equals("macro", ignoreCase = true)) {
                            return true
                        }
                    }
                }
            }
            return false
        }

        /**
         * utility to retrieve correct relationship type abbreviation string from
         * verbose type string
         *
         * @param type
         * @return
         * @see OOXMLConstants
         */
        protected fun getRelationshipTypeAbbrev(type: String): String {
            for (i in OOXMLConstants.relsContentTypes.indices) {
                if (OOXMLConstants.relsContentTypes[i][1].equals(type, ignoreCase = true))
                    return OOXMLConstants.relsContentTypes[i][0]
            }
            return "UNKNOWN TYPE $type"
        }

        /**
         * Strip non-ascii (i.e. xml non-valid) chars from Strings This is utilized for
         * xml attributes, node values can contain quote symbols
         *
         * @param s
         * @return
         */
        fun stripNonAscii(s: String?): StringBuffer {
            val out = StringBuffer()
            if (s == null)
                return out
            /**
             * FROM MS: "Special character" refers to any character outside the standard
             * ASCII character set range of 0x00 - 0x7F, such as Latin characters with
             * accents, umlauts, or other diacritics. The default encoding scheme for XML
             * documents is UTF-8, which encodes ASCII characters with a value of 0x80 or
             * higher differently than other standard encoding schemes. Most often, you see
             * this problem if you are working with data that uses the simple "iso-8859-1"
             * encoding scheme. In this case, the quickest solution is usually the first or
             * example, use the following XML declaration:
             */
            // Legal characters are tab, carriage return, line feed, and the legal
            // characters of Unicode and ISO/IEC 10646
            // XML processors MUST accept any character in the range specified for Char
            // Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] |
            // [#x10000-#x10FFFF] /* any Unicode character, excluding the surrogate blocks,
            // FFFE, and FFFF. */

            /*
         * fro wiki Unicode code points in the following ranges are valid in XML 1.0
         * documents:[1]
         *
         * U+0009, U+000A, U+000D: these are the only C0 controls accepted in XML 1.0;
         * U+0020–U+D7FF, U+E000–U+FFFD: this excludes some (not all) non-characters in
         * the BMP (all surrogates, U+FFFE and U+FFFF are forbidden); U+10000–U+10FFFF:
         * this includes all code points in supplementary planes, including
         * non-characters.
         *
         * The preceding code points ranges contain the following controls which are
         * only valid in certain contexts in XML 1.0 documents, and whose usage is
         * restricted and highly discouraged: U+007F–U+0084, U+0086–U+009F: this
         * includes a C0 control characters, and most (not all) C1 controls.
         */
            for (i in 0 until s.length) {
                val c = s[i]
                val charCode = c.toInt()
                if (charCode == 0x9 || charCode == 0xA || charCode == 0xD || charCode >= 0x20 && charCode <= 0xD7FF
                        || charCode >= 0xE000 && charCode <= 0xFFFD || charCode >= 0x10000 && charCode <= 0x10FFFF) {
                    if (charCode == '&'.toInt())
                        out.append("&amp;")
                    else if (charCode == '"'.toInt())
                        out.append("&quot;")
                    else if (charCode == '<'.toInt())
                        out.append("&lt;")
                    else if (charCode == '>'.toInt())
                        out.append("&gt;")
                    else if (charCode == '\''.toInt())
                        out.append("&apos;")
                    else
                        out.append(c)
                } /*
             * else { // Encoding Q??????:
             * io.starter.toolkit.Logger.log("Skipping Special Char: " + charCode); //
             * skip it }
             */
            }
            /*
         * these translations do not seem to make any difference for Baxter's issue }
         * else if (charCode==8220) { // smart quotes BAXTER ISSUE - TRY
         * out.append("&#8220;"); } else if (charCode==8221) { // ""
         * out.append("&#8221;"); } else if (charCode==8216) { // ""
         * out.append("&#8216;"); } else if (charCode==8217) { // ""
         * out.append("&#8217;"); } else if (charCode==8211){ // en dash
         * out.append("&#8211;"); } else if (charCode==8212){ // em dash
         * out.append("&#8212;"); } else if (charCode==8242) { // single prime
         * out.append("&#8242;"); } else if (charCode==8364) { // Euro Symbol
         * out.append("&#8364;"); } else // copyright= 169 // 10 \n 176=small circle
         * out.append(c); // 169=&copy; }
         */
            return out
        }

        /**
         * Strip non-ascii (i.e. xml non-valid) chars from Strings Node values can
         * contain quote symbols
         *
         * @param s
         * @return
         */
        fun stripNonAsciiRetainQuote(s: String?): StringBuffer {
            val out = StringBuffer()
            if (s == null)
                return out
            for (i in 0 until s.length) {
                val c = s[i]
                val charCode = c.toInt()
                if (charCode >= 32 && charCode <= 126) {
                    if (charCode == '&'.toInt())
                        out.append("&amp;")
                    else if (charCode == '<'.toInt())
                        out.append("&lt;")
                    else if (charCode == '>'.toInt())
                        out.append("&gt;")
                    else
                        out.append(c)
                } else
                    out.append(c)
            }
            return out
        }

        /**
         * deal with the "BOM" input streams...
         *
         *
         * http://bugs.sun.com/bugdatabase/view_bug.do?bug_id=4508058
         *
         *
         *
         *
         * Oct 14, 2010
         *
         * @param in
         * @return
         */
        fun wrapInputStream(`in`: InputStream): InputStream {
// String enc = uin.getEncoding(); // check for BOM mark and skip
            return UnicodeInputStream(`in`, "UTF-8")
        }

        /**
         * used as a way to monitor Zip Entry fetching
         *
         *
         * Oct 14, 2010
         *
         * @param f
         * @param name
         * @return
         */
        fun getEntry(f: ZipFile, name: String): ZipEntry {
            return f.getEntry(name)
        }

        /**
         * parses any .rels file into content List array list
         *
         * @param ii
         * @return
         */
        fun parseRels(ii: InputStream): ArrayList<*> {
            val contentList = ArrayList()
            try {
                val factory = XmlPullParserFactory.newInstance()
                factory.isNamespaceAware = true
                val xpp = factory.newPullParser()

                xpp.setInput(ii, null) // using XML 1.0 specification

                // if(DEBUG) Logger.logInfo("parseRels InputStream has available bytes: " +
                // ii.available());

                var eventType = xpp.eventType

                // if(DEBUG) Logger.logInfo("parseRels XPP Name: " + xpp.getName() );

                // if(DEBUG) Logger.logInfo("parseRels XPP Event Type: " + eventType );

                while (eventType != XmlPullParser.END_DOCUMENT) {
                    try {
                        if (eventType == XmlPullParser.START_TAG) {
                            val tnm = xpp.name
                            if (tnm != null && tnm == "Relationship") {
                                var type = ""
                                var target = ""
                                var rId = ""
                                for (i in 0 until xpp.attributeCount) {
                                    val nm = xpp.getAttributeName(i) // id, Type,
                                    // Target
                                    val v = xpp.getAttributeValue(i)
                                    if (nm.equals("Type", ignoreCase = true))
                                        type = getRelationshipTypeAbbrev(v)
                                    else if (nm.equals("Target", ignoreCase = true))
                                        target = v
                                    else if (nm.equals("id", ignoreCase = true))
                                        rId = v
                                }
                                // 20100426 KSC: unfortunately, need to ensure that commentsX.xml is processed
                                // before vmlDrawingX.xml
                                if (target.indexOf("comments") == -1)
                                    contentList.add(arrayOf(type, target, rId))
                                else
                                // ensure comments are before vmlDrawing so that Notes can be created
                                    contentList.add(0, arrayOf(type, target, rId))
                            } else {
                                // if(DEBUG) Logger.logInfo("parseRels null entry name");
                            }
                        }
                    } catch (ea: Exception) {
                        Logger.logErr("XML Exception in OOXMLAdapter.parseRels. Input file is out of spec.", ea)
                    }

                    eventType = xpp.next()
                }

            } catch (ex: org.xmlpull.v1.XmlPullParserException) {
                Logger.logErr("XML Exception in OOXMLAdapter.parseRels. Input file is out of spec.", ex)

            } catch (e: Exception) {
                Logger.logErr("OOXMLAdapter.parseRels. $e")
            }

            return contentList
        }

        /**
         * utility to retrieve Text element for tag
         *
         * @param xpp
         * @return
         * @throws IOException
         */
        @Throws(IOException::class, XmlPullParserException::class)
        fun getNextText(xpp: XmlPullParser): String? {
            var eventType = xpp.next()
            var ret = ""
            while (eventType != XmlPullParser.END_DOCUMENT && eventType != XmlPullParser.END_TAG
                    && eventType != XmlPullParser.START_TAG && /* true in all cases?? */
                    eventType != XmlPullParser.TEXT) {
                eventType = xpp.next()
            }
            if (eventType == XmlPullParser.TEXT)
                ret = xpp.text

            try {
                return String(ret.toByteArray(), inputEncoding)
                /*
             * KSC: replaced with above *SHOULD* be correct if (!isUTFEncoding) if
             * (xpp.getInputEncoding().equals("UTF-8")) return new String(ret.getBytes(),
             * "UTF-8"); // ensure encoding else return new String(ret.getBytes(),
             * xpp.getInputEncoding()); // ensure encoding
             */
            } catch (e: Exception) {
            }
            // inputEncoding can be null
            return ret
        }

        /**
         * return the file matching rId in the ContentList (String[] type, filename,
         * rId)
         *
         * @param contentList
         * @param rId
         * @return
         */
        fun getFilename(contentList: ArrayList<*>, rId: String): String? {
            for (j in contentList.indices) {
                val s = contentList[j] as Array<String>
                if (s[2] == rId)
                    return s[1]
            }
            return null
        }

        fun deleteDir(f: File): Boolean {
            if (f.isDirectory) {
                val children = f.list()
                for (i in children!!.indices) {
                    val success = deleteDir(File(f, children[i]))
                    if (!success) {
                        // return false;
                    }
                }
            }
            // The directory is now empty so delete it
            f.deleteOnExit()
            return f.delete()
        }

        @Throws(IOException::class)
        fun getTempDir(f: String): String {
            var f = f
            /*
         * File fx = TempFileManager.createTempFile("OOXMLA",".tmp"); File fdir =
         * fx.getParentFile(); String s = ""; if(fdir.isDirectory()) s =
         * fdir.getAbsolutePath(); else{
         */
            var s = System.getProperty("java.io.tmpdir")

            if (!(s.endsWith("/") || s.endsWith("\\")))
                s += "/"
            f = StringTool.stripPath(f)
            s += "extentech/"
            if (f.indexOf('.') > 0)
                s += f.substring(0, f.indexOf('.')) + "/"
            else
                s += "$f/"

            return s
        }

        /**
         * sorts the sheets for incoming workbook xlsx -- used for eventMode only
         *
         *
         * Jan 19, 2011
         *
         * @param cl
         */
        @Throws(Exception::class)
        protected fun sortSheets(cl: ArrayList<*>) {
            // take the array of storages and find sheets

            // use natural sort of the Tree
            val sorted = TreeMap()

            var its: Iterator<*> = cl.iterator()

            while (its.hasNext()) {
                val c = its.next() as Array<String>
                var shtnm = c[1]
                // parse out sheet number
                // xxxxsheet1.xmlxxxx
                var st = shtnm.indexOf("worksheets/sheet")
                if (st > -1) {
                    st += 16
                    shtnm = shtnm.substring(st)
                    shtnm = shtnm.substring(0, shtnm.toLowerCase().indexOf(".xml"))
                    try {
                        val ti = Integer.parseInt(shtnm)
                        // we know the sheet number, add to the tree
                        sorted.put(Integer.valueOf(ti), c)
                    } catch (e: Exception) {
                        Logger.logErr("Could not sort sheets", e)
                        return
                    }

                }
            }
            // now we have the sorted map of "CLs" we can re-order sheets in arraylist
            its = cl.iterator()
            val sort = sorted.values.iterator()
            var clpos = -1
            while (sort.hasNext()) {
                val c = sort.next() as Array<String>
                var found = false
                // now find in cl
                while (its.hasNext() && !found) {
                    clpos++
                    val cx = its.next() as Array<String>
                    if (cx[0] == "sheet") { // replace the sheet entries one by one
                        cl[clpos] = c
                        found = true
                    }
                }
            }
        }

        /**
         * re-save to temp directroy any "pass-through" OOXML files i.e. files/entities
         * not present in 2003-version and thus not processed
         *
         * <br></br>
         * (app.xml, theme.xml ...)
         *
         * @param wbh
         */
        fun refreshPassThroughFiles(wbh: WorkBookHandle) {
            try {
                // retrieve source zip
                val sourceZip = java.util.zip.ZipFile(wbh.file!!)
                OOXMLReader.refreshExternalFiles(sourceZip,
                        OOXMLAdapter.getTempDir(wbh.workBook!!.factory!!.fileName))
                wbh.file = null
            } catch (e: Exception) { // wbh.getFile() can be an XLS file (as source) so Exception is almost always OK
                // (do not report)
                // Logger.logErr("OOXMLAdapter.refreshPassThroughFiles: could not retrieve
                // source ooxml: " + e.toString());
            }

        }
    }

}

/**
 * helper class for array compares ...
 */
internal class intArray {
    private var a: IntArray? = null

    val isZero: Boolean
        get() {
            if (a == null)
                return true
            for (i in a!!.indices)
                if (a!![i] != 0)
                    return false
            return true
        }

    constructor(a: IntArray) {
        this.a = IntArray(a.size)
        for (i in a.indices) {
            this.a[i] = a[i]
        }
    }

    constructor(a: ShortArray) {
        this.a = IntArray(a.size)
        for (i in a.indices) {
            this.a[i] = a[i].toInt()
        }
    }

    fun get(): IntArray? {
        return a
    }

    override fun equals(o: Any?): Boolean {
        val testa = (o as intArray).get()
        if (testa == null || a == null || testa.size != this.a!!.size)
            return false
        for (i in this.a!!.indices) {
            if (this.a!![i] != testa[i])
                return false
        }
        return true
    }
}

internal class objArray(a: Array<Any>) {
    private val a: Array<Any>? = null

    init {
        this.a = a
        /*
         * this.a= new Object[a.length]; for (int i= 0; i < a.length; i++) { this.a[i]=
         * a[i]; }
         */
    }

    fun get(): Array<Any>? {
        return a
    }

    override fun equals(o: Any?): Boolean {
        val testa = (o as objArray).get()
        return java.util.Arrays.equals(this.a, testa)
    }

}
