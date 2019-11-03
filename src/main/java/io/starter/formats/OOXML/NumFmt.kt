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
package io.starter.formats.OOXML

import io.starter.formats.XLS.FormatConstants
import io.starter.formats.XLS.OOXMLAdapter
import io.starter.formats.XLS.Xf
import io.starter.toolkit.Logger
import io.starter.toolkit.StringTool
import org.xmlpull.v1.XmlPullParser

/**
 * numFmt  OOXML element
 *
 *
 * numFmt (Number Format)
 * This element specifies number format properties which indicate how to format and render the numeric value of
 * a cell.
 * Following is a listing of number formats whose formatCode value is implied rather than explicitly saved in
 * the file. In this case a numFmtId value is written on the xf record, but no corresponding numFmt element is
 * written. Some of these Ids are interpreted differently, depending on the UI language of the implementing
 * application.
 *
 *
 * parent:  styleSheet/numFmts element in styles.xml
 * NOTE:  numFmt element also occurs in drawingML, numFmt is replaced with sourceLinked
 * children: none
 */
class NumFmt : OOXMLElement {
    private var formatCode: String? = null
    /**
     * returns the format id assoc with this number format
     *
     * @return
     */
    var formatId: String? = null
        private set
    private var sourceLinked = false

    override val ooxml: String
        get() = getOOXML("")

    constructor(formatCode: String, numFmtId: String, sourceLinked: Boolean) {
        this.formatCode = formatCode
        this.formatId = numFmtId
        this.sourceLinked = sourceLinked
    }

    constructor(n: NumFmt) {
        this.formatCode = n.formatCode
        this.formatId = n.formatId
        this.sourceLinked = n.sourceLinked
    }

    fun getOOXML(ns: String): String {
        val ooxml = StringBuffer()
        ooxml.append("<" + ns + "numFmt")
        // attributes
        ooxml.append(" formatCode=\"" + OOXMLAdapter.stripNonAscii(formatCode) + "\"")
        if (formatId != null) ooxml.append(" numFmtId=\"$formatId\"")
        if (sourceLinked) ooxml.append(" sourceLinked=\"1\"")
        ooxml.append("/>")
        return ooxml.toString()
    }

    override fun cloneElement(): OOXMLElement {
        return NumFmt(this)
    }

    companion object {

        private val serialVersionUID = -206715418106414662L

        fun parseOOXML(xpp: XmlPullParser): OOXMLElement {
            var formatCode: String? = null
            var numFmtId: String? = null
            var sourceLinked = false
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "numFmt") {
                            for (i in 0 until xpp.attributeCount) {
                                val n = xpp.getAttributeName(i)
                                val v = xpp.getAttributeValue(i)
                                if (n == "formatCode") {
                                    formatCode = v
                                } else if (n == "numFmtId")
                                    numFmtId = v
                                else if (n == "sourceLinked")
                                    sourceLinked = v == "1"
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "numFmt") {
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("numFmt.parseOOXML: $e")
            }

            return NumFmt(formatCode, numFmtId, sourceLinked)
        }


        /**
         * returns the OOXML specifying the fill based on this FormatHandle object
         */
        fun getOOXML(xf: Xf): String {
            // Number Format
            val ooxml = StringBuffer()
            if (xf.ifmt > FormatConstants.BUILTIN_FORMATS.size) {    // only input user defined formats ...
                var s = xf.formatPattern
                if (s != null)
                    s = StringTool.replaceText(s, "\"", "&quot;") // replace internal quotes   // 1.6 only s= s.replace('"', "&quot;"); // replace internal quotes
                ooxml.append("<numFmt numFmtId=\"" + xf.ifmt + "\" formatCode=\"" + s + "\"/>")
                ooxml.append("\r\n")
            }
            return ooxml.toString()
        }
    }
}

