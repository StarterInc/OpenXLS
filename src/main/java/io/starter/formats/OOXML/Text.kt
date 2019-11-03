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

import io.starter.OpenXLS.WorkBookHandle
import io.starter.formats.XLS.Font
import io.starter.formats.XLS.OOXMLAdapter
import io.starter.formats.XLS.Sst
import io.starter.formats.XLS.Unicodestring
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.ArrayList
import java.util.Stack

/**
 * text (Comment Text)
 * This element contains rich text which represents the text of a comment. The maximum length for this text is a
 * spreadsheet application implementation detail. A recommended guideline is 32767 chars
 * parent: comment
 * children:	t (text), r (Rich Text Run), rPh (phonetic run), phoneticPr (phonetic properties)
 */
// TODO: finish elements rPh and phoneticPr
// TODO: preserve
class Text : OOXMLElement {
    /**
     * return the String value of this Text (Comment) element
     * Including formatting runs
     *
     * @return
     */
    var commentWithFormatting: Unicodestring? = null
        private set

    override val ooxml: String?
        get() = null

    /**
     * return the String value of this Text (Comment) element
     * i.e. without formatting
     *
     * @return
     */
    val comment: String?
        get() = if (commentWithFormatting != null) commentWithFormatting!!.stringVal else null

    /**
     * create a new comment WITH formatting
     *
     * @param strref
     */
    constructor(str: Unicodestring) {
        this.commentWithFormatting = str
    }

    constructor(t: Text) {
        this.commentWithFormatting = t.commentWithFormatting
    }

    /**
     * create a new comment with NO formatting
     *
     * @param s
     */
    constructor(s: String) {
        this.commentWithFormatting = Sst.createUnicodeString(s, null, Sst.STRING_ENCODING_AUTO)
    }

    /**
     * return the OOXML representation of this Text (Comment) element
     *
     * @param bk
     * @return
     */
    fun getOOXML(bk: io.starter.formats.XLS.WorkBook): String {
        val ooxml = StringBuffer()
        ooxml.append("<text>")
        if (commentWithFormatting != null) {
            var s = OOXMLAdapter.stripNonAsciiRetainQuote(commentWithFormatting!!.stringVal).toString()
            val runs = commentWithFormatting!!.formattingRuns
            if (runs == null) {    // no intra-string formatting
                if (s.indexOf(" ") == 0 || s.lastIndexOf(" ") == s.length - 1) {
                    ooxml.append("<t xml:space=\"preserve\">$s</t>")
                } else {
                    ooxml.append("<t>$s</t>")
                }
                ooxml.append("\r\n")
            } else {    // have formatting runs which split up string into areas with separate formats applied
                /*
				 *  <element name="rPr" type="CT_RPrElt" minOccurs="0" maxOccurs="1"/>
					<element name="t" type="ST_Xstring" minOccurs="1" maxOccurs="1"/>
				 */
                var begIdx = 0
                ooxml.append("<r>")    // new rich text run
                for (j in runs.indices) {
                    val idxs = runs[j] as ShortArray
                    if (idxs[0] > begIdx) {
                        ooxml.append("<t xml:space=\"preserve\">" + OOXMLAdapter.stripNonAscii(s.substring(begIdx, idxs[0].toInt())) + "</t>")
                        ooxml.append("</r>")
                        ooxml.append("\r\n")
                        ooxml.append("<r>")
                        begIdx = idxs[0].toInt()
                    }
                    val rp = Ss_rPr.createFromFont(bk.getFont(idxs[1].toInt()))
                    ooxml.append(rp.ooxml)
                }
                if (begIdx < s.length)
                // output remaining string
                    s = s.substring(begIdx)
                else
                    s = ""
                ooxml.append("<t xml:space=\"preserve\">" + OOXMLAdapter.stripNonAscii(s) + "</t>")
                ooxml.append("\r\n")
                ooxml.append("</r>")
            }
        }
        ooxml.append("</text>")
        ooxml.append("\r\n")
        return ooxml.toString()
    }

    override fun cloneElement(): OOXMLElement {
        return Text(this)
    }

    companion object {

        private val serialVersionUID = 5886384020139606328L

        /**
         * parse this Text element into a unicode string with formatting runs
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): Text {
            var str: Unicodestring? = null
            var s = ""
            var formattingRuns: ArrayList<ShortArray>? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "rPr") {    // intra-string formatting properties
                            val idx = s.length    // index into character string to apply formatting to
                            val rp = Ss_rPr.parseOOXML(xpp, bk) as Ss_rPr    //.cloneElement();
                            val f = rp.generateFont(bk)  // NOW CONVERT ss_rPr to a font!!
                            var fIndex = bk.workBook!!.getFontIdx(f)  // index for specific font formatting
                            if (fIndex == -1)
                            // must insert new font
                                fIndex = bk.workBook!!.insertFont(f) + 1
                            if (formattingRuns == null) formattingRuns = ArrayList()
                            formattingRuns.add(shortArrayOf(Integer.valueOf(idx)!!.toShort(), Integer.valueOf(fIndex)!!.toShort()))
                        } else if (tnm == "t") {
                            /*boolean bPreserve= false;
		            	 if (xpp.getAttributeCount()>0) {
		            		 if (xpp.getAttributeName(0).equals("space") && xpp.getAttributeValue(0).equals("preserve"))
		            			 bPreserve= true;
		            	 }
		            	 */
                            eventType = xpp.next()
                            while (eventType != XmlPullParser.END_DOCUMENT &&
                                    eventType != XmlPullParser.END_TAG &&
                                    eventType != XmlPullParser.TEXT) {
                                eventType = xpp.next()
                            }
                            if (eventType == XmlPullParser.TEXT) {
                                s = s + xpp.text
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG && xpp.name == "text") {
                        str = Sst.createUnicodeString(s, formattingRuns, Sst.STRING_ENCODING_UNICODE)    // create a new unicode string with formatting runs
                        break
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("r.parseOOXML: $e")
            }

            return Text(str)
        }
    }
}

