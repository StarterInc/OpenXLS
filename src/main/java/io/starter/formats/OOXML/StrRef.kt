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

import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Stack

/**
 * OOXML element strRef, string reference, child of tx (chart text) element or cat (category) element
 */
class StrRef(f: String, s: StrCache) : OOXMLElement {
    private val stringRef: String? = null
    private val strCache: StrCache? = null

    /**
     * generate ooxml to define a strRef, part of tx element or cat element
     *
     * @return
     */
    /**
     * strRef contains f + strRef elements
     */
    override val ooxml: String
        get() {
            val tooxml = StringBuffer()
            tooxml.append("<c:strRef>")
            if (this.stringRef != null) tooxml.append("<c:f>" + this.stringRef + "</c:f>")
            if (this.strCache != null) tooxml.append(strCache.ooxml)
            tooxml.append("</c:strRef>")
            return tooxml.toString()
        }

    init {
        this.stringRef = f
        this.strCache = s
    }

    override fun cloneElement(): OOXMLElement {
        return StrRef(this.stringRef, this.strCache)
    }

    companion object {

        private val serialVersionUID = -5992001371281543027L

        /**
         * parse strRef OOXML element
         *
         * @param xpp     XmlPullParser
         * @param lastTag element stack
         * @return spRef object
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): OOXMLElement {
            var f: String? = null
            var s: StrCache? = null

            /**
             * contains (in Sequence)
             * f
             * strRef
             */
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "f") {
                            f = io.starter.formats.XLS.OOXMLAdapter.getNextText(xpp)
                        } else if (tnm == "strCache") {
                            lastTag.push(tnm)
                            s = StrCache.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "strRef") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("title.parseOOXML: $e")
            }

            return StrRef(f, s)
        }
    }

}


/**
 * define OOXML strCache element
 */
internal class StrCache(ptCount: Int, idx: Int, pt: String) : OOXMLElement {
    private val ptCount = -1
    private val idx = -1
    private val pt: String? = null

    /**
     * generate ooxml to define a strCache element, part of strRef element
     *
     * @return
     */
    override val ooxml: String
        get() {
            val tooxml = StringBuffer()
            tooxml.append("<c:strCache>")
            tooxml.append("<c:ptCount val=\"" + this.ptCount + "\"/>")
            tooxml.append("<c:pt idx=\"" + this.idx + "\">")
            tooxml.append("<c:v>" + this.pt + "</c:v>")
            tooxml.append("</c:pt>")
            tooxml.append("</c:strCache>")
            return tooxml.toString()
        }

    init {
        this.ptCount = ptCount
        this.idx = idx
        this.pt = pt
    }

    override fun cloneElement(): OOXMLElement {
        return StrCache(this.ptCount, this.idx, this.pt)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -4914374179641060956L

        /**
         * parse title OOXML element title
         *
         * @param xpp     XmlPullParser
         * @param lastTag element stack
         * @return spCache object
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<*>): StrCache {
            var ptCount = -1
            var idx = -1
            var pt: String? = null

            /**
             * contains (in Sequence)
             * ptCount
             * pt
             */
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "ptCount") {
                            ptCount = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                        } else if (tnm == "pt") {
                            idx = Integer.valueOf(xpp.getAttributeValue(0)).toInt()
                            pt = io.starter.formats.XLS.OOXMLAdapter.getNextText(xpp)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "strCache") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("strCache.parseOOXML: $e")
            }

            return StrCache(ptCount, idx, pt)
        }
    }

}
