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
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserFactory

import java.io.InputStream

/**
 * not fully implemented yet
 * Need to handle fonts, fills ...
 * Until then themes will not be re-created; instead, it will be treated as a pass-through
 */
class Theme : OOXMLElement {
    /**
     * Generic "Office" Color Scheme TODO: Read from theme1.xml
     * 12 colors: 2 darks, 2 lights, 6 accents, 2 hyperlinks
     *
     *
     * NOTE: APPEARS that 1st 4 are "swapped" so in theme1.xml
     * dk1 is index 0, then lt1, dk2, lt2
     * but appears to be a known bug (or a known mystery) that
     * one must swap the 1st two pairs
     */
    var genericThemeClrs = arrayOf("FFFFFF", // text/bg lt1	= window bg (white)
            "000000", // text/bg dk1	= window text (black)
            "EEECE1", // text/bg lt2	= secondary window bg color	  (grayish)
            "1F497D", // text/bg dk2	= secondary window text color (deep blue)
            "4F81BD", // accent1	-- med-deep blue
            "C0504D", // accent2	-- maroon
            "9BBB59", // accent3  -- lime greenish
            "8064A2", // accent4	-- purplish blue
            "4BACC6", // accent5	-- med blue
            "F79646", // accent6	-- orange-coral
            "0000FF", // hlink	-- blue
            "800080")// folhlink	-- dk purple	(followed hlink)

    override val ooxml: String?
        get() = null

    /**
     * given Theme OOXML inputstream, retrieve theme colors for later use
     * This element holds all the different formatting options available to a document through a theme and defines the overall look
     * and feel of the document when themed objects are used within the document.
     *
     * @param bk WorkBookHandle
     * @param ii InputStream
     * @see parseSheetOOXML
     */
    fun parseOOXML(bk: WorkBookHandle, ii: InputStream) {
        try {
            val factory = XmlPullParserFactory.newInstance()
            factory.isNamespaceAware = true
            val xpp = factory.newPullParser()
            xpp.setInput(ii, null) // using XML 1.0 specification
            var eventType = xpp.eventType
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    var idx = 0
                    var tnm = xpp.name
                    if (tnm == "clrScheme") {        // in both themeX.xml, themeOverrideX.xml.  12 colors defined.
                        eventType = xpp.next()
                        while (eventType != XmlPullParser.END_DOCUMENT) {
                            if (eventType == XmlPullParser.START_TAG) {
                                tnm = xpp.name
                                if (tnm == "dk1") {
                                    idx = 1
                                } else if (tnm == "lt1") {
                                    idx = 0
                                } else if (tnm == "dk2") {
                                    idx = 3
                                } else if (tnm == "lt2") {
                                    idx = 2
                                } else if (tnm == "accent1") {
                                    idx = 4
                                } else if (tnm == "accent2") {
                                    idx = 5
                                } else if (tnm == "accent3") {
                                    idx = 6
                                } else if (tnm == "accent4") {
                                    idx = 7
                                } else if (tnm == "accent5") {
                                    idx = 8
                                } else if (tnm == "accent6") {
                                    idx = 9
                                } else if (tnm == "hlink") {
                                    idx = 10
                                } else if (tnm == "folHlink") {
                                    idx = 11
                                } else if (tnm == "sysClr") { // system color attributes val, lastClr
                                    this.genericThemeClrs[idx] = xpp.getAttributeValue("", "lastClr")
                                } else if (tnm == "srgbClr") {
                                    this.genericThemeClrs[idx] = xpp.getAttributeValue(0)
                                }
                            } else if (eventType == XmlPullParser.END_TAG && xpp.name == "clrScheme") {
                                break
                            }
                            eventType = xpp.next()
                        }
                    } else if (tnm == "fmtScheme") {
                        // This element contains the background fill styles, effect styles, fill styles, and line styles which define the style matrix for a theme.
                        // The style matrix consists of subtle, moderate, and intense fills, lines, and effects.
                    } else if (tnm == "fontScheme") {
                        // majorFont		 defines the set of major fonts which are to be used under different languages or locals.
                        // minorFont		 defines the set of minor fonts that are to be used under different languages or locals
                    } else if (tnm == "objectDefaults") {        // only in themeX.xml

                    } else if (tnm == "extraClrSchemeList") {    // only in themeX.xml

                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                }
                eventType = xpp.next()
            }
        } catch (e: Exception) {
            Logger.logErr("Theme.parseOOXML: $e")
        }

        return
    }

    // TODO: implement
    override fun cloneElement(): OOXMLElement? {
        return null
    }

    companion object {

        private val serialVersionUID = -9201334460078323287L

        fun parseThemeOOXML(bk: WorkBookHandle, ii: InputStream): Theme {
            val t = Theme()
            t.parseOOXML(bk, ii)
            return t
        }
    }

}
