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

import java.util.Stack

/**
 * style
 *
 *
 * This element specifies the style information for a shape. This is used to define a shape's appearance in terms of
 * the preset styles defined by the style matrix for the theme.
 *
 *
 * parent:  sp, pic, cnxSp
 * children (required and in sequence):  lnRef, fillRef, effectRef, fontRef
 */
class Style : OOXMLElement {
    private var effectRef: EffectRef? = null
    private var fontRef: FontRef? = null
    private var fillRef: FillRef? = null
    private var lRef: lnRef? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:style>")
            ooxml.append(lRef!!.ooxml)
            ooxml.append(fillRef!!.ooxml)
            ooxml.append(effectRef!!.ooxml)
            ooxml.append(fontRef!!.ooxml)
            ooxml.append("</xdr:style>")
            return ooxml.toString()
        }

    constructor(lr: lnRef, flr: FillRef, er: EffectRef, fr: FontRef) {
        this.lRef = lr
        this.fillRef = flr
        this.effectRef = er
        this.fontRef = fr
    }

    constructor(s: Style) {
        this.lRef = s.lRef
        this.fillRef = s.fillRef
        this.effectRef = s.effectRef
        this.fontRef = s.fontRef
    }

    override fun toString(): String {
        return ooxml
    }

    override fun cloneElement(): OOXMLElement {
        return Style(this)
    }

    companion object {

        private val serialVersionUID = -583023685473342509L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var er: EffectRef? = null
            var fr: FontRef? = null
            var flr: FillRef? = null
            var lr: lnRef? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "effectRef") {
                            lastTag.push(tnm)
                            er = EffectRef.parseOOXML(xpp, lastTag, bk)
                        } else if (tnm == "fontRef") {
                            lastTag.push(tnm)
                            fr = FontRef.parseOOXML(xpp, lastTag, bk)
                        } else if (tnm == "fillRef") {
                            lastTag.push(tnm)
                            flr = FillRef.parseOOXML(xpp, lastTag, bk)
                        } else if (tnm == "lnRef") {
                            lastTag.push(tnm)
                            lr = lnRef.parseOOXML(xpp, lastTag, bk)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "style") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("style.parseOOXML: $e")
            }

            return Style(lr, flr, er, fr)
        }
    }

}

/**
 * effectRef (Effect Reference)
 * This element defines a reference to an effect style within the style matrix. The idx attribute refers the index of
 * an effect style within the effectStyleLst element.
 *
 *
 * parent: many
 * children:  COLORCHOICE
 */
internal class EffectRef : OOXMLElement {
    private var idx: Int = 0
    private var colorChoice: ColorChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:effectRef idx=\"$idx\">")
            if (colorChoice != null) ooxml.append(colorChoice!!.ooxml)
            ooxml.append("</a:effectRef>")
            return ooxml.toString()
        }

    protected constructor(idx: Int, c: ColorChoice) {
        this.idx = idx
        this.colorChoice = c
    }

    protected constructor(er: EffectRef) {
        this.colorChoice = er.colorChoice
        this.idx = er.idx
    }

    override fun cloneElement(): OOXMLElement {
        return EffectRef(this)
    }

    companion object {
        private val serialVersionUID = -7572271663955122478L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): EffectRef {
            var idx = -1
            var c: ColorChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "effectRef") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "idx") {
                                    idx = Integer.valueOf(xpp.getAttributeValue(i)).toInt()
                                }
                            }
                        } else if (tnm == "schemeClr" ||
                                tnm == "hslClr" ||
                                tnm == "prstClr" ||
                                tnm == "scrgbClr" ||
                                tnm == "srgbClr" ||
                                tnm == "sysClr") {
                            lastTag.push(tnm)
                            c = ColorChoice.parseOOXML(xpp, lastTag, bk) as ColorChoice
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "effectRef") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("effectRef.parseOOXML: $e")
            }

            return EffectRef(idx, c)
        }
    }
}

/**
 * fillRef (Fill Reference)
 * This element defines a reference to a fill style within the style matrix. The idx attribute refers to the index of a
 * fill style or background fill style within the presentation's style matrix, defined by the fmtScheme element. A
 * value of 0 or 1000 indicates no background, values 1-999 refer to the index of a fill style within the fillStyleLst
 * element, and values 1001 and above refer to the index of a background fill style within the bgFillStyleLst
 * element. The value 1001 corresponds to the first background fill style, 1002 to the second background fill style,
 * and so on.
 *
 *
 * parent:  many
 * children: COLORCHOICE
 */
internal class FillRef : OOXMLElement {
    private var idx: Int = 0
    private var colorChoice: ColorChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:fillRef idx=\"$idx\">")
            if (colorChoice != null) ooxml.append(colorChoice!!.ooxml)
            ooxml.append("</a:fillRef>")
            return ooxml.toString()
        }

    protected constructor(idx: Int, c: ColorChoice) {
        this.idx = idx
        this.colorChoice = c
    }

    protected constructor(fr: FillRef) {
        this.colorChoice = fr.colorChoice
        this.idx = fr.idx
    }

    override fun cloneElement(): OOXMLElement {
        return FillRef(this)
    }

    companion object {
        private val serialVersionUID = 7691131082710785068L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): FillRef {
            var idx = -1
            var c: ColorChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "fillRef") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "idx") {
                                    idx = Integer.valueOf(xpp.getAttributeValue(i)).toInt()
                                }
                            }
                        } else if (tnm == "schemeClr" ||
                                tnm == "hslClr" ||
                                tnm == "prstClr" ||
                                tnm == "scrgbClr" ||
                                tnm == "srgbClr" ||
                                tnm == "sysClr") {
                            lastTag.push(tnm)
                            c = ColorChoice.parseOOXML(xpp, lastTag, bk).cloneElement() as ColorChoice
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "fillRef") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("fillRef.parseOOXML: $e")
            }

            return FillRef(idx, c)
        }
    }
}

/**
 * fontRef (Font Reference)
 * This element represents a reference to a themed font. When used it specifies which themed font to use along
 * with a choice of color.
 *
 *
 * parent: many
 * children:  COLORCHOICE
 */
internal class FontRef : OOXMLElement {
    private var idx: String? = null
    private var colorChoice: ColorChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:fontRef idx=\"$idx\">")
            if (colorChoice != null) ooxml.append(colorChoice!!.ooxml)
            ooxml.append("</a:fontRef>")
            return ooxml.toString()
        }

    protected constructor(idx: String, c: ColorChoice) {
        this.idx = idx
        this.colorChoice = c
    }

    protected constructor(fr: FontRef) {
        this.colorChoice = fr.colorChoice
        this.idx = fr.idx
    }

    override fun cloneElement(): OOXMLElement {
        return FontRef(this)
    }

    companion object {

        private val serialVersionUID = 2907761758443581273L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): FontRef {
            var idx: String? = null
            var c: ColorChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "fontRef") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "idx") {
                                    idx = xpp.getAttributeValue(i)
                                }
                            }
                        } else if (tnm == "schemeClr" ||
                                tnm == "hslClr" ||
                                tnm == "prstClr" ||
                                tnm == "scrgbClr" ||
                                tnm == "srgbClr" ||
                                tnm == "sysClr") {
                            lastTag.push(tnm)
                            c = ColorChoice.parseOOXML(xpp, lastTag, bk) as ColorChoice
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "fontRef") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("fontRef.parseOOXML: $e")
            }

            return FontRef(idx, c)
        }
    }
}

/**
 * lnRef (Line Reference)
 * This element defines a reference to a line style within the style matrix. The idx attribute refers the index of a
 * line style within the fillStyleLst element
 *
 *
 * parent: many
 * children: COLORCHOICE
 */
internal class lnRef : OOXMLElement {
    private var idx: Int = 0
    private var colorChoice: ColorChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<a:lnRef idx=\"$idx\">")
            if (colorChoice != null) ooxml.append(colorChoice!!.ooxml)
            ooxml.append("</a:lnRef>")
            return ooxml.toString()
        }

    protected constructor(idx: Int, c: ColorChoice) {
        this.idx = idx
        this.colorChoice = c
    }

    protected constructor(lr: lnRef) {
        this.colorChoice = lr.colorChoice
        this.idx = lr.idx
    }

    override fun cloneElement(): OOXMLElement {
        return lnRef(this)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -4349076266006929729L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): lnRef {
            var idx = -1
            var c: ColorChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "lnRef") {
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "idx") {
                                    idx = Integer.valueOf(xpp.getAttributeValue(i)).toInt()
                                }
                            }
                        } else if (tnm == "schemeClr" ||
                                tnm == "hslClr" ||
                                tnm == "prstClr" ||
                                tnm == "scrgbClr" ||
                                tnm == "srgbClr" ||
                                tnm == "sysClr") {
                            lastTag.push(tnm)
                            c = ColorChoice.parseOOXML(xpp, lastTag, bk) as ColorChoice
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "lnRef") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("lnRef.parseOOXML: $e")
            }

            return lnRef(idx, c)
        }
    }
}

