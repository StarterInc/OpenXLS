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
 * spPr
 * This OOXML/DrawingML element specifies the visual shape properties that can be applied to a shape. These properties include the
 * shape fill, outline, geometry, effects, and 3D orientation.
 *
 *
 * parents:  cxnSp, pic, sp
 * children:
 * blipFill (Picture Fill) §5.1.10.14
 * custGeom (Custom Geometry) §5.1.11.8
 * effectDag (Effect Container) §5.1.10.25
 * effectLst (Effect Container) §5.1.10.26
 * extLst (Extension List) §5.1.2.1.15
 * gradFill (Gradient Fill) §5.1.10.33
 * grpFill (Group Fill) §5.1.10.35
 * ln (Outline) §5.1.2.1.24
 * noFill (No Fill) §5.1.10.44
 * pattFill (Pattern Fill) §5.1.10.47
 * prstGeom (Preset geometry) §5.1.11.18
 * scene3d (3D Scene Properties) §5.1.4.1.26
 * solidFill (Solid Fill) §5.1.10.54
 * sp3d (Apply 3D shape properties) §5.1.7.12
 * xfrm (2D Transform)
 */
class SpPr : OOXMLElement {
    private var x: Xfrm? = null
    private var geom: GeomGroup? = null
    /**
     * returns the fill of this shape property set, if any
     *
     * @return
     */
    var fill: FillGroup? = null
        private set
    /**
     * returns the ln element of this shape property set, if any
     *
     * @return
     */
    var ln: Ln? = null
        private set
    private var effect: EffectPropsGroup? = null
    // scene3d, sp3d
    internal var bwMode: String? = null
    // namespace
    private var ns: String? = null

    override// must pass namespace to xfrm
    // geometry choice
    // fill choice
    // ln element
    // effect properties choice
    // scene3d, sp3d
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<" + this.ns + ":spPr")
            if (bwMode != null)
                ooxml.append(" bwMode=\"$bwMode\">")
            else
                ooxml.append(">")
            if (x != null) ooxml.append(x!!.ooxml)
            if (geom != null) ooxml.append(geom!!.ooxml)
            if (fill != null)
                ooxml.append(fill!!.ooxml)
            if (ln != null) ooxml.append(ln!!.ooxml)
            if (effect != null) ooxml.append(effect!!.ooxml)
            ooxml.append("</$ns:spPr>")
            return ooxml.toString()
        }

    /**
     * return the id for the embedded picture (i.e. resides within the file)
     *
     * @return
     */
    /**
     * set the embed attribute for this blip (the id for the embedded picture)
     *
     * @param embed
     */
    var embed: String?
        get() = if (fill != null) fill!!.embed else null
        set(embed) {
            if (fill != null) fill!!.embed = embed
        }

    /**
     * return the id for the linked picture (i.e. doesn't reside in the file)
     *
     * @return
     */
    /**
     * set the link attribute for this blip (the id for the linked picture)
     *
     * @param embed
     */
    var link: String?
        get() = if (fill != null) fill!!.link else null
        set(link) {
            if (fill != null) fill!!.link = link
        }

    /**
     * returns the width of the line of this shape proeprty,
     * or -1 if no line is present
     *
     * @return
     */
    val lineWidth: Int
        get() = if (ln != null) ln!!.width else -1

    /**
     * returns the color of the line of this shape property, or -1 if no line is present
     *
     * @return
     */
    val lineColor: Int
        get() = if (ln != null) ln!!.color else -1

    /**
     * returns the line style of the line of this shape property, or -1 if no line is present
     *
     * @return
     */
    val lineStyle: Int
        get() = if (ln != null) ln!!.lineStyle else -1

    /**
     * return the fill color
     *
     * @return
     */
    val color: Int
        get() = if (fill != null) fill!!.color else 1

    constructor(ns: String) {    // no-param constructor, set up common defaults
        x = Xfrm()
        x!!.setNS("a")
        this.ns = ns
        geom = GeomGroup(PrstGeom(), null)
        bwMode = "auto"
    }

    constructor(x: Xfrm, geom: GeomGroup, fill: FillGroup, l: Ln, effect: EffectPropsGroup, bwMode: String, ns: String) {
        this.x = x
        this.geom = geom
        this.fill = fill
        this.ln = l
        this.effect = effect
        this.bwMode = bwMode
        this.ns = ns
    }

    constructor(clone: SpPr) {
        this.x = clone.x
        this.geom = clone.geom
        this.fill = clone.fill
        this.ln = clone.ln
        this.effect = clone.effect
        this.bwMode = clone.bwMode
        this.ns = clone.ns
    }

    /**
     * create a default shape property with a solid fill and a line
     *
     * @param solidfill
     * @param w
     * @param lnClr
     */
    constructor(ns: String, solidfill: String?, w: Int, lnClr: String) {
        this.ns = ns
        if (solidfill != null)
            this.fill = FillGroup(null, null, null, null, SolidFill(solidfill))
        this.ln = Ln(w, lnClr)
    }

    /**
     * set the namespace for spPr element
     *
     * @param ns
     */
    fun setNS(ns: String) {
        this.ns = ns
    }

    override fun cloneElement(): OOXMLElement {
        return SpPr(this)
    }

    /**
     * add a line for this shape property
     *
     * @param w   line width
     * @param clr html color string
     */
    fun setLine(w: Int, clr: String) {
        ln = Ln()
        ln!!.width = w
        ln!!.setColor(clr)
    }

    /**
     * remove the line, if any, for this shape property
     */
    fun removeLine() {
        ln = null
    }

    /**
     * return true if this shape properties contains a line
     */
    fun hasLine(): Boolean {
        return ln != null
    }

    companion object {

        private val serialVersionUID = 4542844402486023785L

        /**
         * parse shape OOXML element spPr
         *
         * @param xpp     XmlPullParser
         * @param lastTag element stack
         * @return spPr object
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var x: Xfrm? = null
            var fill: FillGroup? = null
            var l: Ln? = null
            var effect: EffectPropsGroup? = null
            var geom: GeomGroup? = null
            // scene3d, sp3d
            var bwMode: String? = null
            var ns: String? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "spPr") {
                            ns = xpp.prefix
                            for (i in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(i)
                                if (nm == "bwMode")
                                    bwMode = xpp.getAttributeValue(i)
                            }
                        } else if (tnm == "xfrm") {
                            lastTag.push(tnm)
                            x = Xfrm.parseOOXML(xpp, lastTag) as Xfrm
                            //x.setNS("a");
                        } else if (tnm == "ln") {
                            lastTag.push(tnm)
                            l = Ln.parseOOXML(xpp, lastTag, bk) as Ln
                        } else if (tnm == "prstGeom" || tnm == "custGeom") {        // GEOMETRY GROUP
                            lastTag.push(tnm)
                            geom = GeomGroup.parseOOXML(xpp, lastTag) as GeomGroup
                        } else if (// FILL GROUP
                                tnm == "solidFill" ||
                                tnm == "noFill" ||
                                tnm == "gradFill" ||
                                tnm == "grpFill" ||
                                tnm == "pattFill" ||
                                tnm == "blipFill") {
                            lastTag.push(tnm)
                            fill = FillGroup.parseOOXML(xpp, lastTag, bk) as FillGroup
                        } else if (tnm == "extLst") {
                            lastTag.push(tnm)
                            ExtLst.parseOOXML(xpp, lastTag) // ignore for now TODO: FINISH
                        } else if (// EFFECT GROUP
                                tnm == "effectLst" || tnm == "effectDag") {
                            lastTag.push(tnm)
                            effect = EffectPropsGroup.parseOOXML(xpp, lastTag) as EffectPropsGroup
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "spPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("spPr.parseOOXML: $e")
            }

            return SpPr(x, geom, fill, l, effect, bwMode, ns)
        }
    }
}

