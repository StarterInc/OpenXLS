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

import java.util.ArrayList
import java.util.HashMap
import java.util.Stack

/**
 * grpSp (Group Shape)
 * This element specifies a group shape that represents many shapes grouped together.
 * Within a group shape 1 each of the shapes that make up the group are
 * specified just as they normally would. The idea behind grouping elements however is that a
 * single transform can apply to many shapes at the same time.
 *
 *
 * parents: absoluteAnchor, grpSp, oneCellAnchor, twoCellAnchor
 * children: nvGrpSpPr (req), grpSpPr (req), choice of: (sp, grpSp, graphicFrame, cxnSp, pic) 0 to unbounded
 */
class GrpSp : OOXMLElement {
    private var nvpr: NvGrpSpPr? = null
    private var sppr: GrpSpPr? = null
    private var choice: ArrayList<OOXMLElement>? = null

    /**
     * return grpSp element OOXML
     */
    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:grpSp>")
            ooxml.append(nvpr!!.ooxml)
            ooxml.append(sppr!!.ooxml)
            if (choice != null) {
                for (i in choice!!.indices) {
                    ooxml.append(choice!![i].ooxml)
                }
            }
            ooxml.append("</xdr:grpSp>")
            return ooxml.toString()
        }

    /**
     * get name attribute of the group shape
     *
     * @return
     */
    /**
     * set name attribute of the group shape
     *
     * @param name
     */
    // set the group name
    var name: String?
        get() = nvpr!!.name
        set(name) {
            nvpr!!.name = name
        }

    /**
     * get macro attribute of the group shape
     *
     * @return
     */
    /**
     * set macro attribute of the group shape
     *
     * @param name
     */
    // Have a shape with an embed?
    // how's about a picture?
    // or a connection shape?
    // Have a shape with a macro?
    // how's about a picture?
    // or a connection shape?
    var macro: String?
        get() {
            var macro: String? = null
            var oe = getObject(SP)
            if (oe != null)
                macro = (oe as Sp).macro
            if (macro == null)
                oe = getObject(PIC)
            if (oe != null && macro == null)
                macro = (oe as Pic).macro
            if (macro == null)
                oe = getObject(CXN)
            if (oe != null && macro == null)
                macro = (oe as CxnSp).macro
            return macro
        }
        set(macro) {
            var m: String? = null
            var oe = getObject(SP)
            if (oe != null)
                m = (oe as Sp).macro
            if (m != null) {
                (oe as Sp).macro = macro
                return
            }
            oe = getObject(PIC)
            if (oe != null)
                m = (oe as Pic).macro
            if (m != null) {
                (oe as Pic).macro = macro
                return
            }
            oe = getObject(CXN)
            if (oe != null)
                m = (oe as CxnSp).macro
            if (m != null)
                (oe as CxnSp).macro = macro
        }


    /**
     * get cthe descr attribute of the group shape
     *
     * @return
     */
    /**
     * set description attribute of the group shape
     * sometimes associated with shape name
     *
     * @param descr
     */
    var descr: String?
        get() = nvpr!!.descr
        set(descr) {
            nvpr!!.descr = descr
        }

    /**
     * return the rid for the embedded object (picture or chart or picture shape) (i.e. resides within the file)
     *
     * @return
     */
    /**
     * set the rid for this object (picture or chart or picture shape) (resides within the file)
     *
     * @param embed
     */
    // Have a shape with an embed?
    // how's about a picture?
    // or a connection shape?
    // Have a shape with an embed?
    // how's about a picture?
    // or a connection shape?
    var embed: String?
        get() {
            var embed: String? = null
            var oe = getObject(SP)
            if (oe != null)
                embed = (oe as Sp).embed
            if (embed == null)
                oe = getObject(PIC)
            if (oe != null && embed == null)
                embed = (oe as Pic).embed
            if (embed == null)
                oe = getObject(CXN)
            if (oe != null && embed == null)
                embed = (oe as CxnSp).embed
            return embed
        }
        set(embed) {
            var e: String? = null
            var oe = getObject(SP)
            if (oe != null)
                e = (oe as Sp).embed
            if (e != null) {
                (oe as Sp).embed = embed
                return
            }
            oe = getObject(PIC)
            if (oe != null)
                e = (oe as Pic).embed
            if (e != null) {
                (oe as Pic).embed = embed
                return
            }
            oe = getObject(CXN)
            if (oe != null)
                e = (oe as CxnSp).embed
            if (e != null)
                (oe as CxnSp).embed = embed
        }

    /**
     * return the id for the linked picture (i.e. doesn't reside in the file)
     *
     * @return
     */
    /**
     * rset the id for the linked picture (i.e. doesn't reside in the file)
     *
     * @return
     */
    // Have a shape with an embed?
    // how's about a picture?
    // or a connection shape?
    // Have a shape with an embed?
    // how's about a picture?
    // or a connection shape?
    var link: String?
        get() {
            var link: String? = null
            var oe = getObject(SP)
            if (oe != null)
                link = (oe as Sp).link
            if (link == null)
                oe = getObject(PIC)
            if (oe != null && link == null)
                link = (oe as Pic).link
            if (link == null)
                oe = getObject(CXN)
            if (oe != null && link == null)
                link = (oe as CxnSp).link
            return link
        }
        set(link) {
            var l: String? = null
            var oe = getObject(SP)
            if (oe != null)
                l = (oe as Sp).link
            if (l != null) {
                (oe as Sp).link = link
                return
            }
            oe = getObject(PIC)
            if (oe != null)
                l = (oe as Pic).link
            if (l != null) {
                (oe as Pic).link = link
                return
            }
            oe = getObject(CXN)
            if (oe != null)
                l = (oe as CxnSp).link
            if (l != null)
                (oe as CxnSp).link = link
        }

    /**
     * return the rid of the chart element, if exists
     *
     * @return
     */
    val chartRId: String?
        get() {
            val oe = getObject(GRAPHICFRAME)
            return if (oe != null) (oe as GraphicFrame).chartRId else null
        }

    /**
     * return the cNvPr id for this element
     *
     * @return
     */
    /**
     * set the cNvPr id for this element
     *
     * @param id
     */
    // get the first??
    // get the first??
    var id: Int
        get() {

            if (choice != null) {
                val oe = choice!![0]
                if (oe is Sp)
                    return oe.id
                else if (oe is Pic)
                    return oe.id
                else if (oe is GraphicFrame)
                    return oe.id
                else if (oe is CxnSp)
                    return oe.id
                else if (oe is GrpSp)
                    return oe.id
            }
            return -1
        }
        set(id) {
            if (choice != null) {
                val oe = choice!![0]
                if (oe is Sp)
                    oe.id = id
                else if (oe is Pic)
                    oe.id = id
                else if (oe is GraphicFrame)
                    oe.id = id
                else if (oe is CxnSp)
                    oe.id = id
                else if (oe is GrpSp)
                    oe.id = id
            }
        }

    constructor(nvpr: NvGrpSpPr, sppr: GrpSpPr, choice: ArrayList<OOXMLElement>) {
        this.nvpr = nvpr
        this.sppr = sppr
        this.choice = choice
    }

    constructor(g: GrpSp) {
        this.nvpr = g.nvpr
        this.sppr = g.sppr
        this.choice = g.choice
    }

    private fun getObject(type: Int): OOXMLElement? {
        if (choice != null) {
            for (i in choice!!.indices) {
                val oe = choice!![i]
                if (oe is Sp && type == SP)
                    return oe
                else if (oe is Pic && type == PIC)
                    return oe
                else if (oe is CxnSp && type == CXN)
                    return oe
                else if (oe is GraphicFrame && type == GRAPHICFRAME)
                    return oe
                else if (oe is GrpSp)
                    return oe.getObject(type)

            }
        }
        return null
    }

    /**
     * set the URI associated with this graphic data
     *
     * @param uri
     */
    fun setURI(uri: String) {
        val oe = getObject(GRAPHICFRAME)
        (oe as GraphicFrame).uri = uri
    }

    /**
     * utility to return the shape properties element for this picture
     * should be depreciated when OOXML is completely distinct from BIFF8
     *
     * @return
     */
    fun getSppr(): SpPr? {
        val oe = getObject(PIC)
        return if (oe != null) (oe as Pic).sppr else null

    }

    /**
     * return if this Group refers to a shape, as opposed a chart or an image
     *
     * @return
     */
    fun hasShape(): Boolean {
        return getObject(SP) != null || getObject(CXN) != null
    }

    override fun cloneElement(): OOXMLElement {
        return GrpSp(this)
    }

    companion object {

        private val serialVersionUID = -3276180769601314853L

        /**
         * parse grpSp element OOXML
         *
         * @param xpp
         * @param lastTag
         * @return
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): GrpSp {
            var nvpr: NvGrpSpPr? = null
            var sppr: GrpSpPr? = null
            val choice = ArrayList<OOXMLElement>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "nvGrpSpPr") {
                            lastTag.push(tnm)
                            nvpr = NvGrpSpPr.parseOOXML(xpp, lastTag)
                        } else if (tnm == "grpSpPr") {
                            lastTag.push(tnm)
                            sppr = GrpSpPr.parseOOXML(xpp, lastTag, bk)
                            // choice of: 0 or more below:
                        } else if (tnm == "sp") {
                            lastTag.push(tnm)
                            choice.add(Sp.parseOOXML(xpp, lastTag, bk))
                        } else if (tnm == "grpSp") {
                            if (nvpr != null) { // if not the initial start attribute, this is a child
                                lastTag.push(tnm)
                                choice.add(GrpSp.parseOOXML(xpp, lastTag, bk))
                            }
                        } else if (tnm == "graphicFrame") {
                            lastTag.push(tnm)
                            choice.add(GraphicFrame.parseOOXML(xpp, lastTag))
                        } else if (tnm == "cxnSp") {
                            lastTag.push(tnm)
                            choice.add(CxnSp.parseOOXML(xpp, lastTag, bk))
                        } else if (tnm == "pic") {
                            lastTag.push(tnm)
                            choice.add(Pic.parseOOXML(xpp, lastTag, bk))
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "grpSp") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("GrpSp.parseOOXML: $e")
            }

            return GrpSp(nvpr, sppr, choice)
        }

        internal var SP = 0
        internal var PIC = 1
        internal var CXN = 2
        internal var GRAPHICFRAME = 3
    }
}

/**
 * nvGrpSpPr 1 (Non-Visual Properties for a Group Shape)
 */
internal class NvGrpSpPr : OOXMLElement {
    var cp: CNvPr? = null
    var cgrpsppr: CNvGrpSpPr? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:nvGrpSpPr>")
            ooxml.append(cp!!.ooxml)
            ooxml.append(cgrpsppr!!.ooxml)
            ooxml.append("</xdr:nvGrpSpPr>")
            return ooxml.toString()
        }

    /**
     * get cNvPr name attribute
     *
     * @return
     */
    /**
     * set cNvPr name attribute
     *
     * @param name
     */
    var name: String?
        get() = if (cp != null) cp!!.name else null
        set(name) {
            if (cp != null)
                cp!!.name = name
        }

    /**
     * get cNvPr descr attribute
     *
     * @return
     */
    /**
     * set cNvPr desc attribute
     *
     * @param name
     */
    var descr: String?
        get() = if (cp != null) cp!!.descr else null
        set(descr) {
            if (cp != null)
                cp!!.descr = descr
        }

    /**
     * return the cNvPr id for this element
     *
     * @return
     */
    /**
     * set the cNvPr id for this element
     *
     * @param id
     */
    var id: Int
        get() = if (cp != null) cp!!.id else -1
        set(id) {
            if (cp != null)
                cp!!.id = id
        }

    constructor(cp: CNvPr, cgrpsppr: CNvGrpSpPr) {
        this.cp = cp
        this.cgrpsppr = cgrpsppr
    }

    constructor(g: NvGrpSpPr) {
        this.cp = g.cp
        this.cgrpsppr = g.cgrpsppr
    }

    override fun cloneElement(): OOXMLElement {
        return NvGrpSpPr(this)
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -4404072268706949318L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): NvGrpSpPr {
            var cp: CNvPr? = null
            var cgrpsppr: CNvGrpSpPr? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "cNvPr") {
                            lastTag.push(tnm)
                            cp = CNvPr.parseOOXML(xpp, lastTag) as CNvPr
                        } else if (tnm == "cNvGrpSpPr") {
                            lastTag.push(tnm)
                            cgrpsppr = CNvGrpSpPr.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "nvGrpSpPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("NvGrpSpPr.parseOOXML: $e")
            }

            return NvGrpSpPr(cp, cgrpsppr)
        }
    }


}

/**
 * grpSpPr (Group Shape Properties)
 * This element specifies the properties that are to be common across all of the shapes
 * within the corresponding group. If there are any conflicting properties within the
 * group shape properties and the individual shape properties then the individual shape properties
 * should take precedence.
 *
 *
 * children:  xfrm, fillproperties (group), effectproperties (group), scene3d, extLst (all optional)
 */
// TODO: FINISH scene3d, extLst
internal class GrpSpPr : OOXMLElement {
    private var xf: Xfrm? = null
    private var bwmode: String? = null
    private var fill: FillGroup? = null
    private var effect: EffectPropsGroup? = null

    override// TODO: Finish
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:grpSpPr")
            if (bwmode != null) {
                ooxml.append(" bwMode=\"$bwmode\"")
            }
            ooxml.append(">")
            if (xf != null) ooxml.append(xf!!.ooxml)
            if (fill != null) ooxml.append(fill!!.ooxml)
            if (effect != null) ooxml.append(effect!!.ooxml)
            ooxml.append("</xdr:grpSpPr>")
            return ooxml.toString()
        }

    constructor(xf: Xfrm, bwmode: String, fill: FillGroup, effect: EffectPropsGroup) {
        this.xf = xf
        this.bwmode = bwmode
        this.fill = fill
        this.effect = effect
    }

    constructor(g: GrpSpPr) {
        this.xf = g.xf
        this.bwmode = g.bwmode
        this.fill = g.fill
        this.effect = g.effect
    }

    override fun cloneElement(): OOXMLElement {
        return GrpSpPr(this)
    }

    companion object {

        private val serialVersionUID = 7464871024304781512L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): GrpSpPr {
            var xf: Xfrm? = null
            var bwmode: String? = null
            var fill: FillGroup? = null
            var effect: EffectPropsGroup? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "grpSpPr") {        // get attributes
                            if (xpp.attributeCount == 1) {
                                bwmode = xpp.getAttributeValue(0)
                            }
                        } else if (tnm == "xfrm") {
                            lastTag.push(tnm)
                            xf = Xfrm.parseOOXML(xpp, lastTag) as Xfrm
                            // scene3d, extLst- finish
                        } else if (tnm == "solidFill" ||
                                tnm == "noFill" ||
                                tnm == "gradFill" ||
                                tnm == "grpFill" ||
                                tnm == "pattFill" ||
                                tnm == "blipFill") {
                            lastTag.push(tnm)
                            fill = FillGroup.parseOOXML(xpp, lastTag, bk) as FillGroup
                        } else if (tnm == "effectLst" || tnm == "effectDag") {
                            lastTag.push(tnm)
                            effect = EffectPropsGroup.parseOOXML(xpp, lastTag) as EffectPropsGroup
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "grpSpPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("GrpSpPr.parseOOXML: $e")
            }

            return GrpSpPr(xf, bwmode, fill, effect)
        }
    }
}

/**
 * cNvGrpSpPr (Non-Visual Group Shape Drawing Properties)
 *
 *
 * parent:		nvGrpSpPr
 * children: 	grpSpLocks, extLst
 */
internal class CNvGrpSpPr : OOXMLElement {
    private var gsl: GrpSpLocks? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:cNvGrpSpPr>")
            if (gsl != null)
                ooxml.append(gsl!!.ooxml)
            ooxml.append("</xdr:cNvGrpSpPr>")
            return ooxml.toString()
        }

    constructor(gsl: GrpSpLocks) {
        this.gsl = gsl
    }

    constructor(g: CNvGrpSpPr) {
        this.gsl = g.gsl
    }

    override fun cloneElement(): OOXMLElement {
        return CNvGrpSpPr(this)
    }

    companion object {

        private val serialVersionUID = -1106010927060582127L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): CNvGrpSpPr {
            var gsl: GrpSpLocks? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "grpSpLocks") {
                            lastTag.push(tnm)
                            gsl = GrpSpLocks.parseOOXML(xpp, lastTag)
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "cNvGrpSpPr") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("CNvGrpSpPr.parseOOXML: $e")
            }

            return CNvGrpSpPr(gsl)
        }
    }
}


/**
 * grpSpLocks (Group Shape Locks)
 * This element specifies all locking properties for a connection shape. These properties inform the generating
 * application about specific properties that have been previously locked and thus should not be changed.
 *
 *
 * parent:		cNvGrpSpPr
 * attributes:  many, all optional
 * children: 	extLst
 */
internal class GrpSpLocks : OOXMLElement {
    private var attrs: HashMap<String, String>? = null

    override// attributes
    val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:grpSpLocks>")
            val i = attrs!!.keys.iterator()
            while (i.hasNext()) {
                val key = i.next()
                val `val` = attrs!![key]
                ooxml.append(" $key=\"$`val`\"")
            }
            ooxml.append("/>")
            return ooxml.toString()
        }

    constructor(attrs: HashMap<String, String>) {
        this.attrs = attrs
    }

    constructor(g: GrpSpLocks) {
        this.attrs = g.attrs
    }

    override fun cloneElement(): OOXMLElement {
        return GrpSpLocks(this)
    }

    companion object {

        private val serialVersionUID = -2592038952923879415L


        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>): GrpSpLocks {
            val attrs = HashMap<String, String>()
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "grpSpLocks") {        // get attributes
                            for (i in 0 until xpp.attributeCount) {
                                attrs[xpp.getAttributeName(i)] = xpp.getAttributeValue(i)
                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "grpSpLocks") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("CNvGrpSpPr.parseOOXML: $e")
            }

            return GrpSpLocks(attrs)
        }
    }
}


