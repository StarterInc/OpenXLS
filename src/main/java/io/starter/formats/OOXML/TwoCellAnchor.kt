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

import io.starter.OpenXLS.ColHandle
import io.starter.OpenXLS.RowHandle
import io.starter.OpenXLS.WorkBookHandle
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Stack

/**
 * twoCellAnchor (Two Cell Anchor Shape Size)
 *
 *
 * This element specifies a two cell anchor placeholder for a group, a shape, or a drawing element.
 * It moves with cells and its extents are in EMU units.
 *
 *
 * This is the root element for charts, images and shapes.
 *
 *
 * parent: wsDr
 * children: from, to, OBJECTCHOICES (sp, grpSp, graphicFrame, cxnSp, pic), clientData
 */
//TODO: finish grpSp Group Shape
// TODO: finish clientData element
class TwoCellAnchor : OOXMLElement {
    /**
     * return the editAs editMovement attribute
     *
     * @return
     */
    /**
     * set the editAs editMovement attribute
     *
     * @return
     */
    var editAs: String? = null
    /**
     * return the Embedded Object's filename as saved on disk
     *
     * @return
     */
    /**
     * set the Embedded Object's filename as saved on disk
     *
     * @param embed
     */
    var embedFilename: String? = null
    private var from: From? = null
    private var to: To? = null
    private var o: ObjectChoice? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:twoCellAnchor")
            if (editAs != null) ooxml.append(" editAs=\"$editAs\"")
            ooxml.append(">")
            if (from != null) ooxml.append(from!!.ooxml)
            if (to != null) ooxml.append(to!!.ooxml)
            ooxml.append(o!!.ooxml)
            ooxml.append("<xdr:clientData/>")
            ooxml.append("</xdr:twoCellAnchor>")
            return ooxml.toString()
        }

    // access methods ******

    /**
     * return the (to, from) bounds of this object
     * by concatenating the bounds for the to and the from
     * bounds represent:  COL, COLOFFSET, ROW, ROWOFFST, COL1, COLOFFSET1, ROW1, ROWOFFSET1
     *
     * @return bounds short[8]  from [4] and to [4]
     */
    /**
     * set the (to, from) bounds of this object
     * bounds represent:  COL, COLOFFSET, ROW, ROWOFFST, COL1, COLOFFSET1, ROW1, ROWOFFSET1
     * NOTE: COL
     *
     * @param bounds int[8]  from [4] and to [4]
     */
    // from bounds
    // to bounds
    var bounds: IntArray
        get() {
            val bounds = IntArray(8)
            System.arraycopy(from!!.getBounds(), 0, bounds, 0, 4)
            System.arraycopy(to!!.getBounds(), 0, bounds, 4, 4)
            return bounds
        }
        set(bounds) {
            val b = IntArray(4)
            System.arraycopy(bounds, 0, b, 0, 4)
            if (from == null)
                from = From(b)
            else
                from!!.setBounds(b)
            System.arraycopy(bounds, 4, b, 0, 4)
            if (to == null)
                to = To(b)
            else
                to!!.setBounds(b)
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
        get() = if (o != null) o!!.name else null
        set(name) {
            if (o != null)
                o!!.name = name
        }

    /**
     * get cNvPr descr attribute
     *
     * @return
     */
    /**
     * set cNvPr descr attribute
     * sometimes associated with shape name
     *
     * @param descr
     */
    var descr: String?
        get() = if (o != null) o!!.descr else null
        set(descr) {
            if (o != null)
                o!!.descr = descr
        }

    /**
     * get macro attribute
     *
     * @return
     */
    /**
     * set Macro attribute
     * sometimes associated with shape name
     *
     * @param descr
     */
    var macro: String?
        get() = if (o != null) o!!.macro else null
        set(macro) {
            if (o != null)
                o!!.macro = macro
        }

    /**
     * get the URI associated with this graphic Data
     */
    /**
     * set the URI associated with this graphic data
     *
     * @param uri
     */
    var uri: String?
        get() = if (o != null) o!!.uri else null
        set(uri) {
            if (o != null)
                o!!.uri = uri
        }

    /**
     * return the rid for the chart defined by this twocellanchor
     *
     * @return
     */
    var chartRId: String?
        get() = if (o != null) o!!.chartRId else null
        set(rId) {
            if (o != null) o!!.chartRId = rId
        }

    /**
     * return the id for the embedded picture or shape (i.e. resides within the file)
     *
     * @return
     */
    /**
     * set the embed for the embedded picture or shape (i.e. resides within the file)
     *
     * @param embed
     */
    var embed: String?
        get() = if (o != null) o!!.embed else null
        set(embed) {
            if (o != null) o!!.embed = embed
        }


    /**
     * return the id for the linked object (i.e. doesn't reside in the file)
     *
     * @return
     */
    /**
     * set the link attribute for this blip (the id for the linked picture)
     *
     * @param embed
     */
    var link: String?
        get() = if (o != null) o!!.link else null
        set(link) {
            if (o != null) o!!.link = link
        }

    /**
     * utility to return the shape properties element (picture element only)
     * should be depreciated when OOXML is completely distinct from BIFF8
     *
     * @return
     */
    val sppr: SpPr?
        get() = if (o != null) o!!.sppr else null

    constructor(editAs: String) {
        this.editAs = editAs
    }

    constructor(editAs: String, f: From, t: To, o: ObjectChoice) {
        this.editAs = editAs
        this.from = f
        this.to = t
        this.o = o
    }

    constructor(tce: TwoCellAnchor) {
        this.editAs = tce.editAs
        this.from = tce.from
        this.to = tce.to
        this.o = tce.o
    }

    override fun cloneElement(): OOXMLElement {
        return TwoCellAnchor(this)
    }

    override fun toString(): String? {
        return this.name
    }

    /**
     * return if this twoCellAnchor element refers to an image rather than a chart or shape
     *
     * @return
     */
    fun hasImage(): Boolean {
        return o!!.hasImage()
    }

    /**
     * return if this twoCellAnchor element refers to a chart as opposed to a shape or image
     *
     * @return
     */
    fun hasChart(): Boolean {
        return if (o != null) o!!.hasChart() else false
    }

    /**
     * return if this twoCellAnchor element refers to a shape, as opposed a chart or an image
     *
     * @return
     */
    fun hasShape(): Boolean {
        return if (o != null) o!!.hasShape() else false
    }

    /**
     * set this twoCellAnchor as a chart element
     * used for
     *
     * @param rid
     * @param name
     * @param bounds
     */
    fun setAsChart(rid: Int, name: String, bounds: IntArray) {
        o = ObjectChoice()
        o!!.`object` = GraphicFrame()
        o!!.name = name
        o!!.chartRId = "rId" + Integer.valueOf(rid)!!.toString()
        o!!.id = rid
        this.bounds = bounds
        // id???
    }

    /**
     * set this twoCellAnchor as an image
     *
     * @param rid
     * @param name
     * @param id
     */
    fun setAsImage(rid: Int, name: String, descr: String, spid: Int, sp: SpPr?) {
        o = ObjectChoice()
        o!!.`object` = Pic()
        o!!.name = name
        o!!.descr = descr
        o!!.embed = "rId" + Integer.valueOf(rid)!!.toString()
        o!!.id = spid
        if (sp != null) {
            (o!!.`object` as Pic).sppr = sp
        }
    }

    companion object {

        private val serialVersionUID = 4180396678197959710L
        // EMU = pixel * 914400 / Resolution (96?)
        val EMU: Short = 1270

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>, bk: WorkBookHandle): OOXMLElement {
            var editAs: String? = null
            var f: From? = null
            var t: To? = null
            var o: ObjectChoice? = null
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "twoCellAnchor") {        // get attributes
                            if (xpp.attributeCount > 0)
                                editAs = xpp.getAttributeValue(0)
                        } else if (tnm == "from") {
                            lastTag.push(tnm)
                            f = From.parseOOXML(xpp, lastTag)
                        } else if (tnm == "to") {
                            lastTag.push(tnm)
                            t = To.parseOOXML(xpp, lastTag)
                        } else if (tnm == "cxnSp" ||
                                tnm == "graphicFrame" ||
                                tnm == "grpSp" ||
                                tnm == "pic" ||
                                tnm == "sp") {
                            o = ObjectChoice.parseOOXML(xpp, lastTag, bk) as ObjectChoice
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "twoCellAnchor") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("twoCellAnchor.parseOOXML: $e")
            }

            return TwoCellAnchor(editAs, f, t, o)
        }


        /**
         * given bounds[8] in BIFF 8 coordinates and convert to OOXML units
         * mostly this means adjusting ROWOFFSET 0 & 1 (bounds[5] & bounds[7]
         * and COLOFFSET 0 & 1 (bounds[1] & bounds[5] to EMU units +
         * adjust according to emperical calc garnered from observation
         *
         * @param sheet
         * @param bbounds short[] bounds in BIFF8 units
         * @return new bounds int[] bounds in OOXML units
         */
        fun convertBoundsFromBIFF8(sheet: io.starter.formats.XLS.Boundsheet, bbounds: ShortArray): IntArray {
            // note on bounds:
            // -- offsets (bounds 1,3,5 & 7) are %'s of column width or row height, respectively
            // below calculations are garnered from OOXML info + by comparing to Excel's values ... may not be 100% but appears to work for both 2003 and 2007 versions
            // NOTE: if change offset algorithm here must modify algorithm in twoCellAnchor from and to classes
            val bounds = IntArray(8)
            bounds[0] = bbounds[0].toInt()
            var cw = ColHandle.getWidth(sheet, bbounds[0].toInt()).toDouble()    ///5.0;
            // below is more correct (???)		bounds[1]= (int)(EMU*cw*(bbounds[1]/1024.0));
            bounds[1] = (cw * 256.0 * (bbounds[1] / 1024.0)).toInt()
            var rh = RowHandle.getHeight(sheet, bbounds[2].toInt()).toDouble() ///2.0;
            bounds[2] = bbounds[2].toInt()
            bounds[3] = (EMU.toDouble() * rh * (bbounds[3] / 256.0)).toInt()
            cw = ColHandle.getWidth(sheet, bbounds[4].toInt()).toDouble()    ///5.0;
            bounds[4] = bbounds[4].toInt()
            //		bounds[5]= (int)(EMU*cw*(bbounds[5]/1024.0));
            bounds[5] = (cw * 256.0 * (bbounds[5] / 1024.0)).toInt()
            rh = RowHandle.getHeight(sheet, bbounds[6].toInt()).toDouble()    ///2.0;
            bounds[6] = bbounds[6].toInt()
            bounds[7] = (EMU.toDouble() * rh * (bbounds[7] / 256.0)).toInt()
            return bounds
        }

        /**
         * convert bounds[8] from OOXML to BIFF8 Units
         * basically must adjust COLOFFSETs and ROWOFFSETs to BIFF8 units
         *
         * @param sheet
         * @param bounds int[]
         * @return bounds short[] (BIFF8 uses short[] + different units)
         */
        fun convertBoundsToBIFF8(sheet: io.starter.formats.XLS.Boundsheet, bounds: IntArray): ShortArray {
            val bbounds = ShortArray(8)
            bbounds[0] = bounds[0].toShort()
            var cw = ColHandle.getWidth(sheet, bounds[0]).toDouble()    ///5.0;
            // below is more correct (???)		bbounds[1]= (short)((bounds[1]*1024)/(TwoCellAnchor.EMU*cw));
            bbounds[1] = (bounds[1] * 1024.0 / (cw * 256)).toShort()
            bbounds[2] = bounds[2].toShort()
            var rh = RowHandle.getHeight(sheet, bounds[2]).toDouble()    ///2.0;
            bbounds[3] = (bounds[3] * 256.0 / (TwoCellAnchor.EMU * rh)).toShort()
            bbounds[4] = bounds[4].toShort()
            cw = ColHandle.getWidth(sheet, bounds[4]).toDouble()    ///5.0;
            //    	bbounds[5]= (short)((bounds[5]*1024)/(TwoCellAnchor.EMU*cw));
            bbounds[5] = (bounds[5] * 1024.0 / (cw * 256)).toShort()
            bbounds[6] = bounds[6].toShort()
            rh = RowHandle.getHeight(sheet, bounds[6]).toDouble()    ///2.0;
            bbounds[7] = (bounds[7] * 256.0 / (TwoCellAnchor.EMU * rh)).toShort()
            return bbounds
        }
    }
}

/**
 * from (Starting Anchor Point)
 * This element specifies the first anchor point for the drawing element. This will be used to anchor the top and left
 * sides of the shape within the spreadsheet. That is when the cell that is specified in the from element is adjusted,
 * the shape will also be adjusted.
 *
 *
 * NOTE: Coordinates are in OOXML units; they are converted to BIFF8 units using twoCellAnchor.convertBoundsToBIFF8
 * parent: oneCellAnchor, twoCellAnchor
 * children: col, colOff, row, rowOff
 */
internal class From : OOXMLElement {
    private var bounds: IntArray? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:from>")
            ooxml.append("<xdr:col>" + bounds!![0] + "</xdr:col>")
            ooxml.append("<xdr:colOff>" + bounds!![1] + "</xdr:colOff>")
            ooxml.append("<xdr:row>" + bounds!![2] + "</xdr:row>")
            ooxml.append("<xdr:rowOff>" + bounds!![3] + "</xdr:rowOff>")
            ooxml.append("</xdr:from>")
            return ooxml.toString()
        }


    constructor(bounds: IntArray) {
        this.bounds = IntArray(4)
        System.arraycopy(bounds, 0, this.bounds!!, 0, 4)
    }

    constructor(f: From) {
        this.bounds = f.bounds
    }

    override fun cloneElement(): OOXMLElement {
        return From(this)
    }

    fun getBounds(): IntArray {
        if (bounds == null) bounds = IntArray(4)
        return bounds
    }

    fun setBounds(bounds: IntArray) {
        System.arraycopy(bounds, 0, this.bounds!!, 0, 4)
    }

    companion object {

        private val serialVersionUID = -4776435343244555855L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<*>): From {
            val bounds = IntArray(4)
            var boundsidx = 0
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "col") {
                            bounds[boundsidx++] = Integer.parseInt(xpp.nextText())
                        } else if (tnm == "colOff") {
                            bounds[boundsidx++] = Integer.parseInt(xpp.nextText())
                        } else if (tnm == "row") {
                            bounds[boundsidx++] = Integer.parseInt(xpp.nextText())
                        } else if (tnm == "rowOff") {
                            bounds[boundsidx++] = Integer.parseInt(xpp.nextText())
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "from") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("from.parseOOXML: $e")
            }

            return From(bounds)
        }
    }
}

/**
 * to (Ending Anchor Point)
 * This element specifies the second anchor point for the drawing element. This will be used to anchor the bottom
 * and right sides of the shape within the spreadsheet. That is when the cell that is specified in the to element is
 * adjusted, the shape will also be adjusted.to (Ending Anchor Point)
 * This element specifies the second anchor point for the drawing element. This will be used to anchor the bottom
 * and right sides of the shape within the spreadsheet. That is when the cell that is specified in the to element is
 * adjusted, the shape will also be adjusted.
 *
 *
 * NOTE: Coordinates are in OOXML units; they are converted to BIFF8 units using twoCellAnchor.convertBoundsToBIFF8
 * parent: twoCellAnchor
 * children: col, colOff, row, rowOff
 */
internal class To : OOXMLElement {
    private var bounds: IntArray? = null

    override val ooxml: String
        get() {
            val ooxml = StringBuffer()
            ooxml.append("<xdr:to>")
            ooxml.append("<xdr:col>" + bounds!![0] + "</xdr:col>")
            ooxml.append("<xdr:colOff>" + bounds!![1] + "</xdr:colOff>")
            ooxml.append("<xdr:row>" + bounds!![2] + "</xdr:row>")
            ooxml.append("<xdr:rowOff>" + bounds!![3] + "</xdr:rowOff>")
            ooxml.append("</xdr:to>")
            return ooxml.toString()
        }

    constructor(bounds: IntArray) {
        this.bounds = IntArray(4)
        System.arraycopy(bounds, 0, this.bounds!!, 0, 4)
    }

    constructor(f: To) {
        this.bounds = f.bounds
    }

    override fun cloneElement(): OOXMLElement {
        return To(this)
    }

    // Access Methods
    fun getBounds(): IntArray {
        if (this.bounds == null) this.bounds = IntArray(4)
        return bounds
    }

    fun setBounds(bounds: IntArray) {
        this.bounds = IntArray(4)
        System.arraycopy(bounds, 0, this.bounds!!, 0, 4)
    }

    companion object {

        private val serialVersionUID = 1500243445505400113L

        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<*>): To {
            val bounds = IntArray(4)
            var boundsidx = 0
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "col") {
                            bounds[boundsidx++] = Integer.parseInt(xpp.nextText())
                        } else if (tnm == "colOff") {
                            bounds[boundsidx++] = Integer.parseInt(xpp.nextText())
                        } else if (tnm == "row") {
                            bounds[boundsidx++] = Integer.parseInt(xpp.nextText())
                        } else if (tnm == "rowOff") {
                            bounds[boundsidx++] = Integer.parseInt(xpp.nextText())
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "to") {
                            lastTag.pop()
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("to.parseOOXML: $e")
            }

            return To(bounds)
        }
    }
}

