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
package io.starter.formats.XLS.charts

import io.starter.formats.OOXML.OOXMLElement
import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser

import java.util.Stack

/**
 * **ThreeD(3D) Chart group is a 3-D Chart Group (0x103A)**
 *
 *
 * anRot 2 		Rotation Angle (0 to 360 degrees), usually 0 for pie, 20 for others  -- def = 0
 * anElev 2 	Elevation Angle (-90 to 90 degrees) (15 is default) 8
 * pcDist 2		Distance from eye to chart (0 to 100) (30 is default) 10
 * pcHeight 2 	Height of plot volume relative to width and depth (100 is default) 12
 * pcDepth 2 	Depth of points relative to width (100 is default) 14
 * pcGap 2 		Space between points  (150 is default - should be 50!!!) 16 grbit 2
 *
 *
 * grbit
 * 0 0x1 fPerspective 1= use perspective transform
 * 1 0x2 fCluster 1= 3-D columns are clustered or stacked
 * 2 0x4 f3DScaling 1= use auto-scaling
 * 3 reserved
 * 4 0x8 fNotPieChart 1= NOT a pie chart
 * 5 0x10 f2DWalls use 2D walls and gridlines (if fPerspective MUST be ignored. if not of type BAR, AREA or
 * PIE, ignore. if BAR and fCluster=0, ignore. specifies whether the walls are rendered in 2-D.
 * If fPerspective is 1 then this MUST be ignored. If the chart group type is not bar, area or pie this MUST be ignored. If the chart
 * group is of type bar and fCluster is 0, then this MUST be ignored. if PIE MUST be 0.
 *
 *
 * "http://www.extentech.com">Extentech Inc.
 */
class ThreeD : GenericChartObject(), ChartObject {
    private var anRot: Short = 0
    private var anElev: Short = 15
    private var pcDist: Short = 30
    private var pcHeight: Short = 100
    private var pcDepth: Short = 100
    private var pcGap: Short = 150
    private var grbit: Short = 0 //
    private var fPerspective: Boolean = false
    private var fCluster: Boolean = false
    private var f3dScaling: Boolean = false
    private var f2DWalls: Boolean = false // 20070905 KSC: parse grbit

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0, 0, 0, 30, 0, 100, 0, 100, 0, -106, 0, 0, 0)

    /**
     * @return String XML representation of this chart-type's options
     */
    override// 20070913 KSC: need to track Cluster, whether on or off
    // sb.append(" formatOptions=\"" + grbit+ "\"");
    val optionsXML: String
        get() {
            val sb = StringBuffer()
            if (anRot.toInt() != 0)
                sb.append(" AnRot=\"$anRot\"")
            if (anElev.toInt() != 15)
                sb.append(" AnElev=\"$anElev\"")
            if (pcDist.toInt() != 30)
                sb.append(" pcDist=\"$pcDist\"")
            if (pcHeight.toInt() != 100)
                sb.append(" pcHeight=\"$pcHeight\"")
            if (pcDepth.toInt() != 100)
                sb.append(" pcDepth=\"$pcDepth\"")
            if (pcGap.toInt() != 150)
                sb.append(" pcGap=\"$pcGap\"")
            if (fPerspective)
                sb.append(" Perspective=\"true\"")
            if (f3dScaling)
                sb.append(" ThreeDScaling=\"true\"")
            if (f2DWalls)
                sb.append(" TwoDWalls=\"true\"")
            sb.append(" Cluster=\"$fCluster\"")
            return sb.toString()
        }

    /**
     * @return truth of "Chart is Clustered"
     */
    /**
     * sets if this chart has clustered bar/columns
     *
     * @param bIsClustered
     */
    var isClustered: Boolean
        get() = fCluster
        set(bIsClustered) {
            fCluster = bIsClustered
            grbit = ByteTools.updateGrBit(grbit, fCluster, 1)
            val b = ByteTools.shortToLEBytes(grbit)
            this.data[12] = b[0]
            this.data[13] = b[1]
        }

    /**
     * return view3d OOXML representation
     *
     * @return
     */
    // rotX == anElev
    // default
    // hPercent -- a height percent between 5 and 500.
    // rotY == anRot
    // default
    // depthPercentage -- This element specifies the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent).
    // rAngAx == !fPerspective
    // perspective == pcDist
    // default
    val ooxml: StringBuffer
        get() {
            val cooxml = StringBuffer()
            cooxml.append("<c:view3D>")
            cooxml.append("\r\n")
            if (anElev.toInt() != 0)
                cooxml.append("<c:rotX val=\"$anElev\"/>")
            if (anRot.toInt() != 0 || anElev.toInt() != 0)
                cooxml.append("<c:rotY val=\"$anRot\"/>")
            if (pcDepth.toInt() != 100)
                cooxml.append("<c:depthPercent val=\"$pcDepth\"/>")
            if (fPerspective)
                cooxml.append("<c:rAngAx val=\"1\"/>")
            if (pcDist.toInt() != 30)
                cooxml.append("<c:perspective val=\"$pcDist\"/>")
            cooxml.append("</c:view3D>")
            cooxml.append("\r\n")
            return cooxml

        }

    override fun init() {
        super.init()
        anRot = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        anElev = ByteTools.readShort(this.getByteAt(2).toInt(), this.getByteAt(3).toInt())
        pcDist = ByteTools.readShort(this.getByteAt(4).toInt(), this.getByteAt(5).toInt())
        pcHeight = ByteTools.readShort(this.getByteAt(6).toInt(), this.getByteAt(7).toInt())
        pcDepth = ByteTools.readShort(this.getByteAt(8).toInt(), this.getByteAt(9).toInt())
        pcGap = ByteTools.readShort(this.getByteAt(10).toInt(), this.getByteAt(11).toInt())
        grbit = ByteTools.readShort(this.getByteAt(12).toInt(), this.getByteAt(13).toInt())
        fPerspective = grbit and 0x1 == 0x1
        fCluster = grbit and 0x2 == 0x2
        f3dScaling = grbit and 0x4 == 0x4
        f2DWalls = grbit and 0x10 == 0x10
    }

    private fun updateRecord() {
        var b = ByteTools.shortToLEBytes(anRot)
        this.data[0] = b[0]
        this.data[1] = b[1]
        b = ByteTools.shortToLEBytes(anElev)
        this.data[2] = b[0]
        this.data[3] = b[1]
        b = ByteTools.shortToLEBytes(pcDist)
        this.data[4] = b[0]
        this.data[5] = b[1]
        b = ByteTools.shortToLEBytes(pcHeight)
        this.data[6] = b[0]
        this.data[7] = b[1]
        b = ByteTools.shortToLEBytes(pcDepth)
        this.data[8] = b[0]
        this.data[9] = b[1]
        b = ByteTools.shortToLEBytes(pcGap)
        this.data[10] = b[0]
        this.data[11] = b[1]
        b = ByteTools.shortToLEBytes(grbit)
        this.data[12] = b[0]
        this.data[13] = b[1]
    }

    /**
     * Handle setting options from XML in a generic manner
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        var bHandled = false
        if (op.equals("AnRot", ignoreCase = true)) { // specifies the clockwise rotation, in degrees, of the 3-D plot area around a vertical line through the center of the 3-D plot area.
            anRot = java.lang.Short.parseShort(`val`) // usually 20
            bHandled = true
        }
        if (op.equals("AnElev", ignoreCase = true)) { // signed integer that specifies the rotation, in degrees, of the 3-D plot area around a horizontal line through the center of the 3-D plot area
            anElev = java.lang.Short.parseShort(`val`) // usually 15
            bHandled = true
        }
        if (op.equals("PcDist", ignoreCase = true)) { // view angle for the 3-D plot area.
            pcDist = java.lang.Short.parseShort(`val`) // usually 30
            bHandled = true
        }
        if (op.equals("PcHeight", ignoreCase = true)) { // specifies the height of the 3-D plot area as a percentage of its width
            pcHeight = java.lang.Short.parseShort(`val`)
            bHandled = true
        }
        if (op.equals("PcDepth", ignoreCase = true)) { // specifies the depth of the 3-D plot area as a percentage of its width.
            pcDepth = java.lang.Short.parseShort(`val`) // usually 100
            bHandled = true
        }
        if (op.equals("PcGap", ignoreCase = true)) { // specifies the width of the gap between the series and the front and back edges of the 3-D plot area as a percentage of the data point depth divided by 2. If
            // fCluster is not set to 1 and chart group type is not a bar, this field also specifies the distance between adjacent series as a percentage of the data point depth.
            pcGap = java.lang.Short.parseShort(`val`) // usually 150
            bHandled = true
        }
        /*
         * 20070905 KSC: parse grbit options if
         * (op.equalsIgnoreCase("FormatOptions")) { grbit=
         * Short.parseShort(val); bHandled= true; }
         */
        if (op.equals("Perspective", ignoreCase = true)) { // specifies whether the 3-D plot area is rendered with a vanishing point.
            fPerspective = java.lang.Boolean.valueOf(`val`).booleanValue()
            grbit = ByteTools.updateGrBit(grbit, fPerspective, 0)
            bHandled = true
        }
        if (op.equals("Cluster", ignoreCase = true)) { // specifies whether data points are clustered together in a bar chart group
            fCluster = java.lang.Boolean.valueOf(`val`).booleanValue()
            grbit = ByteTools.updateGrBit(grbit, fCluster, 1)
            bHandled = true
        }
        if (op.equals("ThreeDScaling", ignoreCase = true)) { // specifies whether the height of the 3-D plot area is automatically determined
            f3dScaling = java.lang.Boolean.valueOf(`val`).booleanValue()
            grbit = ByteTools.updateGrBit(grbit, f3dScaling, 2)
            bHandled = true
        }
        if (op.equals("TwoDWalls", ignoreCase = true)) { // A bit that specifies whether the chart walls are rendered in 2-D
            f2DWalls = java.lang.Boolean.valueOf(`val`).booleanValue()
            grbit = ByteTools.updateGrBit(grbit, f2DWalls, 4)
            bHandled = true
        }
        if (bHandled)
            updateRecord()
        return bHandled
    }

    /**
     * return the desired option setting in string form
     */
    override fun getChartOption(op: String): String? {
        if (op.equals("AnRot", ignoreCase = true)) {
            return anRot.toString()
        }
        if (op.equals("AnElev", ignoreCase = true)) {
            return anElev.toString()
        }
        if (op.equals("PcDist", ignoreCase = true)) {
            return pcDist.toString()
        }
        if (op.equals("PcHeight", ignoreCase = true)) {
            return pcHeight.toString()
        }
        if (op.equals("PcDepth", ignoreCase = true)) {
            return pcDepth.toString()
        }
        if (op.equals("PcGap", ignoreCase = true)) {
            return pcGap.toString()
        }
        if (op.equals("Perspective", ignoreCase = true)) {
            return if (fPerspective) "1" else "0"
        }
        if (op.equals("Cluster", ignoreCase = true)) {
            return if (fCluster) "1" else "0"
        }
        if (op.equals("ThreeDScaling", ignoreCase = true)) {
            return if (f3dScaling) "1" else "0"
        }
        return if (op.equals("TwoDWalls", ignoreCase = true)) {
            if (f2DWalls) "1" else "0"
        } else ""
    }

    /**
     * sets the proper bit for chart type (PIE or not)
     *
     * @param isPieChart true if pie-type chart
     */
    fun setIsPie(isPieChart: Boolean) {
        if (isPieChart)
            grbit = grbit and 0x8 // bit is true if "not a pie"
        else
            grbit = grbit or 0x17
        updateRecord()
    }

    /**
     * sets the Rotation Angle (0 to 360 degrees), usually 0 for pie, 20 for
     * others
     *
     * @param rot
     */
    fun setAnRot(rot: Int) {
        anRot = rot.toShort()
        val b = ByteTools.shortToLEBytes(anRot)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    /**
     * sets the Elevation Angle (-90 to 90 degrees) (15 is default)
     *
     * @param elev
     */
    fun setAnElev(elev: Int) {
        anElev = elev.toShort()
        val b = ByteTools.shortToLEBytes(anElev)
        this.data[2] = b[0]
        this.data[3] = b[1]
    }

    /**
     * sets the Distance from eye to chart (0 to 100) (30 is default)
     *
     * @param elev
     */
    fun setPcDist(dist: Int) {
        pcDist = dist.toShort()
        val b = ByteTools.shortToLEBytes(pcDist)
        this.data[4] = b[0]
        this.data[5] = b[1]
    }

    /**
     * sets the Height of plot volume relative to width and depth (100 is
     * default)
     *
     * @param elev
     */
    fun setPcHeight(dist: Int) {
        pcHeight = dist.toShort()
        val b = ByteTools.shortToLEBytes(pcHeight)
        this.data[6] = b[0]
        this.data[7] = b[1]
    }

    /**
     * sets the Depth of points relative to width (100 is default)
     *
     * @param elev
     */
    fun setPcDepth(depth: Int) {
        pcDepth = depth.toShort()
        val b = ByteTools.shortToLEBytes(pcDepth)
        this.data[8] = b[0]
        this.data[9] = b[1]
    }

    /**
     * sets the Space between points (50 or 150 is default)
     *
     * @param gap
     */
    fun setPcGap(gap: Int) {
        pcGap = gap.toShort()
        val b = ByteTools.shortToLEBytes(pcGap)
        this.data[10] = b[0]
        this.data[11] = b[1]
    }

    fun getPcGap(): Int {
        return pcGap.toInt()
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -7501630910970731901L

        // 20070716 KSC: Need to create new records
        val prototype: XLSRecord?
            get() {
                val td = ThreeD()
                td.opcode = XLSConstants.THREED
                td.data = td.PROTOTYPE_BYTES
                td.init()
                return td
            }

        /**
         * parse shape OOXML element view3D into a ThreeD record
         *
         * @param xpp     XmlPullParser
         * @param lastTag element stack
         * @return spPr object
         */
        fun parseOOXML(xpp: XmlPullParser, lastTag: Stack<String>,
                       cht: OOXMLChart): OOXMLElement? {
            // threeD MUST NOT EXIST in a bar of pie, bubble, doughnut, filled
            // radar, pie of pie, radar, or scatter chart group.
            val td = cht.chartObject.getThreeDRec(true)
            try {
                var eventType = xpp.eventType
                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        var v: String? = null
                        try {
                            v = xpp.getAttributeValue(0)
                        } catch (/* XmlPullParser */e: Exception) {
                        }

                        if (v != null) {
                            if (tnm == "rotX") {
                                td!!.setAnElev(Integer.valueOf(v))
                            } else if (tnm == "rotY") {
                                td!!.setAnRot(Integer.valueOf(v))
                            } else if (tnm == "perspective") {
                                td!!.setPcDist(Integer.valueOf(v))
                            } else if (tnm == "depthPercent") {
                                td!!.setPcDepth(Integer.valueOf(v))
                            } else if (tnm == "rAngAx") {
                                //						if (v!=null)

                            }
                        }
                    } else if (eventType == XmlPullParser.END_TAG) {
                        val endTag = xpp.name
                        if (endTag == "view3D") {
                            lastTag.pop() // pop layout tag
                            break
                        }
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("ThreeD.parseOOXML: $e")
            }

            return null
        }
    }
}
