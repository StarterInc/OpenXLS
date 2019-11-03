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

import io.starter.formats.XLS.Font
import io.starter.formats.XLS.MSODrawing
import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools
import io.starter.toolkit.StringTool

import java.util.HashMap

/**
 * **Legend: Legend Type and Position (0x1015)**
 *
 *
 *
 *
 * 4	x		4		x position of upper-left corner -- MUST be ignored and the x1 field from the following Pos record MUST be used instead.
 * 8	y		4		y position of upper-left corner -- MUST be ignored and the y1 field from the following Pos record MUST be used instead.
 * 12	dx		4		width in SPRC -- MUST be ignored and the x2 field from the following Pos record MUST be used instead.
 * 16	dy		4		height in SPRC -- MUST be ignored and the y2 field from the following Pos record MUST be used instead.
 * 20			1       Undefined and MUST be ignored.
 * 21	wSpacing1		Spacing (0= close, 1= medium, 2= open) (0x1= 40 twips==4 pts)
 * 22	grbit	2		Option Flags
 *
 *
 * grbit Option Flags
 * bits		Mask
 * 0		01h		fAutoPostion	Automatic positioning (1= legend is docked)
 * 1		02h		fAutoSeries		Automatic series distribution
 * 2		04h		fAutoPosX		X positioning is automatic
 * 3		08h		fAutoPosY		Y positioning is automatic
 * 4		10h		fVert			1= vertical legend, 0= horizontal
 * 5		20h		fWasDataTable	1= chart contains data table
 *
 *
 * NOTES:
 * A SPRC is a unit of measurement that is 1/4000th of the height or width of the chart
 * If the field is being used to specify a width or horizontal distance, the SPRC is 1/4000th
 * of the width of the chart.  If the field is being used to specify a height or vertical
 * distance, the SPRC is 1/4000th of the height of the chart.
 *
 *
 * Sequence of records:
 * ATTACHEDLABEL = TextDisp Begin Pos [FontX] [AlRuns] AI [FRAME] [ObjectLink] [DataLabExtContents] [CrtLayout12] [TEXTPROPS] [CRTMLFRT] End
 * LD = Legend Begin Pos ATTACHEDLABEL [FRAME] [CrtLayout12] [TEXTPROPS] [CRTMLFRT] End
 */
class Legend : GenericChartObject(), ChartObject, ChartConstants {
    protected var x_defunct = -1
    protected var y_defunct = -1
    protected var dx_defunct = -1
    protected var dy_defunct = -1        // these vars are now defunct; see doc above + Pos/getSVG for coordinate info
    protected var /*wType= -1, */wSpacing: Byte = -1
    protected var grbit: Short = -1
    protected var fAutoPosition: Boolean = false
    protected var fAutoSeries: Boolean = false
    protected var fAutoPosX: Boolean = false
    protected var fAutoPosY: Boolean = false
    protected var fVert: Boolean = false
    protected var fWasDataTable: Boolean = false
    /**
     * return the coordinates of the legend box, relative to the chart
     *
     * @return int[] coordinates x, y, w, h [fh, legendpos]
     */
    var coords: IntArray? = null
        internal set

    private val PROTOTYPE_BYTES = byteArrayOf(-11, 13, 0, 0, -72, 3, 0, 0, -111, 1, 0, 0, -31, 4, 0, 0, 3, 1, 31, 0)

    /**
     * legend position:
     * controlled by CrtLayout12
     * 0= bottom, 1= corner, 2= top, 3= right, 4= left
     *
     * @return
     */
    /**
     * set Legend Positon:  one of:
     * 0= bottom, 1= corner, 2= top, 3= right, 4= left, 7= not docked
     *
     * @param pos
     */
    // default
    // default if not vert
    var legendPosition: Short
        get() {
            val crt = Chart.findRec(this.chartArr, CrtLayout12::class.java) as CrtLayout12
            if (crt != null)
                return crt.layout.toShort()
            return if (fVert || !fAutoPosition) RIGHT.toShort() else BOTTOM.toShort()
        }
        set(pos) {
            val crt = Chart.findRec(this.chartArr, CrtLayout12::class.java) as CrtLayout12
            if (crt != null)
                crt.layout = pos
        }

    /**
     * return the legend position in string (OOXML) form
     *
     * @return b, l, r, t, tr
     */
    val legendPositionString: String
        get() {
            val lpos = legendPosition.toInt()
            val pos = arrayOf("b", "tr", "t", "r", "l")
            return if (lpos >= 0 && lpos < pos.size) pos[lpos] else "r"
        }

    /**
     * retrieves the specific font for these legends, if set (null if not)
     *
     * @return
     */
    private val legendFont: Fontx?
        get() {
            val td = Chart.findRec(chartArr, TextDisp::class.java) as TextDisp
            return if (td != null) Chart.findRec(td.chartArr, Fontx::class.java) as Fontx else null
        }

    /**
     * tries to get the best match
     *
     * @return
     */
    // shouldn't get here ...
    // this actually doesn't get the actual font for the legend but can't find correct Fontx record!
    val fnt: Font?
        get() {
            try {
                val fx = legendFont
                val f = this.parentChart!!.workBook!!.getFont(fx!!.ifnt)
                return f ?: this.parentChart!!.defaultFont
            } catch (e: NullPointerException) {
                return this.parentChart!!.defaultFont
            }

        }


    /**
     * returns the Data Legend Box svg for this chart
     *
     * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
     * @return
     */
    internal var XOFFSET = 12

    override fun init() {
        super.init()
        val rkdata = this.data
        x_defunct = ByteTools.readInt(this.getBytesAt(0, 4)!!)
        y_defunct = ByteTools.readInt(this.getBytesAt(4, 4)!!)
        dx_defunct = ByteTools.readInt(this.getBytesAt(8, 4)!!)
        dy_defunct = ByteTools.readInt(this.getBytesAt(12, 4)!!)
        // unused:  wType= rkdata[16]; layout position is controled by CrtLayout12 record
        wSpacing = rkdata!![17]
        grbit = ByteTools.readShort(rkdata[18].toInt(), rkdata[19].toInt())
        parseGrbit()
    }

    /**
     * The following records and rules define the significant parts of a legend:
     *
     *
     * The Legend record specifies the layout of the legend and specifies if the legend is automatically positioned.
     * The Pos record, CrtLayout12 record, specify the position of the legend.
     * The sequences of records that conform to the ATTACHEDLABEL (TextDisp ->Pos [FontX] [AlRuns] AI [FRAME] [ObjectLink] [DataLabExtContents] [CrtLayout12] [TEXTPROPS] [CRTMLFRT] )
     * and TEXTPROPS (RichTextStream|TextPropStream) rules specify the default text formatting for the legend entries.
     * The Pos record of the attached label MUST be ignored. The ObjectLink record of the attached label MUST NOT exist.
     * A series can specify formatting exceptions for individual legend entries.
     * The sequence of records that conforms to the FRAME (Frame ->LineFormat AreaFormat [GELFRAME] [SHAPEPROPS]) rule specifies the fill and border formatting properties of the legend.
     */

    protected fun parseGrbit() {
        val grbytes = ByteTools.shortToLEBytes(grbit)
        fAutoPosition = grbytes[0] and 0x01 == 0x01
        fAutoSeries = grbytes[0] and 0x02 == 0x02
        fAutoPosX = grbytes[0] and 0x04 == 0x04
        fAutoPosY = grbytes[0] and 0x08 == 0x08
        fVert = grbytes[0] and 0x10 == 0x10
        fWasDataTable = grbytes[0] and 0x20 == 0x20
    }


    fun setIsDataTable(isDataTable: Boolean) {
        fWasDataTable = isDataTable
        grbit = ByteTools.updateGrBit(grbit, fWasDataTable, 5)
        updateRecord()
    }
    /* unused
	public void setwType(int type) {
		wType= (byte) type;
		updateRecord();
	}
	*/

    fun setVertical(isVertical: Boolean) {
        fVert = isVertical
        grbit = ByteTools.updateGrBit(grbit, fVert, 4)
        updateRecord()
    }

    private fun updateRecord() {
        System.arraycopy(ByteTools.cLongToLEBytes(x_defunct), 0, this.data!!, 0, 4)
        System.arraycopy(ByteTools.cLongToLEBytes(y_defunct), 0, this.data!!, 4, 4)
        System.arraycopy(ByteTools.cLongToLEBytes(dx_defunct), 0, this.data!!, 8, 4)
        System.arraycopy(ByteTools.cLongToLEBytes(dy_defunct), 0, this.data!!, 12, 4)
        // unused this.getData()[16]= wType;
        this.data[17] = wSpacing
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[18] = b[0]
        this.data[19] = b[1]
    }

    /**
     * returns true if this legend is surrounded by a box (the default)
     *
     * @return
     */
    fun hasBox(): Boolean {
        val f = Chart.findRec(chartArr, Frame::class.java) as Frame
        return f?.hasBox() ?: false
        //    	return true; // the default
    }

    fun addBox() {
        var f = Chart.findRec(chartArr, Frame::class.java) as Frame
        if (f == null) {
            f = Frame.prototype as Frame?
            f.addBox(0, -1, -1)
            this.chartArr.add(f)
        }
    }

    /**
     * sets or turns off auto positioning
     * [BugTracker 2844]
     *
     * @param auto
     */
    fun setAutoPosition(auto: Boolean) {
        if (auto && !fAutoPosition) {
            // if setting to autosize/position and it wasn't currently set as so,
            // check Pos and Frame records (if present) as they also controls automatic positioning ((:
            if (this.chartArr.size > 0) {
                try {
                    val p = this.chartArr[0] as Pos
                    p.setAutosizeLegend()
                    val f = Chart.findRec(chartArr, Frame::class.java) as Frame    // find the first one
                    f?.setAutosize()
                } catch (e: Exception) {
                }

            }
        }
        fAutoPosition = auto
        fAutoSeries = auto
        fAutoPosX = auto
        fAutoPosY = auto
        //if (wType==3 || wType==4 && auto)
        //fVert= true;
        grbit = ByteTools.updateGrBit(grbit, fAutoPosition, 0)
        grbit = ByteTools.updateGrBit(grbit, fAutoSeries, 1)
        grbit = ByteTools.updateGrBit(grbit, fAutoPosX, 2)
        grbit = ByteTools.updateGrBit(grbit, fAutoPosY, 3)
        grbit = ByteTools.updateGrBit(grbit, fVert, 4)
        this.updateRecord()
    }

    /**
     * a rough estimate of expanding legend dimensions of
     * 1 normal entry
     */
    fun incrementHeight(h: Float) {
        val p = Chart.findRec(this.chartArr, Pos::class.java) as Pos
        val coords = p.legendCoords    // x, y, w, h, fh, legendpos
        val f = this.fnt
        var fh = 10    // default
        if (f != null)
            fh = (f.fontHeightInPoints * 1.2).toInt()    // a little padding
        if (coords != null) {
            val z = coords[1] - Math.ceil(Pos.convertToSPRC((fh / 2).toFloat(), 0f, h).toDouble()).toInt()
            p.setY(z)
        }
    }

    /**
     * called up change of legend text to adjust width of the legend bounding box
     *
     * @param chartMetrics
     * @param chartType
     * @param legends      String[] text of legends (containing new legend text)
     */
    fun adjustWidth(chartMetrics: HashMap<String, Double>, chartType: Int, legends: Array<String>) {
        val p = Chart.findRec(this.chartArr, Pos::class.java) as Pos
        val coords = p.legendCoords    // x, y, w, h, fh, legendpos
        if (coords != null) {
            val f = this.fnt
            // legend position LEFT and RIGHT display each legend on a separate line (fVert==true)
            // TOP and BOTTOM are displayed horizontally with symbols and spacing between entries (fVert==false)
            val position = this.legendPosition.toInt()
            val cw = chartMetrics["canvasw"].toFloat()
            var x = (Math.ceil(Pos.convertFromSPRC(coords[0].toFloat(), cw, 0f).toDouble()).toInt() - 3).toFloat()
            val w = chartMetrics["w"].toFloat()

            // calculate how much width the legends take up -- algorithm works well for about 80% of the cases ...
            var legendsWidth = 0.0
            val jf = java.awt.Font(f!!.fontName, f.fontWeight, f.fontHeightInPoints.toInt())
            val extras = if (chartType == ChartConstants.LINECHART || chartType == ChartConstants.RADARCHART) 15 else 5    // pad for legend symbols, etc	-
            for (i in legends.indices) {
                if (fVert)
                    legendsWidth = Math.max(legendsWidth, StringTool.getApproximateStringWidth(jf, " " + legends[i] + " "))
                else
                    legendsWidth += StringTool.getApproximateStringWidth(jf, " " + legends[i] + " ") + extras
            }
            if (!fVert)
                legendsWidth -= StringTool.getApproximateStringWidth(jf, " ")    // decrement one space
            else
                legendsWidth += StringTool.getApproximateStringWidth(jf, " ")

            //	    io.starter.toolkit.Logger.log(this.getParentChart().toString() + String.format(": legend box x: %.1f legend box w: %.0f chart x: %.1f w: %.1f cw: %.1f font size: %.0f L.W: %.1f Auto? %b Vertical? %b",
            //		    x, (float)coords[2], chartMetrics.get("x"), w, cw, (float) jf.getSize(), legendsWidth, fAutoPosition, fVert));
            p.setLegendW(legendsWidth.toInt())
            if (x + legendsWidth > cw || position == Legend.RIGHT || position == Legend.CORNER) {
                x = (cw - (legendsWidth + 5)).toFloat()
                if (x < 0) x = 0f
                val z = Math.ceil(Pos.convertToSPRC(x, cw, 0f).toDouble()).toInt()
                p.setX(z)
            }


/*
	    if (position==Legend.RIGHT || position==Legend.CORNER) {	// usual case
		totalWidth+= extras;
		if ((x+totalWidth) > cw) { // if legends will extend over edge of chart, must adjust either x or chart canvas width to accommodate legends
		    x= cw-totalWidth;
		}
	    }
//	    p.setLegendW((int)totalWidth);




	   else if (position==Legend.LEFT) {
		totalWidth+= extras;
		if (totalwidth > x)
		    double originalDist= x-w; 	// original distance between legend and edge of plot area
		    if ((w + totalWidth + originalDist) < cw)
			;

		    if (totalWidth > (cw-w)) {	// can fit in space between plot area and
			//KSC: TESTINGs
			//io.starter.toolkit.Logger.log("Original Lcoords:  w" + w + " cw:" + cw + " coords:" + Arrays.toString(retcoords));
			float newX= (int)Math.ceil(cw-w);
			if (originalDist > 0 && (w+x) > newX)
			    chartMetrics.put("w", newX-originalDist);
			//io.starter.toolkit.Logger.log("After Lcoords:  w" + w + " cw:" + cw + " coords:" + Arrays.toString(retcoords));
		    }
		}
	    } else if (totalWidth > cw) {	// legends are displayed horizontally, can't fit in canvas width
// TODO: must extend cw ?? wrap ???
		/*	float overage= ((x+w+10)-cw);
		if (overage > 0)
		    chartMetrics.put("canvasw", chartMetrics.get("canvasw")+overage);* /
	    }*/
        }
}

/**
 * reset initial position of legend to accommodate nLines of legend text
 */
     fun resetPos(y:Double, h:Double, ch:Double, nLines:Int) {
val p = Chart.findRec(this.chartArr, Pos::class.java) as Pos?
 /*		apparently just setting to 1/2 h works well!!
 * 		Font f= this.getFnt();
		int fh= 10;	// default
		if (f!=null)
			fh=(int)(f.getFontHeightInPoints());
		fh=(int)(f.getFontHeightInPoints()*1.2);	// a little padding*/
        val z = Math.ceil(Pos.convertToSPRC((ch / 2/*+((fh*nLines)/2)*/).toFloat(), 0f, ch.toFloat()).toDouble()).toInt()
p!!.setY(z)

}

/**
 * return the coordinates of the legend box in pixels
 * <br></br>An approximation at this point
 *
 * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
 * @param fh           -- font height in points
 * @return int[4] x, y, w, h
 */
     fun getCoords(charttype:Int, chartMetrics:HashMap<String, Double>, legends:Array<String>, f:java.awt.Font):IntArray {
 // calcs are not 100% ****
        // space between legend entries = 40 twips = 1 twip equals one-twentieth of a printer's point
        val p = Chart.findRec(this.chartArr, Pos::class.java) as Pos?
var coords = p!!.legendCoords
val retcoords = IntArray(6)
val fh = f.getSize()    //*1.2);	// a little padding
retcoords[4] = f.getSize()    // store font height
retcoords[5] = this.legendPosition.toInt()    // store legend position

var canMoveCW = false
if (coords != null)
{
retcoords[0] = Math.ceil(Pos.convertFromSPRC(coords!![0].toFloat(), chartMetrics.get("canvasw").toFloat(), 0f).toDouble()).toInt() - 3
retcoords[1] = Math.ceil(Pos.convertFromSPRC(coords!![1].toFloat(), 0f, chartMetrics.get("canvash").toFloat()).toDouble()).toInt()
}
else
{ // happens upon OOXML
retcoords[0] = (chartMetrics.get("w") + chartMetrics.get("x") + 20.0).toInt()    // start just after right side of plot
retcoords[1] = chartMetrics.get("y").toInt() + (chartMetrics.get("h") / 4).toInt()
coords = IntArray(4)
canMoveCW = true
}
if (coords!![2] != 0)
{
retcoords[2] = (coords!![2] * MSODrawing.PIXELCONVERSION).toInt()
retcoords[2] += 3    // pad slightly
}
else
{
var len = 0.0
for (i in legends.indices)
{
len = Math.max(len, StringTool.getApproximateStringWidth(f, legends[i]))
}
retcoords[2] = Math.ceil(len).toInt()
retcoords[2] += 15 + (if (charttype == ChartConstants.LINECHART || charttype == ChartConstants.RADARCHART) 25 else 5)    // pad for legend symbols, etc	-
 // if now legend box extends over edge reduce plot area width, not canvas width ... EXCEPT for OOXML; in those cases, extend CW
            if (!canMoveCW && ((retcoords[0] + retcoords[2]) > chartMetrics.get("canvasw").toFloat()))
{
val cw = chartMetrics.get("canvasw")
val w = chartMetrics.get("w")
 // KSC: TESTING
 //io.starter.toolkit.Logger.log("Original Lcoords:  w" + w + " cw:" + cw + " coords:" + Arrays.toString(retcoords));
                val ldist = retcoords[0] - w    // original distance between legend and edge of plot area
retcoords[0] = Math.ceil(cw - retcoords[2]).toInt()
if (ldist > 0 && (chartMetrics.get("w") + chartMetrics.get("x")) > retcoords[0])
chartMetrics.put("w", retcoords[0] - ldist)
 //io.starter.toolkit.Logger.log("After Lcoords:  w" + w + " cw:" + cw + " coords:" + Arrays.toString(retcoords));
            }

}
if (canMoveCW)
{
val overage = ((retcoords[0] + retcoords[2] + 10) - (chartMetrics.get("canvasw").toFloat()))
if (overage > 0)
chartMetrics.put("canvasw", chartMetrics.get("canvasw") + overage)
}
if (coords!![3] != 0)
retcoords[3] = (coords!![3] * MSODrawing.PIXELCONVERSION).toInt()
else
retcoords[3] = ((legends.size + 2) * (fh + 2))
return retcoords
}


 fun getMetrics(chartMetrics:HashMap<String, Double>, chartType:Int, s:ChartSeries) {
val legends = s.getLegends()
if (legends == null || legends!!.size == 0) return

val f = fnt
if (f != null)
{
coords = this.getCoords(chartType, chartMetrics, legends, java.awt.Font(f!!.getFontName(), f!!.fontWeight, f!!.fontHeightInPoints.toInt()))
}
else
{    // can't find any font ... shouldn't really happen ...?
coords = this.getCoords(chartType, chartMetrics, legends, java.awt.Font("Arial", 400, 10))
}
}

 fun getSVG(chartMetrics:HashMap<String, Double>, chartobj:ChartType, s:ChartSeries):String {
val svg = StringBuffer()
 // position information fro Pos record:
        /**
 * legend			MDCHART						MDABS						The values x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
 * relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 specify the
 * width and height of the legend, in points.
 * legend			MDCHART						MDPARENT					The values of x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
 * relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 MUST be ignored.
 * The size of the legend is determined by the application.
 * legend			MDKTH						MDPARENT					The values of x1, y1, x2 and y2 MUST be ignored. The legend is located inside a data table.
 *
 */

        val legends = s.getLegends()
val seriescolors = s.seriesBarColors
if (legends == null || legends!!.size == 0) return ""

if (coords == null)
this.getMetrics(chartMetrics, chartobj.chartType, s)
val font:String    // font svg
val fh:Int    // font height
val f = fnt
if (f != null)
{
font = f!!.svg
font = "' " + /*"' vertical-align='bottom' " +*/ font
fh = Math.ceil(f!!.fontHeightInPoints).toInt()
}
else
{    // can't find any font ... shouldn't really happen ...?
font = "' " + "font-family='Arial' font-size='9pt'"
fh = 10
}
 // get legend info in order to get dimensions
        val YOFFSET = coords!![3] / (legends!!.size)

var x = coords!![0]
var y = coords!![1]
val boxw = coords!![2]
val boxh = coords!![3]

svg.append("<g>\r\n")
if (this.hasBox())
{
svg.append(("<rect x='" + x + "' y='" + y +
"' width='" + boxw + "' height='" + boxh +
"' fill='#FFFFFF' fill-opacity='1' stroke='black' stroke-opacity='1' stroke-width='1' " +
"stroke-linecap='butt' stroke-linejoint='miter' stroke-miterlimit='4' fill-rull='evenodd'" +
"/>"))
x += 5    // start of labels (offset from box)
y += (YOFFSET / 3)
}
if (chartobj.chartType == ChartConstants.BARCHART)
{    // same as below except order is reversed
 // draw a little box in appropriate color
            val h = 8    // box size
for (i in legends!!.indices.reversed())
{
svg.append(("<rect x='" + x + "' y='" + y + "' width='" + h + "' height='" + h + "' fill='" + seriescolors!![i] +
"' fill-opacity='1' stroke='black' stroke-opacity='1' stroke-width='1' stroke-linecap='butt' stroke-linejoint='miter' stroke-miterlimit='4' fill-rull='evenodd'" +
"/>"))
svg.append("<text " + GenericChartObject.getScript("legend" + i) + " x='" + (x + XOFFSET) + "' y='" + (y + fh) + font + ">" + legends!![i] + "</text>")
y += YOFFSET
}
}
else if (!(chartobj.chartType == ChartConstants.LINECHART || chartobj.chartType == ChartConstants.SCATTERCHART || chartobj.chartType == ChartConstants.RADARCHART || chartobj.chartType == ChartConstants.BUBBLECHART))
{
 // draw a little box in appropriate color
            val h = 8    // box size
for (i in legends!!.indices)
{
svg.append(("<rect x='" + x + "' y='" + y + "' width='" + h + "' height='" + h + "' fill='" + seriescolors!![i] +
"' fill-opacity='1' stroke='black' stroke-opacity='1' stroke-width='1' stroke-linecap='butt' stroke-linejoint='miter' stroke-miterlimit='4' fill-rull='evenodd'" +
"/>"))
svg.append("<text " + GenericChartObject.getScript("legend" + i) + " x='" + (x + XOFFSET) + "' y='" + (y + fh) + font + ">" + legends!![i] + "</text>")
y += YOFFSET
}
}
else if (chartobj.chartType == ChartConstants.BUBBLECHART)
{
 // little circles
            for (i in legends!!.indices.reversed())
{
svg.append("<circle cx='" + (x + 3) + "' cy='" + (y + 6) + "' r='5' stroke='black' stroke-width='1' fill='" + seriescolors!![i] + "'/>")
svg.append("<text " + GenericChartObject.getScript("legend" + i) + " x='" + (x + XOFFSET) + "' y='" + (y + fh) + font + ">" + legends!![i] + "</text>")
y += YOFFSET
}
}
else
{    // line-type charts/scatter charts/radar charts
 // lines w/markers if nec.
            val markers = chartobj.markerFormats    // markers, if any
var haslines = true
if (chartobj.chartType == ChartConstants.SCATTERCHART)
haslines = chartobj.hasLines        // lines, if any, on scatter plot
svg.append(MarkerFormat.markerSVGDefs) // initial SVG necessary for markers
var w = 25        // w of line + markers
if (!haslines)
w = 10
XOFFSET = w + 4
y += 2    // a bit more padding
for (i in legends!!.indices)
{
if (haslines)
svg.append(("<line x1='" + x + "' y1='" + (y + fh / 2) + "' x2='" + (x + w) + "' y2='" + (y + fh / 2) +
"' stroke='" + seriescolors!![i] + "' stroke-width='2'/>"))
if (markers[i] > 0)
svg.append(MarkerFormat.getMarkerSVG((x + w / 2 - 5).toDouble(), (y + fh / 2).toDouble(), seriescolors!![i], markers[i]))

svg.append("<text " + GenericChartObject.getScript("legend" + i) + " x='" + (x + XOFFSET) + "' y='" + (y + fh) + font + ">" + legends!![i] + "</text>")
y += YOFFSET
}
}
svg.append("</g>\r\n")
return svg.toString()
}

companion object {
/**
 * serialVersionUID
 */
    private val serialVersionUID = -4041111720696805018L
 val BOTTOM = 0
 val CORNER = 1
 val TOP = 2
 val RIGHT = 3
 val LEFT = 4
 val NOT_DOCKED = 7

 fun createDefaultLegend(book:io.starter.formats.XLS.WorkBook):Legend {
val l = Legend.prototype as Legend?
val p = Pos.getPrototype(Pos.TYPE_LEGEND) as Pos
l!!.chartArr.add(p)
val td = TextDisp.getPrototype(ObjectLink.TYPE_TITLE, "", book) as TextDisp
l!!.chartArr.add(td)
return l
}

 val prototype:XLSRecord?
get() {
val l = Legend()
l.opcode = XLSConstants.LEGEND
l.setData(l.PROTOTYPE_BYTES)
l.init()
return l
}
}
}
