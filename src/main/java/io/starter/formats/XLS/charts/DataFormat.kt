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

import io.starter.formats.XLS.XLSRecord
import io.starter.toolkit.ByteTools

/**
 * **DataFormat: Series and Data Point Numbers(0x1006)**
 *
 *
 * 4	xi		2	the zero-based index of the data point within the series specified by yi. (FFFFh means entire series)
 * 6	yi		2	the zero-based index of a Series record
 * 8	iss		2	An unsigned integer that specifies properties of the data series, trendline or error bar, depending on the type of records in sequence of records:
 *
 *
 * If does not contain a SerAuxTrend or SerAuxErrBar record, then this field specifies the plot order of the data series.
 * If the series order was changed, this field can be different from yi. MUST be less than or equal to the number of series in the chart.
 * MUST be unique among iss values for all instances of this record contained in the SERIESFORMAT rule that does not contain a SerAuxTrend or SerAuxErrBar record.
 *
 *
 * If the SERIESFORMAT rule contains a SerAuxTrend record on the chart group, then this field specifies the trendline number for the series.
 *
 *
 * If the SERIESFORMAT rule contains a SerAuxErrBar record on the chart group, then this field specifies a zero-based index into a Series record in the
 * collection of Series records in the current chart sheet substream for which the error bar applies to.
 *
 *
 * 10	grbit	2	flags (0?)		ignored
 *
 *
 *
 *
 * ORDER OF SUB-RECS:
 * [Chart3DBarShape]
 * [LineFormat, AreaFormat, PieFormat]		== lines, fill
 * [SerFormat]			== smoothed lines ...
 * [GelFrame]
 * [MarkerFormat]
 * [AttachedLabel]		== data labels
 * [ShapeProps]
 * [CtrlMltFrt]
 */
class DataFormat : GenericChartObject(), ChartObject {
    var seriesIndex: Short = 0
        private set
    var pointNumber: Short = 0
        private set
    var seriesNumber: Short = 0
        private set

    private val PROTOTYPE_BYTES = byteArrayOf(-1, -1, 0, 0, 0, 0, 0, 0)

    /**
     * return XLSRecord data (2 bytes), which controls bar shape
     */
    val shape: Short
        get() {
            val cs = Chart.findRec(this.chartArr, Chart3DBarShape::class.java) as Chart3DBarShape
            return cs.shape
        }

    /**
     * returns true if this parent chart has smoothed lines (Line, Scatter, Radar charts)
     *
     * @return
     */
    val smoothedLines: Boolean
        get() {
            val sf = Chart.findRec(this.chartArr, Serfmt::class.java) as Serfmt
            return sf?.smoothLine ?: false
        }

    /**
     * returns true if this parent chart has lines (Line, Scatter, Radar charts)
     *
     * @return
     */
    val hasLines: Boolean
        get() {
            val l = Chart.findRec(this.chartArr, LineFormat::class.java) as LineFormat
            return if (l != null) l.lineStyle != LineFormat.NONE else false
        }

    /**
     * returns true if this paernt chart has 3D bubbles (Bubble chart only)
     *
     * @return
     */
    /**
     * sets 3D Bubbles (Bubble Chart only)
     *
     * @param b true if has 3d Bubbles
     */
    var has3DBubbles: Boolean
        get() {
            val sf = Chart.findRec(this.chartArr, Serfmt::class.java) as Serfmt
            if (sf != null)
                return sf.3DBubbles
            return false
        }
        set(b) {
            var sf = Chart.findRec(this.chartArr, Serfmt::class.java) as Serfmt
            if (sf == null) {
                if (b) {
                    sf = Serfmt.prototype as Serfmt?
                    val i = Chart.findRecPosition(this.chartArr, PieFormat::class.java)
                    sf.parentChart = this.parentChart
                    this.chartArr.add(i + 1, sf)
                    sf.setHas3dBubbles(true)
                }
            } else {
                sf.setHas3dBubbles(b)
            }
        }


    /**
     * return if data markers are displayed with a shadow on bubble,
     * scatter, radar, stock, and line chart groups.
     */
    /**
     * data markers are displayed with a shadow on bubble,
     * scatter, radar, stock, and line chart groups.
     *
     * @param b
     */
    var hasShadow: Boolean
        get() {
            val sf = Chart.findRec(this.chartArr, Serfmt::class.java) as Serfmt
            return sf?.hasShadow() ?: false
        }
        set(b) {
            var sf = Chart.findRec(this.chartArr, Serfmt::class.java) as Serfmt
            if (sf == null) {
                if (b) {
                    sf = Serfmt.prototype as Serfmt?
                    val i = Chart.findRecPosition(this.chartArr, PieFormat::class.java)
                    sf.parentChart = this.parentChart
                    this.chartArr.add(i + 1, sf)
                    sf.setHasShadow(true)
                }
            } else {
                sf.setHasShadow(b)
            }
        }

    /**
     * return percentage=distance of pie slice from center of pie as %
     */
    /**
     * percentage=distance of pie slice from center of pie as %
     *
     * @param p
     */
    var percentage: Int
        get() {
            val pf = Chart.findRec(this.chartArr, PieFormat::class.java) as PieFormat
            return pf?.getPercentage()?.toInt() ?: 0

        }
        set(p) {
            var pf = Chart.findRec(this.chartArr, PieFormat::class.java) as PieFormat
            if (pf == null) {
                setHasLines(LineFormat.NONE)
                pf = Chart.findRec(this.chartArr, PieFormat::class.java) as PieFormat
            }
            pf.setPercentage(p.toShort())
        }

    /**
     * returns true if has data labels
     *
     * @return
     */
    val hasDataLabels: Boolean
        get() = this.getAttachedLabelRec(false) != null

    /**
     * return a string of ALL the label options chosen.  One or more of:
     *  * Value
     *  * ValuePerecentage
     *  * CategoryPercentage	// Pie only
     *  * CategoryLabel
     *  * BubbleLabel
     *  * SeriesLabel
     *
     * @return string true or false
     */
    val dataLabelType: String?
        get() {
            val al = this.getAttachedLabelRec(false)
            return al?.type
        }

    /**
     * return the data label int or 0 if no data labels chosen
     *
     * @return
     */
    val dataLabelTypeInt: Int
        get() {
            val al = this.getAttachedLabelRec(false)
            return al?.typeInt ?: 0
        }

    /**
     * returns type of marker, if any <br></br>
     * 0 = no marker <br></br>
     * 1 = square <br></br>
     * 2 = diamond <br></br>
     * 3 = triangle <br></br>
     * 4 = X <br></br>
     * 5 = star <br></br>
     * 6 = Dow-Jones <br></br>
     * 7 = standard deviation <br></br>
     * 8 = circle <br></br>
     * 9 = plus sign
     *
     * @return
     */
    /**
     * 0 = no marker <br></br>
     * 1 = square <br></br>
     * 2 = diamond <br></br>
     * 3 = triangle <br></br>
     * 4 = X <br></br>
     * 5 = star <br></br>
     * 6 = Dow-Jones <br></br>
     * 7 = standard deviation <br></br>
     * 8 = circle <br></br>
     * 9 = plus sign
     */
    // default actually looks like: 2, 1, 5, 4 ...
    // these records come in a set
    // creates default bar shape==0, 0
    // shouldn't get here but it goes
    // dunno, add to end
    var markerFormat: Int
        get() {
            val mf = Chart.findRec(this.chartArr, MarkerFormat::class.java) as MarkerFormat
            return mf?.markerFormat ?: 0

        }
        set(marker) {
            var mf = Chart.findRec(this.chartArr, MarkerFormat::class.java) as MarkerFormat
            if (mf == null) {
                if (chartArr.isEmpty()) {
                    val cs = Chart3DBarShape()
                    cs.opcode = XLSConstants.CHART3DBARSHAPE
                    this.addChartRecord(cs)
                    val lf = LineFormat.prototype as LineFormat?
                    lf!!.parentChart = this.parentChart
                    lf.lineStyle = 5
                    chartArr.add(lf)
                    val af = AreaFormat.prototype as AreaFormat?
                    af!!.parentChart = this.parentChart
                    chartArr.add(af)
                    val pf = PieFormat.prototype as PieFormat?
                    pf!!.parentChart = this.parentChart
                    chartArr.add(pf)
                    mf = MarkerFormat.prototype as MarkerFormat?
                    mf.parentChart = this.parentChart
                    chartArr.add(mf)
                } else {
                    mf = MarkerFormat.prototype as MarkerFormat?
                    mf.parentChart = this.parentChart
                    val z = Chart.findRecPosition(chartArr, PieFormat::class.java)
                    if (z > -1)
                        chartArr.add(z + 1, mf)
                    else
                        chartArr.add(mf)
                }
            }
            mf.setMarkerFormat(marker.toShort())
        }


    /*
	public int[] getDataLabelsPIE(int defaultdl) {
		 i++;
		      	//ArrayList dls= new ArrayList();
		int[] dls= new int[chartArr.size()-i-1];
		int j= 0;
		for (; i < chartArr.size(); i++) {
			if (chartArr.get(i) instanceof DataFormat) {
				DataFormat df= ((DataFormat) chartArr.get(i));
				AttachedLabel al= (AttachedLabel) Chart.findRec(df.chartArr, AttachedLabel.class);
				if (al!=null)
					dls[j]= al.getTypeInt();
				dls[j++]|=defaultdl;
			}
		}
		return dls;
	}
*/

    /**
     * get the bg color identified by this DataFormat
     * <br></br>Usually part of a series group of records
     *
     * @return
     */
    val bgColor: String?
        get() = Frame.getBgColor(chartArr)

    /**
     * returns (creates if necessary) the area format for this series- controls bar/series colors
     *
     * @return AreaFormat record
     */
    private// NOTE: below list of records is what has been observed in Excel 2003 chart files -
    // unsure if need marker format always ?
    val areaFormat: AreaFormat?
        get() {
            var af = Chart.findRec(this.chartArr, AreaFormat::class.java) as AreaFormat
            if (af == null) {
                af = AreaFormat.getPrototype(0) as AreaFormat
                this.addChartRecord(LineFormat.prototype)
                this.addChartRecord(af)
                this.addChartRecord(PieFormat.prototype)
                this.addChartRecord(MarkerFormat.prototype)
            }
            return af
        }

    override fun init() {
        super.init()
        val rkdata = this.data
        pointNumber = ByteTools.readShort(rkdata!![0].toInt(), rkdata[1].toInt())
        seriesIndex = ByteTools.readShort(rkdata[2].toInt(), rkdata[3].toInt())
        seriesNumber = ByteTools.readUnsignedShort(rkdata[4], rkdata[5]).toShort()
    }


    fun initNew() {
        this.opcode = XLSConstants.DATAFORMAT
        this.data = PROTOTYPE_BYTES
        this.init()
        val cs = Chart3DBarShape()
        cs.opcode = XLSConstants.CHART3DBARSHAPE    // creates default bar shape==0, 0
        this.addChartRecord(cs)
    }

    fun setPointNumber(idx: Int) {
        pointNumber = idx.toShort()
        val rkdata = this.data
        val num = ByteTools.shortToLEBytes(idx.toShort())
        rkdata[0] = num[0]
        rkdata[1] = num[1]
        this.data = rkdata
    }

    /**
     * Set the series index
     */
    fun setSeriesIndex(idx: Int) {
        seriesIndex = idx.toShort()
        val rkdata = this.data
        val num = ByteTools.shortToLEBytes(idx.toShort())
        rkdata[2] = num[0]
        rkdata[3] = num[1]
        this.data = rkdata
    }

    fun setSeriesNumber(idx: Int) {
        seriesNumber = idx.toShort()
        val rkdata = this.data
        val num = ByteTools.shortToLEBytes(idx.toShort())
        rkdata[4] = num[0]
        rkdata[5] = num[1]
        this.data = rkdata
    }

    private fun getAttachedLabelRec(bCreate: Boolean): AttachedLabel? {
        var al: AttachedLabel? = null
        al = Chart.findRec(this.chartArr, AttachedLabel::class.java) as AttachedLabel
        if (al == null && bCreate) { // basic options are handled via AttachedLabel rec
            al = AttachedLabel.prototype as AttachedLabel?
            val z = Chart.findRecPosition(this.chartArr, MarkerFormat::class.java)
            if (z > 0)
                chartArr.add(z + 1, al)
            else
                this.addChartRecord(al)
        }
        return al
    }

    private fun getAreaFormatRec(bCreate: Boolean): AreaFormat? {
        var af = Chart.findRec(this.chartArr, AreaFormat::class.java) as AreaFormat
        if (af == null) {
            af = AreaFormat.getPrototype(0) as AreaFormat
            this.addChartRecord(LineFormat.prototype)
            this.addChartRecord(af)
            this.addChartRecord(PieFormat.prototype)
            this.addChartRecord(MarkerFormat.prototype)
        }
        return af
    }

    /**
     * set the shape bit of the associated XLSRecord
     *
     * @param shape
     */
    fun setShape(shape: Int) {
        val cs = Chart.findRec(this.chartArr, Chart3DBarShape::class.java) as Chart3DBarShape
        cs.shape = shape.toShort()
    }

    /**
     * set smooth lines setting (applicable for line, scatter charts)
     *
     * @param smooth
     */
    fun setSmoothLines(smooth: Boolean) {
        var sf = Chart.findRec(this.chartArr, Serfmt::class.java) as Serfmt
        if (sf == null) {
            if (smooth) {
                setHasLines()
                sf = Serfmt.prototype as Serfmt?
                val i = Chart.findRecPosition(this.chartArr, PieFormat::class.java)
                this.chartArr.add(i + 1, sf)
                sf.setSmoothedLine(true)
            }
        } else {
            sf.setSmoothedLine(smooth)
        }
    }

    /**
     * sets this chart to have lines (line chart, radar, scatter ...) of the specific line style
     * <br></br>Style of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
     */
    @JvmOverloads
    fun setHasLines(lineStyle: Int = 0) {
        val l = Chart.findRec(this.chartArr, LineFormat::class.java) as LineFormat
        if (l == null) {    // these come as a group - assume none or only has Chart3DBarShape ...
            var z = Chart.findRecPosition(chartArr, Chart3DBarShape::class.java) + 1
            val lf = LineFormat.getPrototype(lineStyle, -1) as LineFormat
            chartArr.add(z++, lf)
            val af = AreaFormat.prototype as AreaFormat?
            af!!.parentChart = parentChart
            chartArr.add(z++, af)
            val pf = PieFormat.prototype as PieFormat?
            pf!!.parentChart = parentChart
            chartArr.add(z++, pf)
            val mf = MarkerFormat.prototype as MarkerFormat?
            mf!!.parentChart = parentChart
            chartArr.add(z++, mf)

        } else {
            l.lineStyle = lineStyle
        }
    }

    /**
     * sets the data labels to the desired type:
     *  * "ShowValueLabel"
     *  * "ShowValueAsPercent"
     *  * "ShowLabelAsPercent"
     *  * "ShowLabel"
     *  * "ShowSeriesName"
     *  * "ShowBubbleLabel"
     *
     * @param type
     */
    fun setDataLabels(type: String) {
        val al = this.getAttachedLabelRec(true)
        al!!.setType(type, "1")
    }

    /**
     * return if has the specified data label
     *  * "ShowValueLabel"
     *  * "ShowValueAsPercent"
     *  * "ShowLabelAsPercent"
     *  * "ShowLabel"
     *  * "ShowSeriesName"
     *  * "ShowBubbleLabel"
     *
     * @param type String option
     * @return string true or false
     */
    fun getDataLabelType(type: String): String? {
        val al = this.getAttachedLabelRec(false)
        return al?.getType(type)
    }

    /**
     * sets the data labels for the entire chart (as opposed to a specific series/data point).
     * A combination of:
     *  * SHOWVALUE= 0x1;
     *  * SHOWVALUEPERCENT= 0x2;
     *  * SHOWCATEGORYPERCENT= 0x4;
     *  * SHOWCATEGORYLABEL= 0x10;
     *  * SHOWBUBBLELABEL= 0x20;
     *  * SHOWSERIESLABEL= 0x40;
     */
    fun setHasDataLabels(dl: Int) {
        val al = this.getAttachedLabelRec(true)
        al!!.setType(dl.toShort())
    }

    // 0x893 CtlMltFrt
    // [-98, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 0, 0, 0, 0, 13, 19, 0, 6, -128, 34, 0, 35, 0, -67, 64, 0, 0]

    /**
     * sets the color identified by this DataFormat in the group of records
     * belonging to the parent series
     *
     * @param clr color int
     */
    fun setSeriesColor(clr: Int) {
        val af = getAreaFormatRec(true)
        // Finally got the AreaFormat record that governs the color for this series
        af!!.seticvFore(clr)
    }

    /**
     * sets the color identified by this DataFormat in the group of records
     * belonging to the parent series
     *
     * @param clr Hex color string
     */
    fun setSeriesColor(clr: String) {
        val af = getAreaFormatRec(true)
        // Finally got the AreaFormat record that governs the color for this series
        af!!.seticvFore(clr)
    }

    /**
     * sets the color of the desired pie slice
     *
     * @param clr   color int
     * @param slice 0-based pie slice number
     */
    fun setPieSliceColor(clr: String, slice: Int) {
        val af = getAreaFormatPie(slice)
        af.seticvFore(clr)
    }

    /**
     * sets the color of the desired pie slice
     *
     * @param clr   color int
     * @param slice 0-based pie slice number
     */
    fun setPieSliceColor(clr: Int, slice: Int) {
        val af = getAreaFormatPie(slice)
        af.seticvFore(clr)
    }

    /**
     * returns (creates if necessary) the area format for the desired pie slice (pie charts only)
     *
     * @param slice int 0-based slice nmber
     * @return AreaFormat record
     */
    private fun getAreaFormatPie(slice: Int): AreaFormat {
        // must add x number of dataformat recs
        // FINISH- not 100%
        /*		if (i==s.chartArr.size()) {
			df= (DataFormat) DataFormat.getPrototypeWithFormatRecs(0);
			df.setPointNumber(slice);
			PieFormat pf= (PieFormat) Chart.findRec(df.chartArr, PieFormat.class);
			pf.setPercentage((short)25);	// default percentage
			AttachedLabel al= (AttachedLabel) AttachedLabel.getPrototype();
			al.setType("CandP");		// default
			df.addChartRecord(al);
			s.chartArr.add(s.chartArr.size()-1, df);	// -1 to skip SERTOCRT rec
		}
*/
        return Chart.findRec(chartArr, AreaFormat::class.java) as AreaFormat
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3526272512004348462L

        /**
         * Create a new dataformat
         *
         * @param parentChart
         * @return
         */
        // creates default bar shape==0, 0
        val prototype: XLSRecord?
            get() {
                val df = DataFormat()
                df.opcode = XLSConstants.DATAFORMAT
                df.data = df.PROTOTYPE_BYTES
                df.init()
                val cs = Chart3DBarShape()
                cs.opcode = XLSConstants.CHART3DBARSHAPE
                df.addChartRecord(cs)
                return df
            }

        fun getPrototypeWithFormatRecs(parentChart: Chart): XLSRecord {
            return getPrototypeWithFormatRecs(0, parentChart)
        }


        /**
         * Create DataFormat Record that HOPEFULLY reflects the necessary associated recs
         *
         * @param type
         * @return DataFormat Record
         */
        fun getPrototypeWithFormatRecs(seriesNumber: Int, parentChart: Chart): XLSRecord {
            val df = DataFormat()
            df.opcode = XLSConstants.DATAFORMAT
            df.data = df.PROTOTYPE_BYTES
            df.init()
            val cs = Chart3DBarShape()
            cs.opcode = XLSConstants.CHART3DBARSHAPE    // creates default bar shape==0, 0
            df.addChartRecord(cs)
            df.setSeriesNumber(seriesNumber)
            val lf = LineFormat.prototype as LineFormat?
            df.addChartRecord(lf)
            val af = AreaFormat.prototype as AreaFormat?
            af!!.parentChart = parentChart
            df.addChartRecord(af)
            val pf = PieFormat.prototype as PieFormat?
            pf!!.parentChart = parentChart
            df.addChartRecord(pf)
            val mf = MarkerFormat.prototype as MarkerFormat?
            mf!!.parentChart = parentChart
            df.addChartRecord(mf)
            /* pieformat:
	MUST not exist on chart group types other than  ---> doesn't appear true (???)
	pie,
	doughnut,
	bar of pie, or
	pie of pie.
	MUST not exist if the chart group type is doughnut and the series is not the outermost series.
	MUST not exist on the data points on the secondary bar/pie of a bar of pie chart group.
          */
            return df
        }
    }


}
/**
 * sets this chart to have lines (line chart, radar, scatter ...)
 */
