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
package io.starter.OpenXLS

//IMPORT OOXML-specific items

import io.starter.formats.XLS.charts.Ai
import io.starter.formats.XLS.charts.Series


/**
 * Chart Series Handle allows for manipulation of Chart Series within a Chart Handle.
 * <br></br>
 * Charts are typically made up of two data elements, series data and category data.
 * When Microsoft Excel creates a chart, it assigns either the worksheet rows as series data (and the
 * columns as category data), or the worksheet columns as the series data (and the rows as categories).
 * <br></br>
 * Excel tries to minimize the number of series by default, unless both rows and columns are equal, in which case
 * the default is to have the rows as series.
 * <br></br>
 * Typically there will be one series object per row (or column if that is the series) of data.  In most cases, the categories
 * will be the same for every series.
 * <br></br>
 * As there are series for every row of data, when performing actions such as adding a column to a chart that has it's series
 * arranged by row, it will be necessary to modify every series within the chart.  ChartSeriesHandle is available to
 * make this process easier.
 *
 * @see ChartHandle
 */

class ChartSeriesHandle
/**
 * Constructor, used internally.  For public use get series from a ChartHandle
 *
 * @param Series   series
 * @param WorkBook wbk - WorkBook the Chart is attached to
 */
(private val myseries: Series, private val wbh: WorkBookHandle) {

    /**
     * returns the Cell Range String representing the Data in this Series Object.
     *
     * @return String Cell Range representing Series Data e.g. Sheet1!A1:A12
     */
    /**
     * sets the data for this Series to be obtained from a new Cell Range.
     * <br></br>
     * Note that if you set the size of this series to a different amount of data than other Series
     * in the chart you could have unexpected results.
     *
     * @param String seriesRange - a Cell Range representing Series Data e.g. Sheet1!A1:A12
     */
    // 20080215 KSC: Added
    // ""
    // 20070711 KSC: update value (series) count for series
    var seriesRange: String
        get() = myseries.seriesValueAi!!.toString()
        set(seriesRange) {
            myseries.parentChart!!.setMetricsDirty()
            val ai = myseries.seriesValueAi
            ai!!.parentChart = myseries.parentChart
            ai.setSheet(myseries.parentChart!!.sheet)
            if (ai.workBook == null) ai.workBook = wbh.workBook
            ai.changeAiLocation(ai.toString(), seriesRange)
            try {
                val coords = io.starter.OpenXLS.ExcelTools.getRangeCoords(seriesRange)
                myseries.valueCount = coords[4]
                myseries.parentChart!!.setMetricsDirty()
            } catch (e: Exception) {

            }

        }

    /**
     * returns the Cell Range String representing the Categories (usually the X Axis) for the Chart
     * <br></br>Note that Category typically stays constant for all Series in the Chart
     *
     * @return String Category Cell Range e.g. Sheet1!A1:C1
     */
    /**
     * sets the Category from a new Cell Range.
     *
     *
     * IMPORTANT!  In most cases, the category should be set to the same Cell Range for all Series in the chart.
     * <br></br>This method is available for complex charts, but in most cases the ideal way to handle this call is through
     * <br></br>chartHandle.changeCategoryRange.
     *
     * @param String categoryRange - new Category Cell Range
     */
    // 20070711 KSC: update Category Count for this series
    var categoryRange: String
        get() = myseries.categoryValueAi!!.toString()
        set(categoryRange) {
            myseries.parentChart!!.setMetricsDirty()
            val ai = myseries.categoryValueAi
            if (ai!!.workBook == null) ai.workBook = wbh.workBook
            ai.changeAiLocation(ai.toString(), categoryRange)
            try {
                val coords = io.starter.OpenXLS.ExcelTools.getRangeCoords(categoryRange)
                myseries.categoryCount = coords[4]
                myseries.parentChart!!.setMetricsDirty()
            } catch (e: Exception) {

            }

        }

    /**
     * returns the Cell Range String representing Bubble Size for this Series (Bubble Charts only)
     *
     * @return String Bubble Sizes Cell Range e.g. Sheet1!A1:C1
     */
    val bubbleSizes: String
        get() = myseries.bubbleValueAi!!.toString()

    /**
     * returns the Legend text for this Series
     *
     * @return String Legend text for this Series
     * @see getSeriesLegendReference
     */
    /**
     * Set the Legend text for this Series
     * <br></br>note: series legend will then not be linked to a particular cell.  if you want to change the legend
     * reference, use setSeriesLegendRef
     *
     * @param String legend - Legend Text
     * @see setSeriesLegendRef
     */
    var seriesLegend: String
        get() = myseries.legendText
        set(legend) {
            myseries.setLegend(legend, this.wbh)
            myseries.parentChart!!.setMetricsDirty()
        }

    /**
     * returns the Legend Cell Reference, if any
     *
     * @return String Cell Address representing the Legend for this Series
     */
    val seriesLegendReference: String?
        get() {
            val ai = myseries.legendAi
            return ai?.definition
        }

    /**
     * gets the Chart Category Data Type
     *
     * @return int Data Type of the Category
     */
    /**
     * sets the Chart Category Data Type
     *
     * @param int i - Category Data Type
     */
    var categoryDataType: Int
        get() = myseries.categoryDataType
        set(i) {
            myseries.categoryDataType = i
            myseries.parentChart!!.setMetricsDirty()
        }

    /**
     * gets the Series Data Type
     *
     * @return int Data Type of this Series
     */
    /**
     * sets the Series Data Type
     *
     * @param int i - Series Data Type
     */
    var seriesDataType: Int
        get() = myseries.valueDataType
        set(i) {
            myseries.valueDataType = i
        }

    /**
     * returns a constant that represents the bar shape for this Series
     * <br></br>This is an internal method that is not useful to the end user.
     *
     * @return int DataPoint shape for this Series (3d Bar, Pyramid ...)
     */
    /**
     * sets the constant that represents the bar shape for this Series
     * <br></br>This is an internal method that is not useful to the end user.
     *
     * @param int shape - DataPoint shape for this Series (3d Bar, Pyramid ...)
     */
    var shape: Int
        get() = myseries.shape
        set(shape) {
            myseries.shape = shape
            myseries.parentChart!!.setMetricsDirty()
        }

    /**
     * returns the Series (bar or line) color
     * <br></br>NOTE: for Pie charts, use getPieChartSliceColor
     *
     * @return int color constant representing the Series color
     * @see FormatHandle.COLOR_* constants
     *
     * @see getSeriesColorStr
     *
     */
    /**
     * sets the color for this Series (bar or line)
     * <br></br>NOTE: for Pie Charts, use setPieChartSliceColor
     *
     * @param int clr - color constant
     * @see setPieChartSliceColor
     *
     * @see FormatHandle.COLOR_* constants
     */
    var seriesColor: Int
        @Deprecated("")
        get() = FormatHandle.HexStringToColorInt(myseries.seriesColor, FormatHandle.colorBACKGROUND)
        set(clr) {
            myseries.setColor(clr)
            myseries.parentChart!!.setMetricsDirty()
        }

    /**
     * returns the Series (bar or line) color hex string
     * <br></br>NOTE: for Pie charts, use getPieChartSliceColor
     *
     * @return int color constant representing the Series color
     * @see FormatHandle.COLOR_* constants
     *
     * @see getPieChartSliceColor
     */
    val seriesColorStr: String
        get() = myseries.seriesColor

    /**
     * returns true if this chart has Bubble Sizes
     *
     * @return
     */
    fun hasBubbleSizes(): Boolean {
        return myseries.hasBubbleSizes()
    }

    /**
     * set the Cell Range for the Bubbles in this Seeries (Bubble Chart Only)
     *
     * @param String bubbleSizes - Cell Range for Bubble Sizes
     */
    fun setBubbleRange(bubbleSizes: String?) {
        if (bubbleSizes != null && bubbleSizes != "") {
            myseries.parentChart!!.setMetricsDirty()
            val ai = myseries.bubbleValueAi
            if (ai!!.workBook == null) ai.workBook = wbh.workBook
            ai.changeAiLocation(ai.toString(), bubbleSizes)
            ai.setRt(2)
            try {    // also update Bubble Count for this series
                val coords = io.starter.OpenXLS.ExcelTools.getRangeCoords(bubbleSizes)
                myseries.bubbleCount = coords[4]
                myseries.parentChart!!.setMetricsDirty()
            } catch (e: Exception) {

            }

        }
    }

    /**
     * sets Cell Address for the Series Legend
     *
     * @param String legendCell - Cell Address for Legend
     */
    fun setSeriesLegendRef(legendCell: String) {
        myseries.legendRef = legendCell
        myseries.parentChart!!.setMetricsDirty()
    }

    /**
     * defines this Series with Cell References for the Legend, Data Points, Category and Bubble Sizes,
     * if applicable.
     *
     * @param String legendRef - Cell Address representing the Legend for this Series
     * @param String series - Cell Range representing the Data Points for this Series
     * @param String cat - Cell Range representing the Category for this Series
     * (NOTE: The Category Cell Range is typically the same for every Series in the Chart)
     * @param String bubble - Cell Range representing the Bubble Sizes for this Series (Bubble Chart only) or null if none
     */
    fun setSeries(legendRef: String, series: String, cat: String, bubble: String) {
        this.setSeriesLegendRef(legendRef)
        this.seriesRange = series
        this.categoryRange = cat
        this.setBubbleRange(bubble)
        myseries.parentChart!!.setMetricsDirty()
    }

    /**
     * sets the color for this series
     *
     * @param int seriesNumber	- series index
     * @param int clr - color constant
     * @see setSeriesColor
     */
    @Deprecated("use setSeriesColor(int clr) instead")
    fun setSeriesColor(seriesNumber: Int, clr: Int) {
        myseries.setColor(clr)
        myseries.parentChart!!.setMetricsDirty()
    }

    /**
     * sets the color for this Series (bar or line)
     * <br></br>NOTE: for Pie Charts, use setPieChartSliceColor
     *
     * @param String color hex string
     * @see setPieChartSliceColor
     *
     * @see FormatHandle.COLOR_* constants
     */
    fun setSeriesColor(clr: String) {
        myseries.setColor(clr)
        myseries.parentChart!!.setMetricsDirty()
    }

    /**
     * sets the color for a particular Pie Chart slice or wedge
     *
     * @param int clr - color constant
     * @param int slice	- 0-based slice or point number
     * @see FormatHandle.COLOR_* constants
     */
    fun setPieChartSliceColor(clr: Int, slice: Int) {
        myseries.setPieSliceColor(clr, slice)
        myseries.parentChart!!.setMetricsDirty()
    }

    /**
     * returns the desired Pie chart slice color
     *
     * @param int slice	- 0-based slice or wedge index
     * @return int color constant repressenting the color for the desired pie slice
     * @see FormatHandle.COLOR_* constants
     *
     */
    @Deprecated("See getPieChartSliceColorStr")
    fun getPieChartSliceColor(slice: Int): Int {
        return FormatHandle.HexStringToColorInt(myseries.getPieSliceColor(slice)!!, FormatHandle.colorBACKGROUND)
    }

    /**
     * returns the desired Pie chart slice color hex string
     *
     * @param int slice	- 0-based slice or wedge index
     * @return String color hex string representing the color for the desired pie slice
     */
    fun getPieChartSliceColorStr(slice: Int): String? {
        return myseries.getPieSliceColor(slice)
    }

    /*
     * returns the OOXML (Open Office XML) Shape Properties Object for this Series
     * @return SpPr OOXML Shape Properties Object
     * @see SpPr
     *
    public SpPr getSpPr() { return myseries.getSpPr(); }
    */

    /*
     * sets the OOXML (Open Office XML) Shape Properties Object for this Series
     * @param SpPr sp - OOXML Shape Properties Object
     * @see SpPr
     *
    public void setSpPr(SpPr sp) {
    	myseries.setSpPr(sp);
    }*/

    /*
     * returns the OOXML (Open Office XML) Marker Properties Object for this Series
     * @return Marker -OOXML Marker Properties Object
     * @see Marker
     *
    public Marker getMarker() { return myseries.getMarker(); }
    */
    /*
     * sets the OOXML (Open Office XML) Marker Properties Object for this Series
     * @param Marker m - the OOXML Marker Properties Object
     * @see Marker
     *
    public void setMarker(Marker m) {
    	myseries.setMarker(m);
    }*/

    /*
     * returns the OOXML (Open Office XML) Data Label Properties Object for this Series
     * @return  DLbls - OOXML Data Label Properties Object
     *
    public DLbls getDLbls() { return myseries.getDLbls(); }
    */

    /*
     * sets the OOXML (Open Office XML) Data Label Properties Object for this Series
     * @param Dlbls d- OOXML Data Label Properties Object
     * @see DLbls
     *
    public void setDLbls(DLbls d) {
    	myseries.setDLbls(d);
    }*/

    /*
     * returns an array of OOXML (Open Office XML) Data Point Properties Objects defining this Series 
     * @return DPt[] - Array of OOXML Data Point Properties Objects
     * @see Dpt
     *   
    public DPt[] getDPt() { return myseries.getDPt(); }*/
    /*
     * adds an OOXML (Open Office XML) Data Point Property Element to this Series
     * @param DPt d- Data Point Properties Object
     * @see DPt
     *
    public void addDpt(DPt d) {
    	myseries.addDpt(d);
    }*/
}
