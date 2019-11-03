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

/**
 * Constants required for Chart functionality
 */

interface ChartConstants {
    companion object {
        // Chart Types (used in Chart Creation)
        val COLCHART = 0
        val BARCHART = 1
        val LINECHART = 2
        val PIECHART = 3
        val AREACHART = 4 // 20070703 KSC:
        val SCATTERCHART = 5 // ""
        val RADARCHART = 6 // ""
        val SURFACECHART = 7 // ""
        val DOUGHNUTCHART = 8 // ""
        val BUBBLECHART = 9 // ""
        val OFPIECHART = 10
        val PYRAMIDCHART = 11    // column-type pyramid
        val CYLINDERCHART = 12    // column-type cylinder
        val CONECHART = 13        // column-type cone
        val PYRAMIDBARCHART = 14// bar-type pyramid
        val CYLINDERBARCHART = 15    // bar-type cylinder
        val CONEBARCHART = 16    // bar-type cone
        val RADARAREACHART = 17 // ""
        val STOCKCHART = 18

        // Bar Shapes
        val SHAPEDEFAULT = 0
        val SHAPECOLUMN = SHAPEDEFAULT        // default
        val SHAPECYLINDER = 1
        val SHAPEPYRAMID = 256
        val SHAPECONE = 257
        val SHAPEPYRAMIDTOMAX = 516
        val SHAPECONETOMAX = 517

        // Axis Types
        val XAXIS = 0
        val YAXIS = 1
        val ZAXIS = 2
        val XVALAXIS = 3    // an X axis type but VAL records
    }
}