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
import org.json.JSONException
import org.json.JSONObject

/**
 * **Surface: Chart Group is a Surface Chart Group (0x103f) **
 *
 *
 * 4		grbit		2
 *
 *
 * 0		0x1		fFillSurface		1= chart contains color fill for surface
 * 1		0x2		f3DPhongShade		1= this surface chart has shading
 */
class Surface : GenericChartObject(), ChartObject {
    private var grbit: Short = 0
    private var fFillSurface = true
    private var f3DPhoneShade = false
    /**
     * surface charts always contain a 3d record so must determine if 3d separately
     */
    /**
     * set if this surface chart is "truly" 3d as all surface-type charts contain a 3d record
     *
     * @param is3d
     */
    var is3d = false    /* since all surface charts contain a 3d record, must store 3d setting separately */

    private val PROTOTYPE_BYTES = byteArrayOf(0, 0)

    /**
     * returns true if surface chart is wireframe, false if filled
     *
     * @return
     */
    /**
     * sets this surface chart to wireframe (true) or filled(false)
     */
    var isWireframe: Boolean
        get() = !fFillSurface
        set(wireframe) {
            fFillSurface = !wireframe
            grbit = ByteTools.updateGrBit(grbit, fFillSurface, 0)
        }

    /**
     * @return String XML representation of this chart-type's options
     */
    override val optionsXML: String
        get() {
            val sb = StringBuffer()
            if (fFillSurface)
                sb.append(" ColorFill=\"true\"")
            if (f3DPhoneShade)
                sb.append(" Shading=\"true\"")
            return sb.toString()
        }

    /**
     * return the (dojo) type JSON for this Chart Object
     *
     * @return
     */
    val typeJSON: JSONObject
        @Throws(JSONException::class)
        get() {
            val typeJSON = JSONObject()
            val dojoType: String
            if (!this.isStacked)
                dojoType = "Areas"
            else
                dojoType = "StackedAreas"
            typeJSON.put("type", dojoType)
            return typeJSON
        }

    override fun init() {
        super.init()
        grbit = ByteTools.readShort(this.getByteAt(0).toInt(), this.getByteAt(1).toInt())
        fFillSurface = grbit and 0x1 == 0x1
        f3DPhoneShade = grbit and 0x2 == 0x2
        chartType = ChartConstants.SURFACECHART
    }

    // 20070703 KSC:
    private fun updateRecord() {
        val b = ByteTools.shortToLEBytes(grbit)
        this.data[0] = b[0]
        this.data[1] = b[1]
    }

    /**
     * Handle setting options from XML in a generic manner
     */
    override fun setChartOption(op: String, `val`: String): Boolean {
        var bHandled = false
        if (op.equals("ColorFill", ignoreCase = true)) {
            fFillSurface = `val` == "true"
            grbit = ByteTools.updateGrBit(grbit, fFillSurface, 0)
            bHandled = true
        } else if (op.equals("Shading", ignoreCase = true)) {
            f3DPhoneShade = `val` == "true"
            grbit = ByteTools.updateGrBit(grbit, f3DPhoneShade, 1)
            bHandled = true
        }
        if (bHandled)
            updateRecord()
        return bHandled
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -3243029185139320374L

        val prototype: XLSRecord?
            get() {
                val b = Surface()
                b.opcode = XLSConstants.SURFACE
                b.data = b.PROTOTYPE_BYTES
                b.init()
                return b
            }
    }
}
