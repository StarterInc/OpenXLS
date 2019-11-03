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
package io.starter.formats.XLS


/**
 * **Msodrawingselection: MS Office Drawing Selection (EDh)**<br></br>
 *
 *
 * These records contain only data.
 *
 *<pre>
 *
 * offset  name        size    contents
 * ---
 * 4       rgMSODrSelr    var    MSO Drawing Group selectionData
 *
</pre> *
 *
 * @see MSODrawing
 */
class MSODrawingSelection : io.starter.formats.XLS.XLSRecord() {
    var PROTOTYPE_BYTES = byteArrayOf(0, 0, 25, -15, 16, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 4, 0, 0, 1, 4, 0, 0)

    /**
     * bypass continue handling for msodrawingselection, until we start
     * modifing the record
     */
    override var data: ByteArray?
        get() = super.getData()
        set

    /**
     * not a lot going on here...
     */
    override fun init() {
        super.init()
    }

    companion object {

        /**
         *
         */
        private val serialVersionUID = 2799490308252319737L
    }

}
