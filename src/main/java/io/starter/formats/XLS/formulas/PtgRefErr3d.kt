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
package io.starter.formats.XLS.formulas

import io.starter.toolkit.ByteTools


/*
   An Erroneous BiffRec range spanning 3rd dimension of WorkSheets.

	identical to PtgRef3d

 * @see Ptg
 * @see Formula


*/
class PtgRefErr3d : PtgRef3d(), Ptg {

    override val isRefErr: Boolean
        get() = true

    override val string: String?
        get() = if (sheetname == null) "#REF!" else sheetname!! + "!#REF!"

    override val length: Int
        get() = Ptg.PTG_REFERR3D_LENGTH


    override val rowCol: IntArray
        get() = intArrayOf(-1, -1)

    override val value: Any?
        get() = if (sheetname == null) "#REF!" else sheetname!! + "!#REF!"

    override var location: String?
        get() = if (sheetname == null) "#REF!" else sheetname!! + "!#REF!"
        set

    // IDs: 3C (R) 5C (V) 7C (A)
    init {
        record = ByteArray(Ptg.PTG_REFERR3D_LENGTH)
        record[0] = 0x3c
        // record[1]= index to REF entry in EXTERNSHEET

    }

    /*
     Throw this data into a ptgref's
     Ixti can reference sheets that don't exist, causing np error.  As we don't perform any functions
     upon a PTGRef3D error, just swallow
    */
    override fun populateVals() {
        ixti = ByteTools.readShort(record[1].toInt(), record[2].toInt())
        if (ixti > 0) this.sheetname = GenericPtg.qualifySheetname(this.sheetName)
    }

    override fun setLocation(s: Array<String>) {
        sheetname = GenericPtg.qualifySheetname(s[0])
    }

    companion object {


        private val serialVersionUID = 8691902605148033701L
    }

}
