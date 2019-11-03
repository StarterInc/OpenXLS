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

import io.starter.formats.XLS.Name
import io.starter.formats.XLS.WorkBook
import io.starter.toolkit.ByteTools


/*
	This PTG stores an index to a name.  The ilbl field is a 1 based index to the table
	of NAME records in the workbook

	OFFSET      NAME        sIZE        CONTENTS
	---------------------------------------------
	0           ixti        2           index to externsheet
	2           ilbl        2           Index to the NAME table
	4           (reserved)  2   `       Must be 0;

 * @see Ptg
 * @see Formula

*/
class PtgNameX : PtgName(), Ptg, IxtiListener {
    /**
     * @return Returns the ixti.
     */
    /**
     * @param ixti The ixti to set.
     */
    override var ixti: Short = 0
    override var `val`: Int = 0
        internal set(value: Int) {
            super.`val` = value
        }


    override val isOperand: Boolean
        get() = true

    //lookup Name object  in Workbook and return handle
    override// the externsheet reference is negative, there seems to be a problem
    // off the docs.  Just use a placeholder boundsheet, as the PtgRef3D internally will
    // get the value correctly
    //Externsheet x = b.getExternSheet();
    // it's an AddInFormula... -jm
    //Boundsheet[] bound = x.getBoundSheets(ixti);
    val name: Name?
        get() {
            val b = this.parentRec!!.sheet!!.workBook
            var n: Name? = null

            try {
                n = b!!.getName(`val`)
                n!!.setSheet(this.parentRec!!.sheet)
            } catch (e: Exception) {
            }

            return n
        }

    /*
     *
     * returns the string value of the name
        @see io.starter.formats.XLS.formulas.Ptg#getValue()
     */
    override val value: Any?
        get() {

            val b = this.parentRec!!.sheet!!.workBook
            var externalname: String? = null
            try {
                externalname = b!!.getExternalName(`val`)
            } catch (e: Exception) {
            }

            if (externalname != null)
                return externalname

            val n = name
            return n!!.calculatedValue
        }

    override val textString: String
        get() {
            val o = value ?: return ""
            return o.toString()
        }


    override val length: Int
        get() = Ptg.PTG_NAMEX_LENGTH

    override fun addListener() {
        try {
            parentRec!!.workBook!!.externSheet!!.addPtgListener(this)
        } catch (e: Exception) {
            // no need to output here.  NullPointer occurs when a ref has an invalid ixti, such as when a sheet was removed  Worksheet exception could never really happen.
        }

    }


    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        this.populateVals()
    }

    private fun populateVals() {
        ixti = ByteTools.readShort(record[1].toInt(), record[2].toInt())

        `val` = ByteTools.readShort(record[3].toInt(), record[4].toInt()).toInt()
    }

    override fun toString(): String {
        return if (this.parentRec!!.sheet != null)
            value as String?
        else
            "Uninitialized PtgNameX"
    }

    // KSC: Added to handle External names (denoted by PtgNameX records in ExpressionParser)

    /**
     * For creating a ptg namex from formula parser
     */
    override fun setName(name: String) {
        opcode = 0x39    // PtgNameX
        record = ByteArray(Ptg.PTG_NAMEX_LENGTH)
        record[0] = opcode
        val b = this.parentRec!!.sheet!!.workBook
        `val` = b!!.getExtenalNameNumber(name)
        ixti = b.externSheet!!.virtualReference.toShort()
        val bb = ByteTools.shortToLEBytes(ixti)
        record[1] = bb[0]
        record[2] = bb[1]
        val bbb = ByteTools.cLongToLEBytes(`val`)
        record[3] = bbb[0]
        record[4] = bbb[1]
    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 1240996941619495505L
    }
}
    
    