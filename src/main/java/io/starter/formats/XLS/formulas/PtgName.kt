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

import io.starter.formats.XLS.Formula
import io.starter.formats.XLS.FormulaNotFoundException
import io.starter.formats.XLS.Name
import io.starter.formats.XLS.WorkBook
import io.starter.toolkit.ByteTools
import io.starter.toolkit.FastAddVector
import io.starter.toolkit.Logger


/**
 * This PTG stores an index to a name.  The ilbl field is a 1 based index to the table
 * of NAME records in the workbook
 *
 *
 * OFFSET      NAME        sIZE        CONTENTS
 * ---------------------------------------------
 * 0           ilbl        2           Index to the NAME table
 * 2           (reserved)  2   `       Must be 0;
 *
 * @see Ptg
 *
 * @see Formula
 */
open class PtgName : GenericPtg, Ptg, IlblListener {
    internal var ilbl: Short = 0
    override var storedName: String
        internal set

    override val isOperand: Boolean
        get() = true

    // lookup Name object  in Workbook and return handle
    open val name: Name?
        get() {
            val b = this.parentRec!!.workBook
            var n: Name? = null
            try {
                n = b!!.getName(ilbl.toInt())
            } catch (ex: Exception) {
            }

            return n
        }

    open var `val`: Int
        get() = ilbl.toInt()
        set(i) {
            ilbl = i.toShort()
            this.updateRecord()
        }

    /*
     *
     * returns the string value of the name
		@see io.starter.formats.XLS.formulas.Ptg#getValue()
     */
    override// usual case
    // multiple values; create an array
    //    	String s = n.getName();
    //return n;
    val value: Any?
        get() {
            val n = name
            try {
                val p = n!!.cellRangePtgs
                if (p.size == 0) {
                    return "#NAME?"
                } else if (p.size == 1 || this.parentRec !is io.starter.formats.XLS.Array) {
                    return p[0].value
                } else {
                    var retarry = ""
                    for (i in p.indices) {
                        retarry = retarry + p[i].value + ","
                    }
                    retarry = "{" + retarry.substring(0, retarry.length - 1) + "}"
                    val pa = PtgArray()
                    pa.setVal(retarry)
                    return pa
                }
            } catch (e: Exception) {
            }

            return "#NAME?"
        }

    override val textString: String
        get() {
            val n = name ?: return "#NAME!"
            return n.name
        }

    /**
     * Override due to mystery extra byte
     * occasionally found in ptgName recs.
     */
    override val length: Int
        get() = if (record != null) record.size else Ptg.PTG_NAME_LENGTH

    override/*    	Ptg[] p = this.getName().getComponents();
    	for (int i=0;i<p.length;i++){
    	    Ptg[] pcomps = p[i].getComponents();
    		if (pcomps!= null){
    			for (int x=0;x<pcomps.length;x++){
    				v.add(pcomps[x]);
    			}
    		}else{
    			v.add(p[i]);
    		}
    	}
    	Ptg[] retPtgs = new Ptg[v.size()];
    	retPtgs = (Ptg[])v.toArray(retPtgs);
    	return retPtgs;*/ val components: Array<Ptg>?
        get() {
            val v = FastAddVector()
            val p = this.name!!.ptga
            val pcomps = p!!.components
            if (pcomps != null) {
                for (x in pcomps.indices) {
                    v.add(pcomps[x])
                }
            } else {
                v.add(p)
            }
            var retPtgs = arrayOfNulls<Ptg>(v.size)
            retPtgs = v.toTypedArray() as Array<Ptg>
            return retPtgs
        }

    /**
     * return referenced Names' location
     *
     * @see io.starter.formats.XLS.formulas.GenericPtg.getLocation
     */
    override var location: String?
        @Throws(FormulaNotFoundException::class)
        get() {
            if (this.name != null)
                try {
                    return this.name!!.location
                } catch (e: Exception) {
                }

            return null
        }
        set


    override fun init(b: ByteArray) {
        opcode = b[0]
        record = b
        this.populateVals()
        addToRefTracker()
    }

    //default constructor
    constructor() {
        opcode = 0x23        // reference type is default
    }

    /**
     * set the Ptg Id type to one of:
     * VALUE, REFERENCE or Array
     * <br></br>The Ptg type is important for certain
     * functions which require a specific type of operand
     */
    fun setPtgType(type: Short) {
        when (type) {
            Ptg.VALUE -> opcode = 0x43
            Ptg.REFERENCE -> opcode = 0x23
            Ptg.ARRAY -> opcode = 0x63
        }
        record[0] = opcode
    }

    // 20100218 KSC:
    // constructor which sets a specific id
    // to specify whether this PtgName is of value, ref or array type
    // (PtgNameV, PtgNameR or PtgNameA)
    constructor(id: Int) {
        opcode = id.toByte()
        //				0x23=   Ref
        //		ptgId = 0x43;	Value
    }


    /**
     * add this reference to the ReferenceTracker... this
     * is crucial if we are to update this Ptg when cells
     * are changed or added...
     */
    fun addToRefTracker() {
        //Logger.logInfo("Adding :" + this.toString() + " to tracker");
        try {
            if (parentRec != null)
                parentRec!!.workBook!!.refTracker!!.addPtgNameReference(this)
        } catch (ex: Exception) {
            Logger.logErr("PtgRef.addToRefTracker() failed.", ex)
        }

    }

    /**
     * For creating a ptg name from formula parser
     */
    open fun setName(name: String) {
        record = ByteArray(5)
        record[0] = opcode
        val b = this.parentRec!!.workBook
        ilbl = b!!.getNameNumber(name).toShort()
        this.addListener()
        record[1] = ilbl.toByte()
    }


    private fun populateVals() {
        ilbl = ByteTools.readShort(record[1].toInt(), record[2].toInt())
    }

    override fun getIlbl(): Short {
        return ilbl
    }

    override fun storeName(nm: String) {
        storedName = nm
    }


    override fun setIlbl(i: Short) {
        if (ilbl != i) {
            ilbl = i
            this.updateRecord()
        }
    }

    override fun updateRecord() {
        val brow = ByteTools.cLongToLEBytes(ilbl.toInt())
        record[1] = brow[0]
        record[2] = brow[1]
        if (parentRec != null) {
            if (parentRec is Formula)
                (parentRec as Formula).updateRecord()
        }
    }

    override fun toString(): String {
        return if (this.name != null) this.name!!.name else "[Null]"
    }

    override fun addListener() {
        val n = this.name
        if (n != null) {
            n.addIlblListener(this)
            this.storeName(n.name)
        }

    }

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 8047146848365098162L
    }
}
    
    