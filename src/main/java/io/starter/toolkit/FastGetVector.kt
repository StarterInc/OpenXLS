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
package io.starter.toolkit

import java.util.ArrayList
import java.util.Enumeration


/**
 * a Vector class designed to provide forwards compatibility
 * for JDK1.1 programs.
 */
class FastGetVector : ArrayList<Any> {
    private var change_offset = 0
    private val reindex_change_size = 1000

    internal var hits = 0
    internal var misses = 0

    /**
     * reset the hints for all vector elements
     * expense is linear to size but will increase
     * accuracy of subsequent 'indexOf' calls.
     */
    fun resetHints(ignore_records: Boolean) {
        // ExcelTools.benchmark("Re-indexing CompatibleVector" + reindex_change_size++);
        if (!ignore_records) {
            for (t in 0 until this.size) {
                try {
                    (this[t] as CompatibleVectorHints).recordIndexHint = t
                } catch (e: Exception) {
                    return
                }

            }
        }
        change_offset = 0
    }

    constructor() : super() {}

    constructor(i: Int) : super() {}// super(i);

    /*
     * If passed in a Double, it stores it as a Double in the vector
     * in logical order (1,2,3)
     * Returns true if all values can be represented by double and the element is inserted
     */
    fun addOrderedDouble(obj: Double): Boolean {
        try {
            for (i in 0 until super.size) {
                val dd = super.get(i) as Double
                if (dd > obj) {
                    super.add(i, obj)
                    return true
                }
            }
        } catch (e: Exception) {
            return false
        }

        super.add(obj)
        return true
    }

    /*
     * Returns the last object in the collection
     *
     */
    fun last(): Any {
        return super.get(super.size - 1)
    }


    fun add(idx: Int, obj: CompatibleVectorHints) {
        this.change_offset++//
        obj.recordIndexHint = idx
        super.add(idx, obj)
    }

    /*
        public void addAll(CompatibleVector cv){
            for(int i=0;i<cv.size();i++){
                Object b = cv.get(i);
                if(b instanceof CompatibleVectorHints){
                    this.add((CompatibleVectorHints)b);
                }else{
                    this.add((Object)b);
                }
            }
        }
    */
    override fun remove(obj: Any?): Boolean {
        this.change_offset--
        return super.remove(obj)
    }

    override fun clear() {
        this.change_offset = 0
        super.clear()
    }

    override fun toArray(): Array<Any> {
        val obj = arrayOfNulls<Any>(super.size)
        for (i in 0 until super.size) {
            obj[i] = super.get(i)
        }
        return obj
    }

    fun removeAllElements() {
        super.clear()
    }

    fun copyInto(obar: Array<Any>) {
        for (x in obar.indices) {
            super.add(obar)
        }
    }

    fun insertElementAt(ob: Any, i: Int) {
        super.add(i, ob)
    }

    fun lastElement(): Any {
        return super.get(super.size - 1)
    }

    fun elements(): Enumeration<*> {
        return FastGetVectorEnumerator(this)
    }

    fun elementAt(t: Int): Any {
        return super.get(t)
    }

    override fun toArray(obj: Array<Any>): Array<Any> {
        for (i in 0 until super.size) {
            obj[i] = super.get(i)
        }
        return obj
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -6701901995748359720L
    }

}

internal class FastGetVectorEnumerator(itx: FastGetVector) : Enumeration<Any> {

    private var it: FastGetVector? = null
    var x = 0

    init {
        it = itx
    }

    override fun nextElement(): Any {
        return it!!.elementAt(x++)
    }

    override fun hasMoreElements(): Boolean {
        return x < it!!.size
    }

} 
