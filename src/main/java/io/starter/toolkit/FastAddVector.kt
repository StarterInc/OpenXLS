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

import java.util.Enumeration

/**
 * a Vector class designed to provide forwards compatibility for JDK1.1
 * programs.
 *
 *
 *
 *
 * // add; toArray; iterator; insert; get; indexOf; remove // TreeList = 1260
 * 7360; 3080; 160; 170; 3400; 170; // ArrayList = 220 1480; 1760; 6870; 50;
 * 1540; 7200; // LinkedList = 270 7360; 3350; 55860; 290720; 2910; 55200;
 */
class FastAddVector : SpecialArrayList, java.io.Serializable {

    internal var hits = 0
    internal var misses = 0

    constructor() : super() {}

    constructor(i: Int) : super() {}// super(i);

    override fun iterator(): Iterator<*> {
        return SpecialArrayList.Itr()
    }

    /*
     * If passed in a Double, it stores it as a Double in the vector in logical
     * order (1,2,3) Returns true if all values can be represented by double and the
     * element is inserted
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
        obj.recordIndexHint = idx
        super.add(idx, obj)
    }

    override fun remove(obj: Any?): Boolean {
        return super.remove(obj)
    }

    override fun clear() {
        super.clear()
    }

    override fun toArray(): Array<Any> {
        if (true)
            return elementData
        val obj = arrayOfNulls<Any>(super.size)
        for (i in 0 until super.size) {
            obj[i] = super.get(i)
        }
        return obj
    }

    fun removeAllElements() {
        super.clear()
    }

    /**
     * @param obar
     */
    fun copyInto(obar: Array<Any>) {
        for (x in obar.indices) {
            super.add(obar[x])
        }
    }

    fun insertElementAt(ob: Any, i: Int) {
        super.add(i, ob)
    }

    fun lastElement(): Any {
        return super.get(super.size - 1)
    }

    fun elements(): Enumeration<*> {
        return FastAddVectorEnumerator(this)
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

    internal inner class FastAddVectorEnumerator(itx: FastAddVector) : Enumeration<Any> {

        private var it: FastAddVector? = null
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

    companion object {
        /**
         * serialVersionUID
         */
        private const val serialVersionUID = -5615615290731997512L
    }

}
