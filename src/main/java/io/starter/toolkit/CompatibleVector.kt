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

import java.util.NoSuchElementException
import java.util.Vector

/**
 * a Vector class designed to provide forwards compatibility for JDK1.1
 * programs.
 */
class CompatibleVector : Vector<Any> {
    private var change_offset = 0
    private val reindex_change_size = 1000

    internal var hits = 0
    internal var misses = 0

    /* */
    /**
     * Index of element to be returned by subsequent call to next.
     */
    internal var cursor = 0

    /**
     * Index of element returned by most recent call to next or previous. Reset to
     * -1 if this element is deleted by a call to remove.
     */
    internal var lastRet = -1

    /**
     * reset the hints for all vector elements expense is linear to size but will
     * increase accuracy of subsequent 'indexOf' calls.
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

    constructor(i: Int) : super(i) {}

    override fun iterator(): Iterator<*> {
        return Itr()
    }

    /**
     * if the object being checked implements index hints, the lookup can be
     * performed much faster
     *
     *
     * speed of lookups is affected by 'shuffling' positions of vector elements.
     */
    fun indexOf(r: CompatibleVectorHints): Int {
        // return super.indexOf(r);

        var x = r.recordIndexHint
        if (x > 0 && x < super.size)
            if (super.elementAt(x) != null) {
                if (super.elementAt(x) == r) {
                    // Logger.logInfo("hit/miss="+ hits++ + ":"+ misses);
                    // hits++;
                    return x
                }
            }
        x -= change_offset
        if (x > 0 && x < super.size)
            if (super.elementAt(x) == r) {

                // Logger.logInfo("hit/miss="+ hits++ + ":"+ misses);
                // hits++;
                return x
            }
        var t = -1
        if (change_offset > reindex_change_size)
            this.resetHints(false)
        // if(x>0)t = super.indexOf(r,x);
        if (x > 0)
            t = super.indexOf(r)
        if (t < 0)
            t = super.indexOf(r)
        r.recordIndexHint = t
        // Logger.logInfo("hit/miss="+ hits++ + ":"+ misses++);
        return t
        /**/
        /*
         * int idx = r.getRecordIndexHint(); int retval = -1; int recsz = this.size();
         * if(idx >= recsz)idx = recsz-1; if(idx < 0){ idx = this.indexOf((Object)r);
         * r.setRecordIndexHint(idx); return idx; } if(this.get(idx).equals(r))return
         * idx; else{ boolean found = false, checkhi = true, checklo = true; int hi
         * =idx, lo = idx; // int hi = idx + change_offset, lo = idx + change_offset;
         * //Logger.logInfo(change_offset); if(hi < 0)hi = recsz/2; if(lo > 0)lo =
         * recsz/2; while(!found && (checkhi || checklo)){ if(++hi >= recsz)checkhi =
         * false; if(--lo < 0)checklo = false; if(checkhi){ Object b = this.get(hi);
         * if(b.equals(r)){ found = true; retval = hi; // Logger.logInfo(" hi:" + (idx -
         * hi)); }else if(checklo){ Object c = this.get(lo); if(c.equals(r)){ found =
         * true; retval = lo; // Logger.logInfo("lo:" + (idx - lo)); } } } } }
         * Logger.logInfo(" hi:" + (idx - hi)); Logger.logInfo("lo:" + (idx - lo));
         * if((retval == -1)&&(change_offset > 0)){ this.resetHints(false); //
         * Logger.logInfo("Loopy!! " + r); retval = this.indexOf(r); }
         * r.setRecordIndexHint(retval); return retval;
         */
    }

    /**
     * overriding AbstractList so we can have concurrent mods...
     */
    operator fun next(): Any {
        // checkForComodification();
        try {
            val next = get(cursor)
            lastRet = cursor++
            return next
        } catch (e: IndexOutOfBoundsException) {
            throw NoSuchElementException()
        }

    }

    override fun get(idx: Int): Any {
        return super.elementAt(idx)
    }

    /*
     * If passed in a Double, it stores it as a Double in the vector in logical
     * order (1,2,3) Returns true if all values can be represented by double and the
     * element is inserted
     */
    fun addOrderedDouble(obj: Double): Boolean {
        try {
            for (i in 0 until super.size) {
                val dd = super.elementAt(i) as Double
                if (dd > obj) {
                    super.insertElementAt(obj, i)
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
        return super.elementAt(super.size - 1)
    }

    fun add(obj: CompatibleVectorHints?): Boolean {
        this.change_offset++ //
        val idx = super.size
        if (obj != null)
            obj.recordIndexHint = idx
        try {
            super.insertElementAt(obj, idx)
            return true
        } catch (e: Exception) {
            return false
        }

    }

    fun add(idx: Int, obj: CompatibleVectorHints) {
        this.change_offset++//
        obj.recordIndexHint = idx
        super.insertElementAt(obj, idx)
    }

    fun addAll(cv: CompatibleVector) {
        for (i in cv.indices) {
            val b = cv[i]
            if (b is CompatibleVectorHints) {
                this.add(b)
            } else {
                this.add(b)
            }
        }
    }

    override fun remove(obj: Any?): Boolean {
        if (super.remove(obj)) {
            this.change_offset--
            return true
        }
        return false
    }

    override fun clear() {
        this.change_offset = 0
        super.removeAllElements()
    }

    override fun toArray(): Array<Any> {
        val obj = arrayOfNulls<Any>(super.size)
        for (i in 0 until super.size) {
            obj[i] = super.elementAt(i)
        }
        return obj
    }

    override fun toArray(obj: Array<Any>): Array<Any> {
        for (i in 0 until super.size) {
            try {
                obj[i] = super.elementAt(i)
            } catch (e: Exception) {
                Logger.logInfo("CompatibleVector.toArray() failed.")
            }

        }
        return obj
    }

    private inner class Itr : Iterator<Any> {
        /**
         * Index of element to be returned by subsequent call to next.
         */
        internal var cursor = 0

        /**
         * Index of element returned by most recent call to next or previous. Reset to
         * -1 if this element is deleted by a call to remove.
         */
        internal var lastRet = -1

        /**
         * The modCount value that the iterator believes that the backing List should
         * have. If this expectation is violated, the iterator has detected concurrent
         * modification.
         */
        internal var expectedModCount = modCount

        override fun hasNext(): Boolean {
            return cursor != size
        }

        override fun next(): Any {
            val next = get(cursor)
            lastRet = cursor++
            return next
        }

        override fun remove() {
            if (lastRet == -1)
                throw IllegalStateException()

            this@CompatibleVector.removeAt(lastRet)
            if (lastRet < cursor)
                cursor--
            lastRet = -1
            expectedModCount = modCount
        }

    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 6805047965683753637L
    }
}