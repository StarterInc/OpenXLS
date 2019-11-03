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

import io.starter.OpenXLS.ExcelTools
import io.starter.formats.XLS.formulas.PtgArea
import io.starter.formats.XLS.formulas.PtgRef
import io.starter.toolkit.ByteTools

import java.io.Serializable
import java.util.Comparator
import java.util.TreeMap

/**
 * DiscontiguousRefStruct manages discontiguous ranges within one sheet.  Each of the ranges is managed via a Ref object,
 * be it range or single cell references.
 */
class DiscontiguousRefStruct : Serializable {
    // KSC: try to decrease processing time by using refPtgs treemap  private ArrayList sqrefs = new ArrayList();
    private val allrefs = refPtgs(refPtgComparer())
    internal var parentRec: XLSRecord? = null

    /**
     * Returns the refs in R1C1 format
     *
     * @return
     */
    /* KSC: try to decrease processing time by using refPtgs treemap         			    	String[] s = new String[sqrefs.size()];
        for (int i=0;i<s.length;i++) {
                s[i] = ((SqRef)sqrefs.get(i)).getLocation();
        }*/ val refs: Array<String>
        get() {

            val s = arrayOfNulls<String>(allrefs.size)
            val ptgs = allrefs.values.iterator()
            var i = 0
            while (ptgs.hasNext()) {
                try {
                    val pr = ptgs.next() as PtgRef
                    s[i++] = pr.location
                } catch (ex: Exception) {
                }

            }
            return s
        }

    /**
     * Returns the binary record of this SQRef for output to Biff8
     */
    /* KSC: try to decrease processing time by using refPtgs treemap
 		short s = (short) sqrefs.size();
        byte[] retData = ByteTools.shortToLEBytes(s);
        for (int i=0;i<sqrefs.size();i++) {
            SqRef sqr = (SqRef)sqrefs.get(i);
            retData = ByteTools.append(sqr.getRecordData(), retData);
         }*///        byte[] retData = ByteTools.shortToLEBytes((short)allrefs.size()); now done in updateData
    //        Object[] refs= allrefs.values().toArray();
    //        for (int i= refs.length-1; i >= 0; i--) {
    //              retData = ByteTools.append(getRecordData((PtgRef)refs[i]), retData);
    val recordData: ByteArray
        get() {

            var retData = ByteArray(0)
            val ptgs = allrefs.values.iterator()
            while (ptgs.hasNext()) {
                try {
                    val pr = ptgs.next() as PtgRef
                    retData = ByteTools.append(retData, getRecordData(pr))
                } catch (ex: Exception) {
                }

            }
            return retData
        }

    /**
     * Return the number of references this structure contains
     *
     * @return
     */
    //        return sqrefs.size();
    val numRefs: Int
        get() = allrefs.size

    /**
     * Get the bounding rectangle of all references in this structure in the format
     * {topRow,leftCol,bottomRow,rightCol}
     *
     *
     * Note that if you are including 3dReferences this could return inconsistent results
     *
     * @return
     */
    /* KSC: try to decrease processing time by using refPtgs treemap
        int[] retValues = {0,0,0,0};
        for (int i=0;i<sqrefs.size();i++) {
            SqRef sref = (SqRef)sqrefs.get(i);
            int[] locs = sref.getIntLocation();
            for (int x=0;x<2;x++) {
                if((locs[x]<retValues[x])||i==0)retValues[x]=locs[x];
            }
            for (int x=2;x<4;x++) {
                if((locs[x]>retValues[x])||i==0)retValues[x]=locs[x];
            }
        }*//*        Object[] refs= allrefs.values().toArray();
        for (int i= refs.length-1; i >= 0; i--) {*///        	PtgRef pr= (PtgRef)refs[i];
    val rowColBounds: IntArray
        get() {
            val retValues = intArrayOf(Integer.MAX_VALUE, Integer.MAX_VALUE, 0, 0)
            val ptgs = allrefs.values.iterator()
            var i = 0
            while (ptgs.hasNext()) {
                val pr = ptgs.next() as PtgRef
                val locs = pr.intLocation
                for (x in 0..1) {
                    if (locs!![x] < retValues[x]) retValues[x] = locs!![x]
                }
                for (x in 2..3) {
                    if (locs!![x] > retValues[x]) retValues[x] = locs!![x]
                }
                i++
            }
            return retValues
        }

    /**
     * Handles creating a SqRefStruct via a string passed in, this should be in the format
     * reference1,reference2,
     * or
     * 'A1:A5,G21,H33'
     */
    constructor(ranges: String, parentRec: XLSRecord) {
        this.parentRec = parentRec
        val refs = ranges.split(",".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
        for (i in refs.indices) {
            try {
                if (refs[i] != "") {
                    /* KSC: try to decrease processing time by using refPtgs treemap         			SqRef pref = new SqRef(refs[i],parentRec);
        			sqrefs.add(pref);*/
                    allrefs.add(refs[i], parentRec)
                }
            } catch (e: NumberFormatException) {
                //keep going
            }

        }
    }

    /**
     * Takes a binary array representing a sqref struct.
     *
     *
     * 2               rgbSqref            var         Array of 8 byte sqref structures, format shown below
     * in sqref class
     *
     * @param sqrefs
     */
    constructor(sqrefrec: ByteArray, parentRec: XLSRecord) {
        this.parentRec = parentRec
        var i = 0
        while (i < sqrefrec.size) {
            val sref = ByteArray(8)
            System.arraycopy(sqrefrec, i, sref, 0, 8)
            /* KSC: try to decrease processing time by using refPtgs treemap            SqRef s = new SqRef(sref, parentRec);
            sqrefs.add(s);*/
            allrefs.add(sref, parentRec)
            i += 8
        }
    }

    /**
     * Return toString as an array of references
     */
    override fun toString(): String {
        var result = "["
        /* KSC: TODO FINISH        Iterator i = sqrefs.iterator();
        while(i.hasNext()) {
            result += i.next().toString();
            if(i.hasNext())result += ", ";
        }*/
        result += "]"
        return result
    }

    /**
     * Adds a ref to the existing group of refs
     *
     * @param range
     */
    fun addRef(range: String) {
        /* KSC: try to decrease processing time by using refPtgs treemap            SqRef sr = new SqRef(range, this.parentRec);
        sqrefs.add(sr);*/
        allrefs.add(range, this.parentRec)
    }

    /**
     * Determines if the reference structure encompasses the
     * reference value passed in.
     *
     * @param rowcol
     * @return
     */
    fun containsReference(rowcol: IntArray): Boolean {
        /* KSC: try to decrease processing time by using refPtgs treemap
    	for (int i=0;i<sqrefs.size();i++) {
            SqRef sqr = (SqRef)sqrefs.get(i);
            if (sqr.contains(rowcol))return true;
        }*/
        return allrefs.containsReference(rowcol)
    }

    private fun getRecordData(myPtg: PtgRef): ByteArray {
        var retData = ByteArray(0)
        val rc = myPtg.rowCol
        if (rc[0] >= 65536)
        // TODO: if XLSX, would this get here????
            rc[0] = -1
        if (rc[2] >= 65536)
            rc[2] = -1
        retData = ByteTools.append(ByteTools.shortToLEBytes(rc[0].toShort()), retData)
        retData = ByteTools.append(ByteTools.shortToLEBytes(rc[2].toShort()), retData)
        retData = ByteTools.append(ByteTools.shortToLEBytes(rc[1].toShort()), retData)
        retData = ByteTools.append(ByteTools.shortToLEBytes(rc[3].toShort()), retData)
        return retData
    }

    /**
     * SQRef is a Ref (PtgArea) with specific methods to get and set via a different byte array than
     * ptgArea normally uses.
     *
     *
     * Sqref Structures
     *
     *
     * OFFSET       NAME            SIZE        CONTENTS
     * -----
     * 2            rwFirst             2           First row in reference
     * 4            rwLast              2           Last row in reference
     * 6            colFirst            2           First column in reference
     * 8            colLast             2           Last column in reference
     */
    internal inner class SqRef : Serializable {
        var myPtg: PtgRef


        val location: String?
            get() = myPtg.location

        val intLocation: IntArray?
            get() = myPtg.intLocation

        /**
         * Get the sqref as a byte array in the standardized SqRef structure
         *
         * @return
         */
        // TODO: if XLSX, would this get here????
        val recordData: ByteArray
            get() {
                var retData = ByteArray(0)
                val rc = myPtg.rowCol
                if (rc[0] >= 65536)
                    rc[0] = -1
                if (rc[2] >= 65536)
                    rc[2] = -1
                retData = ByteTools.append(ByteTools.shortToLEBytes(rc[0].toShort()), retData)
                retData = ByteTools.append(ByteTools.shortToLEBytes(rc[2].toShort()), retData)
                retData = ByteTools.append(ByteTools.shortToLEBytes(rc[1].toShort()), retData)
                retData = ByteTools.append(ByteTools.shortToLEBytes(rc[3].toShort()), retData)
                return retData
            }

        /**
         * Construct an Sqref structure from a range string
         *
         * @param range
         * @param parentRec
         */
        constructor(range: String?, parentRec: XLSRecord) {
            // add handling for different ref types, for now utilize a PtgArea
            if (range != null && range != "") {
                myPtg = PtgArea(range, parentRec)
                myPtg.addToRefTracker()
            }
        }


        override fun toString(): String? {
            return myPtg.toString()
        }

        /**
         * Constructor for usage with any sort of ptgRef
         *
         * @param ptg
         */
        constructor(ptg: PtgRef) {
            myPtg = ptg
        }


        /**
         * Determine if the ref contains the rowcol passed in
         *
         * @return
         */
        operator fun contains(rowcol: IntArray): Boolean {
            return if (myPtg is PtgArea) {
                (myPtg as PtgArea).contains(rowcol)
            } else false
        }

        /**
         * Construct an Sqref structure from a byte array
         * in the format specified above for an sqref
         *
         *
         * We do a string conversion for the absolute reference translation
         * which all (?) sqrefs share.
         *
         * @param range
         * @param parentRec
         */
        constructor(b: ByteArray, parentRec: XLSRecord) {
            val row = ByteTools.readShort(b[0].toInt(), b[1].toInt()).toInt()
            val col = ByteTools.readShort(b[4].toInt(), b[5].toInt()).toInt()
            var row2 = ByteTools.readShort(b[2].toInt(), b[3].toInt()).toInt()
            if (row2 < 0)
            // if row truly references MAXROWS_BIFF8 comes out -
                row2 += XLSConstants.MAXROWS_BIFF8

            val col2 = ByteTools.readShort(b[6].toInt(), b[7].toInt()).toInt()
            val rc = intArrayOf(row, col, row2, col2)
            val bool = booleanArrayOf(true, true, true, true)
            val location = ExcelTools.formatRangeRowCol(rc, bool)
            myPtg = PtgArea(location, parentRec)
            myPtg.addToRefTracker()
        }

        companion object {
            private const val serialVersionUID = -7923448634000437926L
        }

    }

    companion object {

        private const val serialVersionUID = -7923448634000437926L
    }

}

internal class refPtgs
/**
 * set the custom Comparitor for tracked Ptgs
 * Tracked Ptgs are referened by a unique key that is based upon it's location and it's parent record
 *
 * @param c
 */
(c: Comparator<*>) : TreeMap<*, *>(c), Serializable {

    /**
     * single access point for unique key creation from a ptg location and the ptg's parent
     * would be so much cleaner without the double precision issues ... sigh
     *
     * @param o
     * @return
     */
    @Throws(IllegalArgumentException::class)
    private fun getKey(o: Any?): Any {
        val loc = (o as PtgRef).hashcode
        if (loc == -1)
        // may happen on a referr (should have been caught earlier) or if not initialized yet
            throw IllegalArgumentException()

        val ploc = (o as PtgRef).parentRec!!.hashCode().toLong()
        return longArrayOf(loc, ploc)
    }

    /**
     * access point for unique key creation from separate identities
     *
     * @param loc  -- location key for a ptgref
     * @param ploc -- location key for a parent rec
     * @return Object key
     */
    private fun getKey(loc: Long, ploc: Long): Any {
        return longArrayOf(loc, ploc)
    }

    /**
     * override of add to record ptg location hash + parent has for later lookups
     */
    fun add(o: Any): Boolean {
        try {
            super.put(this.getKey(o), o)
        } catch (e: IllegalArgumentException) {    // SHOULD NOT HAPPEN -- happens upon RefErrs but they shouldnt be added ...
        }

        return true
    }

    fun add(b: ByteArray, parentRec: XLSRecord): Boolean {
        val row = ByteTools.readShort(b[0].toInt(), b[1].toInt()).toInt()
        val col = ByteTools.readShort(b[4].toInt(), b[5].toInt()).toInt()
        var row2 = ByteTools.readShort(b[2].toInt(), b[3].toInt()).toInt()
        if (row2 < 0)
        // if row truly references MAXROWS_BIFF8 comes out -
            row2 += XLSConstants.MAXROWS_BIFF8

        val col2 = ByteTools.readShort(b[6].toInt(), b[7].toInt()).toInt()
        val rc = intArrayOf(row, col, row2, col2)
        val bool = booleanArrayOf(true, true, true, true)
        val location = ExcelTools.formatRangeRowCol(rc, bool)
        val pa = PtgArea(location, parentRec)
        pa.addToRefTracker()
        return this.add(pa)
    }

    fun add(range: String?, parentRec: XLSRecord?): Boolean {
        // add handling for different ref types, for now utilize a PtgArea
        if (range != null && range != "") {
            val pa = PtgArea(range, parentRec)
            pa.addToRefTracker()
            return this.add(pa)
        }
        return false
    }

    /**
     * override of the contains method to look up ptg location + parent record via hashcode
     * to see if it is already contained within the store
     */
    operator fun contains(o: Any): Boolean {
        return super.containsKey(getKey(o))
    }

    fun containsReference(rc: IntArray): Boolean {
        val loc = PtgRef.getHashCode(rc[0], rc[1])    // get location in hashcode notation
        var m: Map<*, *>? = this.subMap(getKey(loc, 0), getKey(loc + 1, 0))    // +1 for max parent
        val key: Any
        if (m != null && m!!.size > 0) {
            val ii = m!!.keys.iterator()
            while (ii.hasNext()) {
                key = ii.next()
                val testkey = (key as LongArray)[0]
                return if (testkey == loc) {    // longs to remove parent hashcode portion of double
                    //System.out.print(": Found ptg" + this.get((Integer)locs.get(key)));
                    true
                } else
                    break    // shouldn't hit here
            }
        }
        // now see if test cell falls into any areas
        m = this.tailMap(getKey(SECONDPTGFACTOR, 0)) // ALL AREAS ...
        if (m != null) {
            val ii = m!!.keys.iterator()
            while (ii.hasNext()) {
                key = ii.next()
                val testkey = (key as LongArray)[0]
                val firstkey = (testkey / SECONDPTGFACTOR).toDouble()
                val secondkey = (testkey % SECONDPTGFACTOR).toDouble()
                if (firstkey.toLong() <= loc && secondkey.toLong() >= loc) {
                    val col0 = firstkey.toInt() % XLSRecord.MAXCOLS
                    val col1 = secondkey.toInt() % XLSRecord.MAXCOLS
                    val rw0 = (firstkey / XLSRecord.MAXCOLS).toInt() - 1
                    val rw1 = (secondkey / XLSRecord.MAXCOLS).toInt() - 1
                    if (this.isaffected(rc, intArrayOf(rw0, col0, rw1, col1))) {
                        //System.out.print(": Found area " + ((PtgRef)this.get(index)));
                        return true
                    }
                } else if (firstkey > loc)
                // we're done
                    break
            }
        }
        //io.starter.toolkit.Logger.log("");

        return false
    }

    /**
     * returns true if cell coordinates are contained within the area coordinates
     *
     * @param cellrc
     * @param arearc
     * @return
     */
    private fun isaffected(cellrc: IntArray, arearc: IntArray): Boolean {
        if (cellrc[0] < arearc[0]) return false // row above the first ref row?
        if (cellrc[0] > arearc[2]) return false // row after the last ref row?

        return if (cellrc[1] < arearc[1]) false else cellrc[1] <= arearc[3] // col before the first ref col?
// col after the last ref col?
    }

    /**
     * remove this PtgRef object via it's key
     */
    override fun remove(o: Any?): Any {
        return super.remove(getKey(o))
    }

    fun toArray(): Array<Any> {
        return this.values.toTypedArray()
    }

    companion object {
        private const val serialVersionUID = -7923448634000437926L
        val SECONDPTGFACTOR = XLSRecord.MAXCOLS.toLong() + XLSRecord.MAXROWS.toLong() * XLSRecord.MAXCOLS
    }
}

/**
 * custom comparitor which compares keys for TrackerPtgs
 * consisting of a long ptg location hash, long parent record hashcode
 */
internal class refPtgComparer : Comparator<*>, Serializable {
    override fun compare(o1: Any, o2: Any): Int {
        val key1 = o1 as LongArray
        val key2 = o2 as LongArray
        if (key1[0] < key2[0]) return -1
        if (key1[0] > key2[0]) return 1
        if (key1[0] == key2[0]) {
            if (key1[1] == key2[1]) return 0    // equals
            if (key1[1] < key2[1]) return -1
            if (key1[1] > key2[1]) return 1
        }
        return -1
    }
}
