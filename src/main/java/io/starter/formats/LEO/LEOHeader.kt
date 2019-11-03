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
package io.starter.formats.LEO

// NIO based API requires JDK1.3+

import io.starter.toolkit.Logger

import java.io.Serializable
import java.nio.ByteBuffer


/**
 * Header record containing information on the Storage records in the LEOFile.
 *
 *
 * Header (block 1) -- 512 (0x200) bytes
 * Field Description Offset Length Default value or const
 * FILETYPE Magic number identifying this as a LEO filesystem. 0x0000 Long 0xE11AB1A1E011CFD0
 * UK1 Unknown constant 0x0008 Integer 0
 * UK2 Unknown Constant 0x000C Integer 0
 * UK3 Unknown Constant 0x0014 Integer 0
 * UK4 Unknown Constant (revision?) 0x0018 Short 0x003B
 * UK5 Unknown Constant (version?) 0x001A Short 0x0003
 * UK6 Unknown Constant 0x001C Short -2
 * LOG_2_BIG_BLOCK_SIZE Log, base 2, of the big block size 0x001E Short 9 (2 ^ 9 = 512 bytes)
 * LOG_2_SMALL_BLOCK_SIZE Log, base 2, of the miniFAT size 0x0020 Integer 6 (2 ^ 6 = 64 bytes)
 * UK7 Unknown Constant 0x0024 Integer 0
 * UK8 Unknown Constant 0x0028 Integer 0
 * FAT_COUNT Number of elements in the FAT array 0x002C Integer required
 * PROPERTIES_START Block index of the first block of the property table 0x0030 Integer required
 * UK9 Unknown Constant 0x0034 Integer 0
 * UK10 Unknown Constant 0x0038 Integer 0x00001000
 * MINIFAT_START Block index of first big block containing the mini FATallocation table (MINIFAT) 0x003C Integer -2
 * UK11 Unknown Constant 0x0040 Integer 1
 * extraDIFAT Block index of the first block in the Extended Block Allocation Table 0x0044 Integer -2
 * extraDIFATCount Number of elements in the Extended File Allocation Table (to be added to the FAT) 0x0048 Integer 0
 * FAT_ARRAY Array of block indicies constituting the File Allocation Table (FAT) 0x004C, 0x0050, 0x0054 ... 0x01FC Integer[ ] -1 for unused elements, at least first element must be filled.
 * N/A Header block data not otherwise described in this table N/A
 */
class LEOHeader : Serializable {
    /**
     * get the Header bytes
     */
    @Transient
    var bytes: ByteBuffer? = null
        private set
    private var numFATSectors = -1            // FAT= array of Sector numbers
    /**
     * get the position of the Root Storage/1st Directory
     */
    var rootStartPos = -1
        private set                // Root Directory sector start
    private var miniFATStart = -2            // miniFAT stores sectors for storages < 4096
    private var extraDIFATStart = -1        // if more than 109 FAT sectors are needed this stores start position of remaining DIFAT
    private var numExtraDIFATSectors = -1    // usually 109 FAT sectors is enough; if > -1 more DIFAT is stored in other sectors
    private var numMiniFATSectors = 0
    /**
     * return the minimum stream size
     * usually 4096 or 8 blocks
     *
     * @return
     */
    internal var minStreamSize = 4096
        private set    // = 8 blocks minimum for a stream

    /**
     * get the FAT Sectors (sectors which hold the FAT or the chain that references all sectors in the file)
     */
    // more than 109 goes into DIFAT
    // START OF DIFAT (indexes the FAT)
    // start of the 1st 109 secIds (4 bytes each==436 bytes)
    val difat: IntArray
        get() {
            val numblks = Math.min(this.numFATSectors, 109)
            val FAT = IntArray(numblks)
            val pos = DIFATPOSITION
            bytes!!.position(pos)
            for (i in 0 until numblks) {
                FAT[i] = bytes!!.int
                FAT[i]++
            }
            return FAT
        }

    /**
     * initialize from existing data
     */
    fun init(): Boolean {
        return this.init(this.bytes!!)
    }


    /**
     * initialize the leo doc header
     */
    fun init(dta: ByteBuffer): Boolean {
        this.bytes = dta
        var pos: Int

        if (dta.limit() < 1)
            throw InvalidFileException("Empty input file.")

        try {
            // sanity check -- is it a valid WorkBook?
            for (t in 0..15) {
                if (dta.get(t) != majick[t]) {
                    throw InvalidFileException("File is not valid OLE format: Bad Majick.")
                }
            }
        } catch (e: Exception) {
            throw InvalidFileException("File is not valid OLE format:" + dta.limit())
        }

        // get the rootstart 0x30
        pos = 0x30
        bytes!!.position(pos)    // secId of 1st storage pos (==rootstart)[48]
        rootStartPos = bytes!!.int
        rootStartPos++
        rootStartPos *= BIGBLOCK.SIZE

        // has DIFAT Extra sectors? -- extended data blocks for large XLS files
        pos = 0x44    // [68]
        bytes!!.position(pos)
        extraDIFATStart = bytes!!.int    //  sector id of 1st sector of extra DIFAT sectors or -2 of no extra sectors used
        numExtraDIFATSectors = bytes!!.int    // total # sectors used for the DIFAT -- if # sectors or blocks > 109, more sectors are used

        // get minimum size of standard stream
        pos = 56
        bytes!!.position(pos)
        minStreamSize = bytes!!.int

        // starting at 0x2c (44) THE numFATSectors
        pos = 0x2C
        bytes!!.position(pos)
        numFATSectors = bytes!!.int    // total # blocks used in the FAT chain (each block can hold 128 sector ids)


        // 0x3c (60)  THE miniFAT start
        pos = 60
        bytes!!.position(pos)
        miniFATStart = bytes!!.int    // secId of 1st sector of the miniFAT sector chain or -2 if non-existant

        // pos 0x40= # miniFAT sectors [64]
        pos = 0x40
        bytes!!.position(pos)
        numMiniFATSectors = bytes!!.int
        return true
    }

    /**
     * get the number of the Extended DIFAT sectors
     */
    fun getNumExtraDIFATSectors(): Int {
        return numExtraDIFATSectors
    }

    /**
     * get the position of the extra DIFAT sector start
     */
    fun getExtraDIFATStart(): Int {
        return this.extraDIFATStart
    }

    /**
     * get the position of the miniFAT Start Block
     */
    fun getMiniFATStart(): Int {
        return miniFATStart
    }

    internal fun initConstants() {
        var pos = 0x08
        bytes!!.position(pos)
        bytes!!.putInt(0)

        pos = 0xc
        bytes!!.position(pos)
        bytes!!.putInt(0)

        pos = 0x10
        bytes!!.position(pos)
        bytes!!.putInt(0)

        pos = 0x14
        bytes!!.position(pos)
        bytes!!.putInt(0)

        pos = 0x18
        bytes!!.position(pos)
        bytes!!.putShort(0x3e.toShort())

        pos = 0x1a
        bytes!!.position(pos)
        bytes!!.putShort(0x3.toShort())

        pos = 0x1c
        bytes!!.position(pos)
        bytes!!.putShort((-2).toShort())

        pos = 0x1e
        bytes!!.position(pos)
        bytes!!.putShort(9.toShort())

        pos = 0x20
        bytes!!.position(pos)
        bytes!!.putInt(6)

        pos = 0x24
        bytes!!.position(pos)
        bytes!!.putInt(0)

        pos = 0x28
        bytes!!.position(pos)
        bytes!!.putInt(0)

        pos = 0x34
        bytes!!.position(pos)
        bytes!!.putInt(0)

        pos = 0x38
        bytes!!.position(pos)
        bytes!!.putInt(0x00001000)

    }

    internal fun initMajickBytes() {
        // create the majick number
        this.bytes!!.position(0)
        bytes!!.put(majick)
    }

    internal fun setData(dta: ByteBuffer?) {
        this.bytes = dta
    }

    /**
     * Set the FAT or sector chain index
     *
     * @param FAT
     */
    internal fun setFAT(FAT: IntArray) {
        val pos = DIFATPOSITION
        bytes!!.position(pos)
        for (i in FAT.indices) {
            if (FAT.size * 4 >= this.bytes!!.limit() - 4) {
                // TODO: Create extra DIFAT sectors FAT.length too big...
                Logger.logWarn("WARNING: LEOHeader.setFAT() creating Extra FAT Sectors Not Implemented.  Output file too large.")
            } else {
                bytes!!.putInt(FAT[i] - 1)    // todo: why decrement here?  necessitates an increment in LEOFile.writeBytes
            }
        }
        // fill rest with -1's === Empty Sector denotation
        var i = bytes!!.position()
        while (i < 512) {
            bytes!!.putInt(-1)
            i += 4
        }

        setNumFATSectors(FAT.size)
    }

    internal fun setRootStorageStart(i: Int) {
        val pos = 0x30
        bytes!!.position(pos)
        bytes!!.putInt(i)
    }

    internal fun setMiniFATStart(i: Int) {
        val pos = 0x3c
        this.miniFATStart = i
        bytes!!.position(pos)
        bytes!!.putInt(i)
        miniFATStart = i
    }

    internal fun setExtraDIFATStart(i: Int) {
        val pos = 0x44
        bytes!!.position(pos)
        bytes!!.putInt(i)
        extraDIFATStart = i
    }

    internal fun setNumExtraDIFATSectors(i: Int) {
        val pos = 0x48
        bytes!!.position(pos)
        bytes!!.putInt(i)
        numExtraDIFATSectors = i
    }

    internal fun setNumFATSectors(i: Int) {
        val pos = 0x2c
        bytes!!.position(pos)
        bytes!!.putInt(i)
        numFATSectors = i
    }

    internal fun getNumFATSectors(): Int {
        return numFATSectors
    }

    /**
     * set the number of miniFAT sectors (each sector=64 bytes)
     */
    internal fun setNumMiniFATSectors(i: Int) {
        val pos = 0x40
        bytes!!.position(pos)
        bytes!!.putInt(i)
        numMiniFATSectors = i
    }

    /**
     * return the number of miniFAT sectors
     *
     * @return
     */
    internal fun getNumMiniFATSectors(): Int {
        return numMiniFATSectors
    }

    companion object {

        /**
         * serialVersionUID
         */
        private const val serialVersionUID = -422489164065975273L
        val HEADER_SIZE = 0x200
        private val DIFATPOSITION = 76    // DIFAT start position WITHIN HEADER (DIFAT=where to find the FAT i.e the index to the FAT)
        val majick = byteArrayOf(0xd0.toByte(), 0xcf.toByte(), 0x11.toByte(), 0xe0.toByte(), 0xa1.toByte(), 0xb1.toByte(), 0x1a.toByte(), 0xe1.toByte(), 0x00.toByte(), 0x00.toByte(), 0x00.toByte(), 0x00.toByte(), 0x00.toByte(), 0x00.toByte(), 0x00.toByte(), 0x00.toByte())

        /**
         * structure of typical header:
         * // 24, 25= 62, 0  (rev #) = 0x3E
         * // 26, 27= 3, 0   (vers #) = 0x03
         * // 28, 29= -2, -1 (little endian)
         * // 30, 31= 9, 0   (size of sector (9= 512, 7= 128))
         * // 32, 33= 6, 0   (size of short sector (6=64 bytes))
         * // 34->43= 0      (unused)
         * // 44->47= Total # sectors in SAT ***
         * // 48->51= Sector id of 1st sector
         * // 52->55= 0      (unused)
         * // 56->59= Minimum size of a standard stream (usually 4096)
         * // 60->63= SecID of first sector in short-SAT
         * // 64->67= Total # sectors in Short-SAT ***
         * // 68->71= SecID of 1st sector of MSAT or -2 if no additional sectors are needed
         * // 72->75= Total # sectors used for the MSAT
         * // 76->436= MSAT (1st 109 sectors)
         */
        private fun displayHeader(data: ByteBuffer) {
            for (i in 0 until data.capacity()) {
                io.starter.toolkit.Logger.log("[" + i + "] " + data.get(i))
            }
        }

        /**
         * get a BIGBLOCK containing this Header's data
         * ie: create a new header byte block.
         */
        internal fun getPrototype(FAT: IntArray): LEOHeader {

            val retval = LEOHeader()

            // get an empty BIGBLOCK
            val retblock = BlockFactory.getPrototypeBlock(Block.BIG) as BIGBLOCK
            retval.setData(retblock.byteBuffer)
            retval.initMajickBytes()
            retval.initConstants()

            // recreate the FAT and indexes to ...
            retval.setFAT(FAT)
            retval.setMiniFATStart(-2)    // we don't rebuild miniFATs at this time ...
            retval.setRootStorageStart(1)
            retval.setExtraDIFATStart(-2)
            retval.setNumExtraDIFATSectors(0)
            return retval
        }
    }
}