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

import io.starter.formats.XLS.XLSConstants
import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger

import java.io.OutputStream
import java.io.UnsupportedEncodingException
import java.nio.ByteBuffer
import java.util.AbstractList
import java.util.ArrayList


/**
 * Defines a 'file' in the LEO filesystem.  contains pointers
 * to Block storages as well as the storage type etc.
 *
 *
 *
 *
 * Header (128 bytes):
 * Directory Entry Name (64 bytes): This field MUST contain a Unicode string for the storage or stream name encoded in UTF-16. The name MUST be terminated with a UTF-16 terminating null character. Thus storage and stream names are limited to 32 UTF-16 code points, including the terminating null character. When locating an object in the compound file except for the root storage, the directory entry name is compared using a special case-insensitive upper-case mapping, described in Red-Black Tree. The following characters are illegal and MUST NOT be part of the name: '/', '\', ':', '!'.
 * Directory Entry Name Length (2 bytes): This field MUST match the length of the Directory Entry Name Unicode string in bytes. The length MUST be a multiple of 2, and include the terminating null character in the count. This length MUST NOT exceed 64, the maximum size of the Directory Entry Name field.
 * Object Type (1 byte): This field MUST be 0x00, 0x01, 0x02, or 0x05, depending on the actual type of object. All other values are not valid.
 * 0= Unknown or unallocated
 * 1= Storage Object
 * 2- Stream Object
 * 5= Root Storage Object
 * Color Flag (1 byte): This field MUST be 0x00 (red) or 0x01 (black). All other values are not valid.
 * Left Sibling ID (4 bytes): This field contains the Stream ID of the left sibling. If there is no left sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).
 * Right Sibling ID (4 bytes): This field contains the Stream ID of the right sibling. If there is no right sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).
 * Child ID (4 bytes): This field contains the Stream ID of a child object. If there is no child object, then the field MUST be set to NOSTREAM (0xFFFFFFFF).
 * CLSID (16 bytes): This field contains an object classGUID, if this entry is a storage or root storage. If there is no object class GUID set on this object, then the field MUST be set to all zeroes. In a stream object, this field MUST be set to all zeroes. If not NULL, the object class GUID can be used as a parameter to launch applications.
 * State Bits (4 bytes): This field contains the user-defined flags if this entry is a storage object or root storage object. If there are no state bits set on the object, then this field MUST be set to all zeroes.
 * Creation Time (8 bytes): This field contains the creation time for a storage object. The Windows FILETIME structure is used to represent this field in UTC. If there is no creation time set on the object, this field MUST be all zeroes. For a root storage object, this field MUST be all zeroes, and the creation time is retrieved or set on the compound file itself.
 * Modified Time (8 bytes): This field contains the modification time for a storage object. The Windows FILETIME structure is used to represent this field in UTC. If there is no modified time set on the object, this field MUST be all zeroes. For a root storage object, this field MUST be all zeroes, and the modified time is retrieved or set on the compound file itself.
 * Starting Sector Location (4 bytes): This field contains the first sector location if this is a stream object. For a root storage object, this field MUST contain the first sector of the mini st	ream, if the mini stream exists.
 * Stream Size (8 bytes): This 64-bit integer field contains the size of the user-defined data, if this is a stream object. For a root storage object, this field contains the size of the mini stream.
 *
 *
 * ??
 * offset  type value           const?  function
 * 00:     stream $pps_rawname       !  name of the pps
 * 40:     word $pps_sizeofname      !  size of $pps_rawname
 * 42:     byte $pps_type		      !  type of pps (1=storage|2=stream|5=root)
 * 43:	    byte $pps_uk0		      !  ?
 * 44:	    long $pps_prev	          !  previous pps
 * 48:     long $pps_next            !  next pps
 * 4c:     long $pps_dir             !  directory pps
 * 50:     stream 00 09 02 00        .  ?
 * 54:     long 0                    .  ?
 * 58:     long c0                   .  ?
 * 5c:     stream 00 00 00 46        .  ?
 * 60:     long 0                    .  ?
 * 64:     long $pps_ts1s            !  timestamp 1 : "seconds"		creation time
 * 68:     long $pps_ts1d            !  timestamp 1 : "days"
 * 6c:     long $pps_ts2s            !  timestamp 2 : "seconds"		modified time
 * 70:     long $pps_ts2d            !  timestamp 2 : "days"
 * 74:     long $pps_sb              !  starting block of property
 * 78:     long $pps_size            !  size of property
 * 7c:     long                      .  ?
 */
open class Storage : BlockByteReader {

    internal var DIR_COLOR_RED: Byte = 0
    internal var DIR_COLOR_BLACK: Byte = 1


    /**
     * return the existing header headerData for this storage
     */
    @Transient
    var headerData = ByteBuffer.allocate(128)

    // properties of this storage file.
    internal var name = ""
    internal var nameSize = -1

    var storageType: Byte = -1
    var directoryColor: Byte = -1
    var prevStorageID = -1
    var nextStorageID = -1
    var childStorageID = -1

    var sz = -1
    var miniStreamStorage = false
    var isSpecial = false

    var myblocks: MutableList<*>? = null
    private var startBlock = 0
    private var SIZE = -1
    /**
     * returns whether this Storage's headerData blocks are contained
     * in the Small or Big Block arrays.
     */
    /**
     * sets whether this Storage's headerData blocks are contained
     * in the Small or Big Block arrays.
     */
    var blockType = -1
        set(type) {
            field = type
            when (type) {
                Block.BIG -> SIZE = BIGBLOCK.SIZE
                Block.SMALL -> SIZE = SMALLBLOCK.SIZE
            }
        }
    var initialized = false
    internal var idxs: AbstractList<*>? = CompatibleVector()

    internal var lastblock: Block? = null

    /**
     * returns a new BlockByteReader
     *
     * @return
     */
    val blockReader: BlockByteReader
        get() = BlockByteReader(myblocks, actualFileSize)

    val blockVect: List<*>?
        get() = myblocks

    /**
     * get the position of this Storage in the file bytes
     */
    val filePos: Int
        get() = (getStartBlock() + 1) * SIZE

    /**
     * return the underlying byte array for this
     * Storage
     */
    /**
     * set the underlying byte array for this
     * Storage
     */
    open var bytes: ByteArray
        get() = LEOFile.getBytes(myblocks)
        set(b) {
            if (this.miniStreamStorage) {
                this.setMiniFATSectorBytes(b)
                return
            }
            val bs = BlockFactory.getBlocksFromByteArray(b, Block.BIG)
            if (myblocks != null) {
                myblocks!!.clear()
                lastblock = null
            } else {
                myblocks = ArrayList(bs!!.size)
            }
            for (d in bs!!.indices)
                this.addBlock(bs[d])
        }

    /**
     * return the underlying byte array for this
     * Storage
     */
    val byteStream: OutputStream
        get() = LEOFile.getByteStream(myblocks)

    /**
     * get the size of the Storage headerData
     */
    /**
     * set the size of the Storage headerData
     */
    var actualFileSize: Int
        get() {
            val pos = 0x78
            headerData.position(pos)
            this.sz = headerData.int
            return this.sz
        }
        set(i) {
            val pos = 0x78
            this.sz = i
            headerData.position(pos)
            headerData.putInt(i)
        }

    /**
     * return this Storage's existing headerData Blocks.
     */
    internal var blocks: Array<Block>?
        get() {
            if (myblocks!!.size < 1) {
                return this.initBigBlocks()
            }
            var blox = arrayOfNulls<Block>(this.myblocks!!.size)
            blox = myblocks!!.toTypedArray() as Array<Block>
            return blox
        }
        set(blks) {
            myblocks = ArrayList()
            for (t in blks.indices)
                this.addBlock(blks[t])
        }

    /**
     * get the number of Block records
     * that this Storage needs to store its
     * byte array
     */
    val sizeInBlocks: Int
        get() = myblocks!!.size


    /**
     * remove a block from this Storage's headerData
     */
    internal fun removeBlock(b: Block) {
        myblocks!!.remove(b)
    }

    fun getStorageType(): Int {
        return storageType.toInt()
    }

    fun getDirectoryColor(): Int {
        return directoryColor.toInt()
    }

    /**
     * set the value of the prevProp variable
     */
    fun setDirectoryColor(o: Int) {
        val pos = 0x43
        headerData.position(pos)
        headerData.put(o.toByte())
        this.directoryColor = o.toByte()
    }


    fun getName(): String {
        return name
    }

    override fun toString(): String {
        return getName() + " n:" + nextStorageID + " p:" + prevStorageID + " c:" + childStorageID + " sz:" + sz
    }

    /**
     * sets the storage name
     *
     * @param nm
     */
    fun setName(nm: String) {
        try {
            val b = nm.toByteArray(charset(XLSConstants.UNICODEENCODING))
            var pos = 0                // unicode name bytes
            headerData.position(pos)
            headerData.put(b)
            pos = 0x40                    // short name size
            headerData.position(pos)
            headerData.putShort((b.size + 2).toShort())
        } catch (e: UnsupportedEncodingException) {
        }

        name = nm
    }

    /**/
    constructor() {
        // empty constructor...
    }

    /**
     * create a new Storage record with a byte
     * array containing its record headerData.
     */
    constructor(buff: ByteBuffer) {
        headerData = buff
        var pos = 0x40
        headerData.position(pos)
        nameSize = headerData.short.toInt()
        if (nameSize > 0) {
            // get the name
            pos = 0
            val namebuf = ByteArray(nameSize)
            try {
                for (i in 0 until nameSize) {
                    namebuf[i] = headerData.get(i)
                }
            } catch (e: Exception) {
            }

            try {
                name = String(namebuf, XLSConstants.UNICODEENCODING)
            } catch (e: UnsupportedEncodingException) {
                Logger.logWarn(
                        "Storage error decoding storage name $e")
            }

            name = name.substring(0, name.length - 1)
            // if this line fails, your header BBD index is wrong...

        } else {
            // empty storage
        }
        pos = 0x42
        headerData.position(pos)
        storageType = headerData.get()
        directoryColor = headerData.get()
        prevStorageID = headerData.int
        nextStorageID = headerData.int
        childStorageID = headerData.int

        sz = this.actualFileSize
        if (sz > 0 && sz < BIGBLOCK.SIZE)
            miniStreamStorage = true
        if (LEOFile.DEBUG)
            Logger.logInfo("Storage: " + name + " storageType: " + storageType + " directoryColor:" + directoryColor +
                    " prevSID:" + prevStorageID + " nextSID:" + nextStorageID + " childSID:" + childStorageID + " sz:" + sz)
    }

    /**
     * get the value of the prevProp variable
     */
    fun getPrevStorageID(): Int {
        return this.prevStorageID
    }

    /**
     * set the value of the prevProp variable
     */
    fun setPrevStorageID(o: Int) {
        val pos = 0x44
        headerData.position(pos)
        headerData.putInt(o)
        this.prevStorageID = o
    }

    /**
     * get the value of the nextProp variable
     */
    fun getNextStorageID(): Int {
        return this.nextStorageID
    }

    /**
     * set the value of the nextProp variable
     */
    fun setNextStorageID(o: Int) {
        val pos = 0x48
        headerData.position(pos)
        headerData.putInt(o)
        this.nextStorageID = o
    }

    /**
     * get the value of the child storage id
     */
    fun getChildStorageID(): Int {
        return this.childStorageID
    }

    /**
     * set the value of the child storage id
     */
    fun setChildStorageID(o: Int) {
        val pos = 0x4C
        headerData.position(pos)
        headerData.putInt(o)
        this.childStorageID = o
    }

    /**
     * return the underlying byte array for this
     * Storage
     */
    fun writeBytes(out: OutputStream) {
        val itx = myblocks!!.iterator()
        while (itx.hasNext()) {
            (itx.next() as Block).writeBytes(out)
        }
    }

    /**
     * set the underlying byte array for this
     * Storage
     */
    fun writeBytes(out: OutputStream, blen: Int) {
        val bs = BlockFactory.getBlocksFromOutputStream(out, blen, Block.BIG)
        if (myblocks != null) {
            myblocks!!.clear()
            lastblock = null
        } else {
            myblocks = ArrayList(bs!!.size)
        }
        for (d in bs!!.indices)
            this.addBlock(bs[d])
    }

    /**
     * set the underlying byte array for this
     * Storage
     */
    fun setOutputBytes(b: OutputStream, blen: Int) {
        val bs = BlockFactory.getBlocksFromOutputStream(b, blen, Block.BIG)
        if (myblocks != null) {
            myblocks!!.clear()
            lastblock = null
        } else {
            myblocks = ArrayList(bs!!.size)
        }
        for (d in bs!!.indices)
            this.addBlock(bs[d])
    }

    /**
     * Sets bytes on a miniFAT storage
     *
     * @param b
     */
    private fun setMiniFATSectorBytes(b: ByteArray) {
        val bs = BlockFactory.getBlocksFromByteArray(b, Block.SMALL)
        if (myblocks != null) {
            myblocks!!.clear()
            lastblock = null
        } else {
            myblocks = ArrayList(bs!!.size)
        }
        for (d in bs!!.indices)
            this.addBlock(bs[d])
    }

    /**
     * sets bytes for this storage; length of newbytes determines
     * whether the storage is miniFAT or regular
     * <br></br>Incldes padding of bytes to ensure blocks are a factor of required block size
     *
     * @param newbytes
     */
    fun setBytesWithOverage(newbytes: ByteArray) {
        var newbytes = newbytes
        val actuallen = newbytes.size
        myblocks = ArrayList()    // clear out
        if (newbytes.size < StorageTable.BIGSTORAGE_SIZE) {    // usual case
            val overage = newbytes.size % 128
            if (overage > 0) {
                val b = ByteArray(128 - overage)
                newbytes = ByteTools.append(b, newbytes)
            }
            val smallblocks = BlockFactory.getBlocksFromByteArray(newbytes, Block.SMALL)
            for (i in smallblocks!!.indices)
                this.addBlock(smallblocks[i])
            this.blockType = Block.SMALL
        } else {
            val overage = newbytes.size % BIGBLOCK.SIZE
            if (overage > 0) {
                val b = ByteArray(BIGBLOCK.SIZE - overage)
                newbytes = ByteTools.append(b, newbytes)
            }
            val blocks = BlockFactory.getBlocksFromByteArray(newbytes, Block.BIG)
            for (i in blocks!!.indices)
                this.addBlock(blocks[i])
            this.blockType = Block.BIG
        }
        this.actualFileSize = actuallen
    }

    /**
     * Associate this storage with it's data
     * (obtained by walking the miniFAT sector index to reference blocks in the miniStream
     */
    @Throws(LEOIndexingException::class)
    fun initFromMiniStream(miniStream: List<*>, miniFAT: IntArray?) {
        if (this.getStartBlock() < 0)
            return

        if (miniFAT == null) { // error: trying to access smallblocks but no smallblock container found
            if (LEOFile.DEBUG)
                Logger.logWarn("initMiniFAT: no miniFAT container found")
            return
        }
        myblocks = ArrayList()
        var endloop = false
        var thisBlock: Block? = null

        var idx = this.getStartBlock()
        while (idx >= 0) {
            when (idx) {
                -1    // unused sector, shouldn't get to here
                    , -2    // end of sector marker, exit
                -> endloop = true

                else -> if (idx >= miniStream.size) {
                    Logger.logWarn("MiniStream Error initting Storage: " + this.getName())
                } else {
                    thisBlock = miniStream[idx] as Block    // miniFAT is 0-based (no header sector at position 0 as in regular FAT)
                    this.addBlock(thisBlock)
                }
            }
            if (endloop)
                break
            idx = miniFAT[idx]    // otherwise, walk the sector id chain
        }
        if (LEOFile.DEBUG) {
            if (Math.ceil(this.actualFileSize / 64.0).toInt() != myblocks!!.size)
                Logger.logErr("Number of miniStream Sectors does not equal storage size.  Expected: " + Math.ceil(this.actualFileSize / 64.0).toInt() + ". Is: " + myblocks!!.size)
        }
        this.initialized = true
    }

    /**
     * associate this storage with its blocks of data
     * (obtained by walking down the FAT sector index chain to obtain blocks from the dta block store)
     *
     *
     * for some bizarre reason, at idx 30208 the extraDIFAT jumps 5420 to 35628
     */
    fun init(dta: List<*>, FAT: IntArray, keepStartBlock: Boolean) {
        myblocks = ArrayList()
        var endloop = false
        if (getStartBlock() < 0)
            return
        var thisbb: Block? = null
        var nextIdx = 0 //, lastIdx = 0, specialOffset = 1;

        // ksc: for root block and miniFAT cont., we add start block to block list
        if (keepStartBlock) {
            // for root storages, add rootstart block
            thisbb = dta[startBlock + 1] as Block
            this.addBlock(thisbb) //;
        }
        var i = startBlock
        while (i < FAT.size) {
            nextIdx = FAT[i]

            when (nextIdx) {

                -4 // extraDIFAT sector
                -> Logger.logInfo(
                        "INFO: Storage.init() encountered extra DIFAT sector.")

                -3 // special block	= DIFAT - defines the FAT
                -> if (this.actualFileSize > 0) {
                    if (LEOFile.DEBUG)
                        Logger.logWarn(
                                "WARNING: Storage.init() Special block containing headerData.")
                    this.isSpecial = true

                    thisbb = dta[i++] as Block
                    if (!thisbb.isSpecialBlock) {
                        this.addBlock(thisbb) //;
                    }
                    nextIdx = i
                } else {
                    endloop = true
                }

                -1 // unused
                -> endloop = true

                -2 // end of Storage - keep end block
                -> {
                    if (i + 1 < dta.size) {
                        // get the "padding" block for later retrieval
                        thisbb = dta[i + 1] as Block
                        if (thisbb == null)
                            break
                        this.addBlock(thisbb) //
                        //}
                    }
                    endloop = true
                }

                else // normal block
                -> {
                    if (dta.size > nextIdx)
                        thisbb = dta[nextIdx] as Block
                    if (thisbb == null)
                        break
                    if (nextIdx != i + 1) {
                        //the next is a jumper, pickup the orphan
                        if (LEOFile.DEBUG)
                            Logger.logInfo(
                                    "INFO: Storage init: jumper skipping: $i")
                        val skipbb = dta[i + 1] as Block

                        this.addBlock(skipbb) //
                    } else if (!thisbb.isSpecialBlock) { // just skip as probably a bbdix in the midst of the secid chain
                        this.addBlock(thisbb) //

                    }
                }
            }// ksc
            i = nextIdx
            if (endloop)
                break
        }

        if (LEOFile.DEBUG) {
            val sz = this.actualFileSize
            if (sz != 0) {
                if (Math.ceil(sz / 512.0) != myblocks!!.size.toDouble())
                    Logger.logWarn("Storage.init:  Number of blocks do not equal storage size")
            }
        }
        this.initialized = true
    }

    /**
     * adds a block of data to this storage
     *
     * @param b
     */
    fun addBlock(b: Block?) {
        if (lastblock != null)
            lastblock!!.setNextBlock(b)

        if (b!!.initialized) {
            if (LEOFile.DEBUG) Logger.logWarn("ERROR: $this - Block is already initialized.")
            return
        }
        b.storage = this
        b.initialized = true
        if (myblocks == null)
            myblocks = ArrayList()
        myblocks!!.add(b)
        lastblock = b
    }

    /**
     * set the storage type
     */
    fun setStorageType(i: Int) {
        val pos = 0x42
        headerData.position(pos)
        headerData.put(i.toByte())
        storageType = i.toByte()
    }

    /**
     * set the starting block in the Block table
     * array for this Storage's headerData.
     */
    fun setStartBlock(i: Int) {
        this.startBlock = i
        val pos = 0x74
        this.headerData.position(pos)
        this.headerData.putInt(i)
    }

    /**
     * get the starting block for the Storage headerData
     */
    fun getStartBlock(): Int {
        val pos = 0x74
        headerData.position(pos)
        startBlock = headerData.int
        return startBlock
    }

    /**
     * return this Storage's headerData in new Blocks
     */
    private fun initBigBlocks(): Array<Block>? {
        // byte[] bb = this.getBytes();
        if (this.length > 0) {
            val blks = BlockFactory.getBlocksFromByteArray(this.bytes, Block.BIG)
            var t = 0
            myblocks!!.clear()
            while (t < blks!!.size) {
                if (t + 1 < blks.size)
                    blks[t].setNextBlock(blks[t + 1])
                this.addBlock(blks[t])
                t++
            }
            return blks
        }
        return null
    }

    /**
     * Track BB Index info
     */
    fun addIdx(x: Int) {
        idxs!!.add(Integer.valueOf(x))
    }

    override fun equals(other: Any?): Boolean {
        return other!!.toString() == this.toString()
    }

    /**
     * clear out object references in prep for closing workbook
     */
    fun close() {
        if (myblocks != null)
            myblocks!!.clear()
        if (idxs != null)
            idxs!!.clear()
        lastblock = null
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = -2065921767253066667L
        // make an enum in > 1.4
        internal var TYPE_INVALID: Byte = 0
        internal var TYPE_DIRECTORY: Byte = 1
        internal var TYPE_STREAM: Byte = 2
        internal var TYPE_LOCKBYTES: Byte = 3
        internal var TYPE_PROPERTY: Byte = 4
        internal var TYPE_ROOT: Byte = 5
    }
}