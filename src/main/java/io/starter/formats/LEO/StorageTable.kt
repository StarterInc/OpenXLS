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

import io.starter.formats.XLS.BiffRec
import io.starter.formats.XLS.WorkBookException
import io.starter.formats.XLS.XLSRecordFactory
import io.starter.toolkit.ByteTools
import io.starter.toolkit.CompatibleVector
import io.starter.toolkit.Logger

import java.io.File
import java.io.FileOutputStream
import java.io.IOException
import java.io.Serializable
import java.nio.ByteBuffer
import java.nio.ByteOrder
import java.util.*


/**
 * The directory system for an LEO file
 *
 * [Extentech Inc.](http://www.extentech.com)
 */
class StorageTable : Serializable {

    private var myheader: LEOHeader? = null

    /**
     * get the CompatibleVector of all Directories within the LEO file
     */
    internal var allDirectories = CompatibleVector()    // all directories in the LEO file

    // Directory collection
    /**
     * get the Hashtable of Storages within the LEO file
     */
    internal var storageHash = Hashtable(100, 0.9f)

    private val dupct = 0

    /**
     * return the number of existing directories
     *
     * @return
     */
    val numDirectories: Int
        get() = allDirectories.size

    /**
     * init the directories and gather their associated data (if present)
     * including those directories whose data is gathered from the miniStream
     */
    internal var miniFAT: IntArray? = null    // index into the miniStream

    /**
     * get the Root Storage Record
     */
    internal val rootStorage: RootStorage
        @Throws(StorageNotFoundException::class)
        get() = this.getDirectoryByName("Root Entry") as RootStorage?


    /**
     * clear out object references in prep for closing workbook
     */
    fun close() {
        myheader = null

        val ii = storageHash.keys.iterator()
        while (ii.hasNext()) {
            val s = storageHash.get(ii.next()) as Storage
            s.close()
        }
        storageHash = Hashtable(100, 0.9f)


        for (i in allDirectories.indices) {
            val s = allDirectories[i] as Storage
            s.close()
        }
        allDirectories.clear()
    }

    /**
     * initialize the directory entry array
     */
    fun init(dta: ByteBuffer, h: LEOHeader, blockvect: List<*>, FAT: IntArray) {
        this.myheader = h
        val data = LEOFile.getBytes(this.initDirectoryStream(blockvect, FAT))
        val psbsize = data.size
        val numRecs = psbsize / DIRECTORY_SIZE
        if (LEOFile.DEBUG) {
            Logger.logInfo("Number of Directories: $numRecs")
            Logger.logInfo("Directories:  " + Arrays.toString(data))
        }

        var pos = 0
        for (i in 0 until numRecs) {
            val b = ByteBuffer.allocate(DIRECTORY_SIZE)
            b.order(ByteOrder.LITTLE_ENDIAN)
            b.put(data, pos, DIRECTORY_SIZE)
            pos += DIRECTORY_SIZE
            var rec: Storage? = null
            try {
                rec = Storage(b)
            } catch (ex: Exception) {
                throw WorkBookException("StorageTable.init failed:$ex", WorkBookException.UNSPECIFIED_INIT_ERROR)
            }

            if (i == 0) {
                rec = RootStorage(b)
                if (rec.getName() != "Root Entry")
                    rec.setName("Root Entry")    // can happen upon a mac-sourced file
            }
            this.addStorage(rec, -1)
        }
    }

    /**
     * Create a new Storage and add to the directory array
     * <br></br>NOTE: add any associated data separately
     *
     * @param name      Storage name - must be unique
     * @param type      Storage Type - 1= storage 2= stream 5= root 0=unknown or unallocated
     * @param insertIdx id where to insert, or -1 to insert at end
     * @return
     */
    fun createStorage(name: String, type: Int, insertIdx: Int): Storage {
        var s: Storage? = null
        try {
            val b = ByteBuffer.allocate(DIRECTORY_SIZE)
            b.order(ByteOrder.LITTLE_ENDIAN)
            b.put(ByteArray(DIRECTORY_SIZE))
            s = Storage(b)
            s.setName(name)
            s.setStorageType(type)
            s.blocks = arrayOfNulls(0) // init empty storages with 0 blocksW
            s.setPrevStorageID(-1)
            s.setNextStorageID(-1)
            s.setChildStorageID(-1)
            addStorage(s, insertIdx)
        } catch (ex: Exception) {
            throw WorkBookException("Storage.createStorage failed:$ex", WorkBookException.UNSPECIFIED_INIT_ERROR)
        }

        return s
    }

    fun initDirectories(blockvect: List<*>, FAT: IntArray) {
        val rcs = allDirectories.elements()
        var miniStream: List<*>? = null
        var totrecsize = 0

        while (rcs.hasMoreElements()) {
            val rec = rcs.nextElement() as Storage
            val name = rec.getName()
            val recsize = rec.actualFileSize
            //			if (LEOFile.DEBUG)
            //				Logger.logInfo("Initializing Directory: "	+ name + ".  Start=" + rec.getStartBlock() + ". Size=" + recsize);
            totrecsize += recsize
            if (name == "Root Entry") {
                // also sets miniFAT ... ugly, I know ...
                miniStream = this.initMiniStream(blockvect, FAT)    // grab the mini stream (short sector container) (if any), indexed by miniFAT
                if (LEOFile.DEBUG && miniFAT != null)
                    Logger.logInfo("miniFAT: " + Arrays.toString(miniFAT))
            } else if (recsize > 0 && recsize < BIGSTORAGE_SIZE) {
                rec.blockType = Block.SMALL
                this.initStorage(rec, miniStream, miniFAT, Block.SMALL)    // the miniStream is indexed by the miniFAT
                // Regular Sector file storage
            } else if (recsize >= BIGSTORAGE_SIZE) {
                this.initStorage(rec, blockvect, FAT, Block.BIG)

                // a storage-less directory
            } else if (recsize == 0) {
                rec.blocks = arrayOfNulls(0) // init empty storages with 0 blocksW
            } else {
                if (LEOFile.DEBUG)
                    Logger.logWarn("Storage has no Block Type.")
            }// this storage has it's data in the miniStream
        }
        if (LEOFile.DEBUG)
            Logger.logInfo(
                    "Total Size used by Directories : $totrecsize")

    }

    /**
     * // whenever a stream is shorter than a specific length
     * // it is stored as a short-stream or ministream.  ministreamentries do not directly
     * // use sectors to store their data, but are all embedded in a specific
     * // internal control stream.
     * // first used sector is obtained from the root store ==> miniFAT chain
     * // it's secID chain is contained in the miniFAT
     * // The data used by all of the short-sectors container stream are concatenated
     * // in order of the secID chain.  There is no header, so the first
     * // mini sector (secId= 0) is always located at position 0 in the mini Stream
     * // The miniFAT is the same as the FAT except the secID chains refer
     * // to miniSectors (64 bytes) rather than regular sectors or blocks (512 bytes)
     */
    fun initMiniStream(blockvect: List<*>, FAT: IntArray): List<*>? {
        val pos = myheader!!.miniFATStart
        if (pos == -2)
            return null // no miniStream sectors

        val miniFATSectors = getMiniFAT(pos, blockvect, FAT)
        if (miniFATSectors!!.size > 0) {
            try {
                val sbz = miniFATSectors.iterator()
                while (sbz.hasNext()) {
                    val sbb = sbz.next() as BIGBLOCK
                    sbb.isDepotBlock = true
                }

                miniFAT = LEOFile.readFAT(miniFATSectors)

                val rootStore: RootStorage
                try {
                    rootStore = this.getDirectoryByName("Root Entry") as RootStorage?
                } catch (e: StorageNotFoundException) {
                    throw InvalidFileException("Error parsing OLE File. OLE FileSystem Out of Spec: No Root Entry.")
                }

                // capture the short-stream container stream
                val miniStreamStart = rootStore.startBlock
                val miniStreamSize = rootStore.actualFileSize

                val miniStream = Storage()
                miniStream.startBlock = miniStreamStart
                miniStream.actualFileSize = miniStreamSize
                miniStream.setName("miniStream")
                // obtain the miniStream from the regular bigblock store
                this.initStorage(miniStream, blockvect, FAT, Block.BIG)
                // now that we have the entire miniStream , break it up into mini Sector-sized blocks
                // NOTE: only miniStreamSize bytes are usable - ignore rest
                val b = ByteArray(miniStreamSize)
                System.arraycopy(miniStream.bytes, 0, b, 0, miniStreamSize)
                val miniStreamBlocks = BlockFactory.getBlocksFromByteArray(b, Block.SMALL)
                val miniStreamBlockList = ArrayList()
                for (i in miniStreamBlocks!!.indices) {    // should equal sbbsize
                    miniStreamBlockList.add(miniStreamBlocks[i])
                }
                return miniStreamBlockList
            } catch (e: LEOIndexingException) {
                Logger.logWarn("initSBStorages: Error obtaining sbdIdx")
            }

        }
        return null
    }

    /**
     * extract the miniFAT sector index from the blockvect
     *
     * @param pos
     * @param blockvect
     * @param FAT
     * @return
     */
    private fun getMiniFAT(pos: Int, blockvect: List<*>, FAT: IntArray): List<*>? {
        var miniFATContainer: Storage? = null
        miniFATContainer = Storage()
        miniFATContainer.blockType = Block.BIG
        miniFATContainer.startBlock = pos
        miniFATContainer.setStorageType(5) // set as root type to distinguish from regular storages
        if (LEOFile.DEBUG)
            Logger.logInfo("StorageTable.getMiniFAT() Initializing miniFAT Container.")
        miniFATContainer.init(blockvect, FAT, true)
        //	miniFAT.setName("SBidx");
        // miniFAT index
        return miniFATContainer.blockVect
    }

    /**
     * add a Storage to the directory array
     *
     * @param Storage   storage to insert
     * @param insertIdx -1 if add at end, otherwise insert at spot
     */
    internal fun addStorage(rec: Storage, insertIdx: Int) {
        val nm = rec.getName()
        if (storageHash.get(nm) != null && nm != "") {
            /*	KSC: with 2012 code changes, this breaks output:
 * 		if (LEOFile.DEBUG)
				Logger.logInfo(
					"INFO: StorageTable.addStorage() Dupe Storage Name: " + nm);
			nm = nm + "|^" + dupct++;
			rec.setName(nm);
 Does not appear necessary
*/
        }
        storageHash.put(rec.getName(), rec)
        if (insertIdx == -1)
            allDirectories.add(rec)
        else {
            allDirectories.add(insertIdx, rec)
            for (i in allDirectories.indices) {    // adjust prev, next, child ids if necessary
                val s = allDirectories[i] as Storage
                if (s.getChildStorageID() >= insertIdx)
                    s.setChildStorageID(s.getChildStorageID() + 1)
                if (s.getPrevStorageID() >= insertIdx)
                    s.setPrevStorageID(s.getPrevStorageID() + 1)
                if (s.getNextStorageID() >= insertIdx)
                    s.setNextStorageID(s.getNextStorageID() + 1)

            }
        }

        if (false && LEOFile.DEBUG) {
            Logger.logInfo(
                    /*"INFO: StorageTable.addStorage() Storage size: "
					+*/ rec.getName()
                    + " Size: "
                    + rec.actualFileSize
                    + " Start Block: "
                    + rec.startBlock)
        }
    }

    /* remove a Storage
     */
    internal fun removeStorage(st: Storage) {
        this.storageHash.remove(st)
        this.allDirectories.remove(st)
    }

    /**
     * init a Storage
     */
    internal fun initStorage(rec: Storage, sourceblocks: List<*>?, idx: IntArray?, blocktype: Int) {
        var name: String? = rec.getName()
        val recsize = rec.actualFileSize
        if (LEOFile.DEBUG)
            Logger.logInfo(
                    "Initializing Storage: " + name + " Retrieving Data." +
                            " Size: "
                            + recsize
                            + " type: "
                            + blocktype
                            + " startidx: "
                            + rec.startBlock + if (rec.blockType == Block.SMALL) " MiniFAT" else "")
        rec.blockType = blocktype
        if ("Root Entry" == name)
        // ksc: shouldn't!
            return


        if (rec.blockType == Block.BIG)
            rec.init(sourceblocks, idx, false)
        else if (rec.blockType == Block.SMALL)
            rec.initFromMiniStream(sourceblocks, idx)

        if (LEOFile.DEBUG) {
            if (rec.bytes != null) {
                if (name == null)
                    name = "noname.dat"
                if (name[0] == '')
                    name = name.substring(1)
                if (name[0] == '')
                    name = name.substring(1)
                // if(blocktype == Block.BIG)
                StorageTable.writeitout(rec.blockVect, "$name.stor")
            }
        }
        rec.idxs = null
    }

    /**
     * get the directory BLOCKS or Sectors
     * contains the header info for all of the directories
     */
    internal fun initDirectoryStream(blockvect: List<*>, FAT: IntArray): List<*>? {
        //getDirectoryBlocks()
        var directories: Storage? = null
        directories = Storage()
        val pstart = this.myheader!!.rootStartPos / BIGBLOCK.SIZE - 1
        directories.startBlock = pstart
        directories.blockType = Block.BIG
        directories.setStorageType(5) // set to root directory
        directories.init(blockvect, FAT, true) // get additional directory stores, if any
        //		directories.setName("StorageTable");
        //		if (LEOFile.DEBUG)
        //			StorageTable.writeitout(directories.getBlockVect(),
        //									"directoryStorage.dat");
        return directories.blockVect


    }

    /**
     * generate new RootStorage bytes from
     * all of the Storage directories
     */
    internal fun rebuildRootStore(): ByteArray {
        /*
         * Free (unused) directory entries are marked with Object Type 0x0 (unknown or unallocated).
         * The entire directory entry should consist of all zeroes except for the child, right sibling,
         * and left sibling pointers, which should be initialized to NOSTREAM (0xFFFFFFFF).
         */
        while (allDirectories.size % 4 != 0) { // add "null" storages to ensure multiples of 4 (128*4=512==minimum size)
            this.createStorage("", 0, -1)
        }
        val e = allDirectories.elements()
        val bytebuff = ByteArray(allDirectories.size * DIRECTORY_SIZE)
        var pos = 0
        while (e.hasMoreElements()) {
            val s = e.nextElement() as Storage
            val buff = this.getDirectoryHeaderBytes(s)
            System.arraycopy(buff!!.array(), 0, bytebuff, pos, DIRECTORY_SIZE)
            pos += DIRECTORY_SIZE
        }
        return bytebuff
    }

    /**
     * get the directory by name.  throws StorageNotFoundException if not found.
     */
    @Throws(StorageNotFoundException::class)
    fun getDirectoryByName(name: String): Storage? {
        if (storageHash.get(name) != null)
            return storageHash.get(name)
        throw StorageNotFoundException("Storage: $name not located")
    }

    /**
     * returns the Stream ID of the named directory, or -1 if not found
     *
     * @param name
     * @return
     * @throws StorageNotFoundException
     */
    @Throws(StorageNotFoundException::class)
    fun getDirectoryStreamID(name: String): Int {
        for (i in allDirectories.indices) {
            val s = allDirectories[i] as Storage
            if (s.getName() == name)
                return i
        }
        return -1
    }

    /**
     * get the child directory (storage or stream) of the named storage
     *
     * @param name
     * @throws StorageNotFoundException
     * @return Storage or null
     */
    @Throws(StorageNotFoundException::class)
    fun getChild(name: String): Storage? {
        val s = getDirectoryByName(name)
        if (s != null) {
            val child = s.getChildStorageID()
            if (child > -1)
                return allDirectories[child] as Storage
        }
        return null
    }

    /**
     * get the next directory (storage or stream) of the named storage
     *
     * @param name
     * @throws StorageNotFoundException
     * @return Storage or null
     */
    @Throws(StorageNotFoundException::class)
    fun getNext(name: String): Storage? {
        val s = getDirectoryByName(name)
        if (s != null) {
            val next = s.getNextStorageID()
            if (next > -1)
                return allDirectories[next] as Storage
        }
        return null
    }

    /**
     * get the previous directory (storage or stream) of the named storage
     *
     * @param name
     * @throws StorageNotFoundException
     * @return Storage or null
     */
    @Throws(StorageNotFoundException::class)
    fun getPrevious(name: String): Storage? {
        val s = getDirectoryByName(name)
        if (s != null) {
            val prev = s.getNextStorageID()
            if (prev > -1)
                return allDirectories[prev] as Storage
        }
        return null
    }

    /**
     * get the header record for a Directory
     */
    internal fun getDirectoryHeaderBytes(thisStorage: Storage): ByteBuffer? {
        // create a new byte buffer
        var buff: ByteBuffer? = null
        buff = thisStorage.headerData
        // 20100304 KSC: don't reset anything, just return
        return buff
    }

    fun DEBUG() {
        io.starter.toolkit.Logger.log("DIRECTORY CONTENTS:")
        for (i in allDirectories.indices) {
            val s = allDirectories[i] as Storage
            val n = s.getName()
            Logger.logInfo("Storage: " + n + " storageType: " + s.getStorageType() + " directoryColor:" + s.getDirectoryColor() +
                    " prevSID:" + s.getPrevStorageID() + " nextSID:" + s.getNextStorageID() + " childSID:" + s.getChildStorageID() + " sz:" + s.actualFileSize)
            // special storages
            if (n == "Root Entry") {
                Logger.logInfo("Root Header: " + Arrays.toString(s.headerData.array()))

                /***********************************
                 * // KSC: TESTING for XLS-97:
                 * !!!  Set creation time and modified time on root storage to 0
                 * int p = 100;
                 * s.getHeaderData().position(p);
                 * Long tsCreated= s.getHeaderData().getLong();
                 * Long tsModified= s.getHeaderData().getLong();
                 */
                if (s.myblocks != null) {
                    val zz = 0
                    if (s.myblocks!![zz] is io.starter.formats.LEO.BIGBLOCK)
                        io.starter.toolkit.Logger.log("BLOCK 1:\t" + zz + "-" + Arrays.toString((s.myblocks!![zz] as io.starter.formats.LEO.BIGBLOCK).bytes))
                    else
                        io.starter.toolkit.Logger.log("BLOCK 1:\t" + zz + "-" + Arrays.toString((s.myblocks!![zz] as io.starter.formats.LEO.SMALLBLOCK).bytes))
                }
            } else if (n == "Workbook") {
                //skip
            } else if (n == "\u0001CompObj") {
                val bytes = s.blockReader
                val len = bytes.length
                val rec = io.starter.formats.XLS.XLSRecord()        // 4 bytes are header ...
                rec.byteReader = bytes
                rec.length = len
                val slen = ByteTools.readInt(rec.getBytesAt(24, 4))    // actually position 28
                if (slen >= 0) {
                    val ss = String(rec.getBytesAt(28, slen))        // AnsiUserType= a display name of the linked object or embedded object.
                    io.starter.toolkit.Logger.log("\tOLE Object:$ss")
                }
                // AnsiClipboardFormat (variable)
                //				io.starter.toolkit.Logger.log("\t" + Arrays.toString(rec.getData()));
            } else if (n.startsWith("000")) {        // pivot cache
                if (s.myblocks != null) {
                    for (zz in s.myblocks!!.indices)
                        if (s.myblocks!![zz] is io.starter.formats.LEO.BIGBLOCK)
                            io.starter.toolkit.Logger.log("\t" + zz + "-" + Arrays.toString((s.myblocks!![zz] as io.starter.formats.LEO.BIGBLOCK).bytes))
                        else
                            io.starter.toolkit.Logger.log("\t" + zz + "-" + Arrays.toString((s.myblocks!![zz] as io.starter.formats.LEO.SMALLBLOCK).bytes))
                }
                val bytes = s.blockReader
                val len = bytes.length
                var z = 0
                while (z <= len - 4) {
                    val headerbytes = bytes.getHeaderBytes(z)
                    val opcode = ByteTools.readShort(headerbytes[0].toInt(), headerbytes[1].toInt())
                    val reclen = ByteTools.readShort(headerbytes[2].toInt(), headerbytes[3].toInt()).toInt()
                    val rec = XLSRecordFactory.getBiffRecord(opcode)
                    rec.byteReader = bytes
                    rec.offset = z
                    rec.length = reclen.toShort()
                    rec.init()
                    io.starter.toolkit.Logger.log("\t\t" + rec.toString())
                    z += reclen + 4
                }
            } else {
                if (s.myblocks != null) {
                    for (zz in s.myblocks!!.indices)
                        if (s.myblocks!![zz] is io.starter.formats.LEO.BIGBLOCK)
                            io.starter.toolkit.Logger.log("\t" + zz + "-" + Arrays.toString((s.myblocks!![zz] as io.starter.formats.LEO.BIGBLOCK).bytes))
                        else
                            io.starter.toolkit.Logger.log("\t" + zz + "-" + Arrays.toString((s.myblocks!![zz] as io.starter.formats.LEO.SMALLBLOCK).bytes))
                }
            }
        }
    }

    companion object {

        /**
         * serialVersionUID
         */
        private const val serialVersionUID = 3399830613453524580L
        val TABLE_SIZE = 0x200
        val DIRECTORY_SIZE = 0x80
        val BIGSTORAGE_SIZE = 4096        // default, should read from LEOHeader

        /**
         * write out the directory bytes for debugging
         */
        fun writeitout(blocks: List<*>?, name: String) {
            try {
                val fos = FileOutputStream(File(System.getProperty("user.dir") + "\\storages\\" + name))
                fos.write(LEOFile.getBytes(blocks))
            } catch (a: IOException) {
            }

        }
    }
}