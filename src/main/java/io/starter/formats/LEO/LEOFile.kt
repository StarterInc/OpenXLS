/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.formats.LEO

import io.starter.formats.XLS.WorkBookException
import io.starter.toolkit.*

import java.io.*
import java.nio.ByteBuffer
import java.nio.ByteOrder
import java.util.*

/**
 * LEOFile is an archive format compatible with other popular archive formats such as OLE.
 *
 *
 * It contains multiple files or "Storages" which can
 * subsequently contain data in popular desktop application formats.
 */
class LEOFile : Serializable {
    var DEBUGLEVEL = 0
    private var bigBlocks: MutableList<*>? = null
    private var readok = false
    // TODO: fix temp file issues and implement shutdown cleanup
    // -- currently no way to delete -jm

    var header: LEOHeader? = null
        private set
    /**
     * return the Directory Array for this LEOFile
     * <br></br>holds all directories (storages and streams) in correct order
     * <br></br>Directories in turn can reference streams of data
     *
     * @return
     */
    var directoryArray: StorageTable? = null
        private set
    private var fb: FileBuffer? = null
    var fileName = "New Spreadsheet"
        internal set
    internal var encryptionStorageOverage: ByteArray? = null
    internal var encryptedXLSX = false

    /**
     * get all Directories in this LEO file
     */
    val allDirectories: Array<Storage>
        get() {
            val v = directoryArray!!.allDirectories
            var s = arrayOfNulls<Storage>(v.size)
            s = v.toTypedArray() as Array<Storage>
            return s
        }

    /**  Simply setting the byte size on the book storage
     *
     * private void checkIfSmallBlockOutput(int blen) throws StorageNotFoundException {
     * Storage book;
     * try {
     * book = storageTable.getStorageByName("Workbook");
     * }catch(StorageNotFoundException e) {
     * book = storageTable.getStorageByName("EncryptedPackage");
     * }
     *
     * if (blen >= StorageTable.BIGSTORAGE_SIZE) {
     * if (book.getBlockType() == Block.SMALL) {
     * Logger.logWarn(
     * "WARNING: Modifying SmallBlock Workbook: "+ this.getFileName() +
     * ". Modifying SmallBlock-type Workbooks " +
     * " can cause corrupt output files. See: " +
     * "http://extentech.com/uimodules/docs/docs_detail.jsp?meme_id=195");
     * throw new WorkBookException("SmallBlock File Detected reading: "+this.getFileName(), WorkBookException.SMALLBLOCK_FILE);
     * // this.convertSBtoBB(book);
     * }
     * book.setActualFileSize(blen);
     * } else
     * book.setActualFileSize(StorageTable.BIGSTORAGE_SIZE); // set at the minimum
     * }
     */

    /**
     * get the bytes of the DOC file.
     */
    // get the bytes for the WordDocument substream only
    val docBlockBytes: BlockByteReader
        get() {
            val doc: Storage?
            try {
                doc = directoryArray!!.getDirectoryByName("WordDocument")
            } catch (e: StorageNotFoundException) {
                throw InvalidFileException(
                        "InvalidFileException: Not Word '97 or later version.  Unsupported file format.")
            }

            return doc!!.blockReader
        }

    /**
     * get the bytes of the XLS file.
     */
    // get the bytes for the Workbook substream only
    val xlsBlockBytes: BlockByteReader
        get() {
            var book: Storage?
            try {
                book = directoryArray!!.getDirectoryByName("Workbook")
            } catch (e: StorageNotFoundException) {
                try {
                    book = directoryArray!!.getDirectoryByName("Book")
                } catch (e1: StorageNotFoundException) {
                    Logger.logInfo("Not Excel '97 (BIFF8) or later version.  Unsupported file format.")
                    throw InvalidFileException(
                            "InvalidFileException: Not Excel '97 (BIFF8) or later version.  Unsupported file format.")
                }

            }

            return book!!.blockReader
        }

    /**
     * Get the minimum blocks that this file can encompass
     *
     * @return
     */
    val minBlocks: Int
        get() = header!!.minStreamSize / BIGBLOCK.SIZE

    /**
     * reads in the FAT sectors from the compound file using info from the LEOFile header
     *
     * @return FATSectors == one or more sectors used to define the Sector Chain Index (==File Allocation Table == FAT)
     */
    private// the Index for the FAT(which is the
    // File Allocation Table or Sector
    // Chain Array), id's of 1st 109
    // sectors or blocks
    // **** read in the FAT index ****//
    // **** First get the inital blocks before the extra DIFAT
    // sectors ****//
    // each FAT entry is a 1-based index
    // if (DEBUG)Logger.logInfo("INFO: LEOFile Got A FAT index
    // at: "+ String.valueOf(FATidx));
    // Usually caused by CHECK ByteStreamer.writeOut()
    // dlen/filler handling
    // ******* read in any Exrra DIFAT sectors *******//
    // -1 because must always have
    // room for end of sector mark
    // fails
    // at
    // 41k
    // recs
    // jpm
    // test
    // the index to the next extra DIFAT sector is the last int
    // in the chain
    val fatSectors: CompatibleVector
        get() {
            val FATSectors = CompatibleVector()
            val DIFAT = header!!.difat
            if (DEBUG)
                Logger.logInfo("FAT Blocks:\n" + Arrays.toString(DIFAT))
            var FATidx = 0
            var blockidx = 0
            val bsz = bigBlocks!!.size - 1
            while (blockidx < Math.min(DIFAT.size, MAXDIFATLEN)) {
                FATidx = DIFAT[blockidx]
                if (FATidx <= bsz) {
                    val bbd = bigBlocks!![FATidx] as BIGBLOCK
                    bbd.isDepotBlock = true
                    FATSectors.add(bbd)
                } else {
                    Logger.logErr("LEOFile.init failed. FAT Index Attempting to fetch Block past end of blocks.")
                    throw InvalidFileException(
                            "Input file truncated. LEOFile.init failed. FAT Index Attempting to fetch Block past end of blocks.")
                }
                blockidx++
            }
            val numExtraDIFATSectors = header!!.numExtraDIFATSectors
            val extraDIFATStart = header!!.extraDIFATStart + 1
            var chainIndex = extraDIFATStart
            for (x in 0 until numExtraDIFATSectors) {
                val xind = bigBlocks!![chainIndex] as BIGBLOCK
                val extraSectorBytes = ByteBuffer.wrap(xind.bytes)
                extraSectorBytes.position(0)
                extraSectorBytes.order(ByteOrder.LITTLE_ENDIAN)
                val numExtraSectorElements = Math
                        .min(header!!.numFATSectors - blockidx, IDXBLOCKSIZE)
                for (i in 0 until numExtraSectorElements) {
                    if (i < IDXBLOCKSIZE - 1) {
                        val bbloc = extraSectorBytes.int + 1
                        if (bbloc <= bsz) {
                            val bbd = bigBlocks!![bbloc] as BIGBLOCK
                            bbd.isDepotBlock = true
                            bbd.setIsExtraSector(true)
                            FATSectors.add(bbd)
                            blockidx++
                        } else {
                            Logger.logErr("LEOFile.init failed. Attempting to fetch Invalid Extra Sector Block.")
                            throw InvalidFileException(
                                    "LEOFile.init failed. Attempting to fetch Invalid Extra Sector Block.")
                        }
                    } else {
                        chainIndex = extraSectorBytes.int + 1
                    }
                }
            }
            return FATSectors
        }

    /**
     * Close the underlying IO streams for this LEOFile.
     *
     *
     * Also deletes the temp file.
     *
     * @throws Exception
     */
    @Throws(IOException::class)
    fun close() {
        if (fb != null)
            fb!!.close()
        // KSC: close out other object refs
        fb = null
        // header.getBytes().clear();
        header = null// new LEOHeader();
        if (directoryArray != null) {
            directoryArray!!.close()
            directoryArray = null
        }
        if (bigBlocks != null) {
            for (i in bigBlocks!!.indices) {
                val b = bigBlocks!![i] as BlockImpl
                b.close()
            }
            bigBlocks!!.clear()
        }
        bigBlocks = null
        // FAT= null;
    }

    /**
     * just closes the filebuffer withut clearing out buffers and storage tables
     */
    @Throws(IOException::class)
    fun closefb() {
        if (fb != null)
            fb!!.close()
        fb = null
    }

    fun shutdown() {
        try {
            this.close()
        } catch (e: Exception) {
            if (DEBUGLEVEL > 0)
                Logger.logWarn("could not close workbook cleanly.$e")
        }

    }

    /**
     * a new LEO file containing LEO archive entries
     *
     * @param String a file path containing a valid LEOfile
     */
    constructor(fname: String) {
        if (fname.indexOf(".ser") > -1) {
            this.initFromPrototype(fname)
            return
        }
        this.fileName = fname
        fb = LEOFile.readFile(fname)
        this.initWrapper(fb!!.buffer)
    }

    /**
     * Create a leo file from a prototype string/path
     *
     *
     * PROTOTYPE_LEO_ENCRYPTED
     *
     * @param fname
     */
    private fun initFromPrototype(fname: String) {
        try {
            val b = ResourceLoader.getBytesFromJar(fname) ?: throw io.starter.formats.XLS.WorkBookException(
                    "Required Class files not on the CLASSPATH.  Check location of .jar file and/or jarloc System property.",
                    WorkBookException.LICENSING_FAILED)
            val bbf = ByteBuffer.wrap(b)
            bbf.order(ByteOrder.LITTLE_ENDIAN)
            this.initWrapper(bbf)
        } catch (e: Exception) {
            throw InvalidFileException(
                    "WorkBook could not be instantiated: $e")
        }

    }

    /**
     * Instantiate a Leo file from an encrypted document.  We currently have a hack
     * in place for encrypted documents due to their having truncated bigblocks that our
     * interface does not correctly handle.
     *
     * @param a       file containing a valid LEOfile (XLS BIFF8)
     * @param whether to use a temp file
     * @param if      the file is encrypted xlsx format
     */
    constructor(fpath: File, usetempfile: Boolean, encryptedXLSX: Boolean) {
        this.encryptedXLSX = encryptedXLSX
        this.fileName = fpath.absolutePath
        fb = LEOFile.readFile(fpath, usetempfile)
        this.initWrapper(fb!!.buffer)
    }

    /**
     * a new LEO file containing LEO archive entries
     *
     * @param a       file containing a valid LEOfile (XLS BIFF8)
     * @param whether to use a temp file
     */
    constructor(fpath: File, usetempfile: Boolean) {
        this.fileName = fpath.absolutePath
        fb = LEOFile.readFile(fpath, usetempfile)
        this.initWrapper(fb!!.buffer)
    }

    /**
     * a new LEO file containing LEO archive entries
     *
     * @param a file containing a valid LEOfile (XLS BIFF8)
     */
    constructor(fpath: File, DEBUGLEVEL: Int) {
        this.fileName = fpath.absolutePath
        this.DEBUGLEVEL = DEBUGLEVEL
        fb = LEOFile.readFile(fpath)
        this.initWrapper(fb!!.buffer)
    }

    /**
     * a new LEO file containing LEO archive entries
     *
     * @param byte[] a byte array containing a valid LEOfile
     */
    constructor(bytebuff: ByteBuffer) {
        this.initWrapper(bytebuff)
    }

    /**
     * This is just removing some duplicate code from our constructors
     *
     *
     * We should add some exception handling in here!
     */
    fun initWrapper(bytebuff: ByteBuffer?) {
        var FAT = this.init(bytebuff!!)
        if (FAT != null) {
            directoryArray!!.initDirectories(bigBlocks, FAT)
            // KSC: TESTING: XLS-97
            if (DEBUG || DEBUGLEVEL > 200)
                directoryArray!!.DEBUG()
            FAT = null
            readok = true
        } else
            readok = false
    }

    fun clearAfterInit() {
        this.bigBlocks!!.clear()
    }

    /**
     * Reads in an encrypted LEO stream.  May be able to remove this and use the standard
     * constructor, we shall see.
     *
     * @param encryptedFile
     */
    fun readEncryptedFile(encryptedFile: File) {
        fb = readFile(encryptedFile)
        this.initWrapper(fb!!.buffer)
    }

    /**
     * return whether the LEOFile contains a valid workbook
     */
    fun hasWorkBook(): Boolean {
        var book: Storage?
        try {
            book = directoryArray!!.getDirectoryByName("Workbook")
        } catch (e: StorageNotFoundException) {
            try {
                book = directoryArray!!.getDirectoryByName("Book")
            } catch (e1: StorageNotFoundException) {
                return false
            }

        }

        return true
    }

    /**
     * return whether the LEOFile contains a valid doc
     */
    fun hasDoc(): Boolean {
        if (readok) {
            try {
                val doc = this.directoryArray!!
                        .getDirectoryByName("WordDocument")
            } catch (e: StorageNotFoundException) {
                return false
            }

            return true
        }
        return false
    }

    /**
     * return whether the LEOFile contains the _SX_DB_CUR Pivot Table Cache storage (required for Pivot Table)
     *
     * @return
     */
    fun hasPivotCache(): Boolean {
        if (readok) {
            try {
                val doc = this.directoryArray!!.getDirectoryByName("_SX_DB_CUR")
            } catch (e: StorageNotFoundException) {
                return false
            }

            return true
        }
        return false
    }

    /**
     *
     */
    constructor() : super() {}

    /**
     * Create a LEOFile from an input stream.  Unfortunately because of our byte backer
     * we need to write to a temporary file in order to do this
     *
     * @param stream
     */
    @Throws(IOException::class)
    constructor(stream: InputStream) {
        val target = TempFileManager.createTempFile("ExtenXLS_temp", ".leo")
        JFileWriter.writeToFile(stream, target)

        this.fileName = target.absolutePath
        fb = LEOFile.readFile(target)
        this.initWrapper(fb!!.buffer)
        target.deleteOnExit()
        target.delete()
    }

    constructor(fx: File) {
        // TODO Auto-generated constructor stub
    }

    /**
     * Creates all the necessary headers and associated storages/structures that,
     * along with the actual workbook data, makes up the document
     *
     *
     * Basic Structure
     * 1st sector or block= Header + 1st 109 sector chains (==FAT===array of sector numbers)
     * Location of FAT==DIFAT
     * Header denotes sector of "Root Directory" which in turn denotes sectors of all other directories
     * (Workbook is but one of several directories)
     * Each directory denotes the start block of any associated data into the FAT (sector chain index)
     * (if directory is >0 an <4096 in size, it's associated data is stored in miniStorage
     *
     *
     *
     *
     *
     *
     *
     *
     *
     *
     *
     *
     *
     *
     *
     *
     * Ok, so this is a little odd, it looks as if it collects all of the storages, and handles
     * all of the preprocessing necessary to create FAT etc.
     * returns a list of all the storages, presumably to be handled by ByteStreamer.
     *
     *
     *
     *
     * A little more information:
     * A list of storages is generated.  These storages do not include
     * a) the workbook storage(!) which is meant to be the first record written after receiving the byte array
     * b) individual miniFAT storages, these are all combined into one container storage that is not represented in the header
     * c) the LeoHeader, which is written to the output stream already
     *
     *
     *
     *
     * It seems as if this method needs several things
     * 1) better header comments on its functions
     * 2) break out logic to private methods to keep overall functions understandable.
     * 3) Handle file formats other than XLS (ie doc, encrypted xls, ppt, etc)
     * 4) be clear about its usage to outputStream.  Why only the 1 storage written, maybe none would be better?
     * 5) probably throw an exception
     *
     *
     * 6) WHAT IS WITH THE STORAGES THAT ARE NOT STORAGES? -- these are non-directory storages used to write out the bytes contained within in the correct order
     * for the document.  These non-directory storages include the FAT and the minFAT blocks and miniStream
     *
     * @param outputstream to write out
     * @param book         to write
     * @param book         bytes size
     * @return storages to write out in streamer
     */
    @Synchronized
    fun writeBytes(out: OutputStream, workbook_byte_size: Int): List<*> {
        /***** storages to be written - rebuild from saved storages and recreated workbook storage */
        val storages = Vector()
        // header variables
        var numFATSectors = 0 // num FAT indexes
        var rootstart = 0 // block position where root sector is located
        var numExtraDIFATSectors = 0 // num extra sectors/blocks
        var extraDIFATStart = -2 // extra sector start (-2 for none)
        var numMiniFATSectors = 0 // num short sectors/miniFAT indexes
        var miniFATStart = -2 // start block position of short sector container
        // blocks (-2 for none)
        var sbidxpos = -2 // start position of short sector idx chain (-2 for
        // none)
        var sbsz = 0 // size of short sector container
        val DIFAT: IntArray // holds the sector id chains for the 1st 109 sector index
        // blocks (DIFAT)
        val extraDIFAT: IntArray // holds the sector id chains for any sector index
        // blocks (DIFAT) > 109
        var numblocks = 1 // header is 1st

        // workbook directory
        var book: Storage? = null
        var isEncrypted = false
        try {
            book = directoryArray!!.getDirectoryByName("Workbook")
        } catch (e3: StorageNotFoundException) {
            try {
                book = directoryArray!!.getDirectoryByName("EncryptedPackage")
                isEncrypted = !true
            } catch (e1: StorageNotFoundException) {
                // this is an error state, ideally we would be throwing a
                // (non-runtime) exception?
                throw InvalidFileException(
                        "Input LEO file not valid for output")
            }

        }

        // root directory
        var rootStore: RootStorage? = null
        try {
            rootStore = directoryArray!!.rootStorage
        } catch (e2: StorageNotFoundException) {
            // this is an error state, ideally we would be throwing a
            // (non-runtime) exception?
            throw InvalidFileException(
                    "Input LEO file not valid for output")
        }

        // get existing directories + store all except workbook +
        // root (which are rebuilt)
        val e = directoryArray!!.allDirectories.elements()
        while (e.hasMoreElements()) {
            val thisStore = e.nextElement() as Storage
            if (thisStore !== book && thisStore !== rootStore) { // &&
                // thisStore!=encryptionInfo?
                storages.add(thisStore)
                if (thisStore.blockType == Block.SMALL) { // count number
                    // of miniStream
                    // blocks
                    thisStore.miniStreamStorage = true // TODO: investigate why
                    // this setting gets
                    // lost!
                    numMiniFATSectors += thisStore.blockVect!!.size
                }
            }
        }

        // if have miniStream sectors, rebuild miniFAT index
        // + convert mini blocks to one or more big blocks (=miniFAT
        // container)
        if (numMiniFATSectors > 0) {
            val sbs = buildMiniFAT(storages, numMiniFATSectors) // returns
            // two
            // non-directory
            // storages:
            // miniStream
            // Storage
            // +
            // miniFAT
            // index
            storages.add(sbs[0])
            numMiniFATSectors = sbs[1].blockVect!!.size // trap number of
            // miniStream
            // block indexes
            storages.add(sbs[1])
        }

        // we need to count up the total block array
        // count number of blocks:
        // add blocks necessary for root storage
        numblocks += Math.ceil((directoryArray!!.allDirectories.size / 4).toDouble()).toInt() // each
        // 4
        // directories==1
        // block
        // (512
        // byte
        // sector)
        // count Storage Blocks (except root and workbook)
        for (t in storages.indices) {
            val nstr = storages.get(t) as Storage // saved stores containt
            // thier original blocks
            // + padding block
            // signifying end of
            // sector/block
            if (!nstr.miniStreamStorage)
                numblocks += nstr.sizeInBlocks
        }

        // take current workbook byte size and calculate # bigblocks
        book!!.miniStreamStorage = false // if had any initially, now it's big
        // blocks

        var workbook_idx_block_size = LEOFile
                .getSizeInBlocks(workbook_byte_size, BIGBLOCK.SIZE)
        workbook_idx_block_size = Math
                .max(workbook_idx_block_size, minBlocks - 1)// ensure
        // minimum #
        // blocks
        numblocks += workbook_idx_block_size + 1// to account for end of block
        // sector
        if (isEncrypted) { // Encrypted workbooks are already set to correct
            // size
            numblocks--
            workbook_idx_block_size--
        }
        // Given amount of blocks to write, calculate id chain
        // necessary to describe
        numFATSectors = LEOFile.getNumFATSectors(numblocks) // get # of secids
        // necessary to
        // describe
        // blocks

        // now include the FAT itself
        // total number of Sectors(blocks)= number of Sectors + FAT
        // chain block(s)
        numFATSectors = getNumFATSectors(numFATSectors + numblocks)
        val FAT = IntArray(IDXBLOCKSIZE * numFATSectors)
        for (i in FAT.indices)
            FAT[i] = -1 // create byte array of -1's 128 bytes - total number
        // of sector ids

        // allocate FAT array index
        // important positions that will be recorded in the header
        // or FAT/miniFAT
        DIFAT = IntArray(Math.min(numFATSectors, MAXDIFATLEN)) // 1 up to the
        // 1st 109
        // sector index
        // blocks used
        // to describe
        // this file
        extraDIFAT = IntArray(Math.max(numFATSectors - MAXDIFATLEN, 0)) // if
        // necessary,
        // sector
        // index
        // blocks
        // in
        // excess
        // of
        // 109

        /************************ now start to lay out block positions of all directories' associated data  */
        /************************ note: order of directories in storages is order that they will be written out
         * this order is indexed in the FAT via the block indexes used for each sector
         * upon return, storages contains all the directories, their associated data, the FAT ...
         * in the calling method, the workbook records (==associated data) is written first because
         * the workbook storage is first in the FAT.
         * Eventually, want to change this so all other storages and the FAT are written first, then
         * the calling method only needs to write the workbook records.
         */
        // first start block index with workbook blocks (they will
        // be written out in ByteStreamer before all other blocks
        // except the header)
        var blockpos = 0 // initial sector position
        book.startBlock = blockpos
        for (t in 0 until workbook_idx_block_size) {
            FAT[blockpos++] = blockpos
        }
        FAT[blockpos++] = -2 // end of sector flag
        if (!isEncrypted) {
            // this is a questionable line. It is setting the storage
            // length to that of the padded bigblock.
            // causes errors in encrypted files
            book.actualFileSize = (workbook_idx_block_size + 1) * BIGBLOCK.SIZE
            if (DEBUG)
                Logger.logInfo("Workbook actual bytes: " + book.actualFileSize)
        }

        // now rest of "static" stores (summary, document summary,
        // comp obj ...)
        for (t in storages.indices) {
            val nstr = storages.get(t) as Storage
            if (!nstr.miniStreamStorage) { // miniStream blocks are handled
                // separately
                if (nstr.getName() == "miniStream") { // miniStream (short
                    // sector container)
                    miniFATStart = blockpos // this goes in rootstore
                    // startblock to denote start of
                    // short sector container
                    sbsz = nstr.actualFileSize // actual size of short
                    // sector (miniFAT)
                    // container - stored in
                    // root storage
                } else if (nstr.getName() == "miniFAT") { // miniFAT (short
                    // sector) index
                    sbidxpos = blockpos
                }
                nstr.startBlock = -2 // default, not linked to any blocks
                // 20100325 KSC: since original blocks are kept for
                // non-workbook storage,
                // and since size is not under our control, keep original
                // size - otherwise can error
                // nstr.setActualFileSize(0); // default, 0 bytes
                // record index of block position for each block of the
                // storage
                val blks = nstr.blocks
                if (blks != null) {
                    nstr.startBlock = blockpos // start of block chain
                    for (i in 0 until blks.size - 1) { // -1 to trap
                        // last block
                        // with special
                        // end-of-block
                        // flag (-2)
                        FAT[blockpos++] = blockpos // start of this store's
                        // blocks as indexed in
                        // block index
                    }
                    FAT[blockpos++] = -2 // end of sector/block flag
                    // nstr.setActualFileSize((blks.length)*512); // set file
                    // size -- see above why we can't do this at this time
                }
            }
        }

        // after other storages, add root store:
        // must rebuild after info from above as root storage is
        // built from from all other storages
        rootStore!!.startBlock = miniFATStart // root store start points to
        // miniFAT container, or -2 if
        // none
        rootStore.actualFileSize = sbsz // 0 or miniFAT container size
        // handle rootstorage blocks = data of all other directory
        // storages
        rootStore.bytes = directoryArray!!.rebuildRootStore() // now rebuild
        // rootstore from
        // all other
        // storages
        rootstart = blockpos
        for (i in 0 until rootStore.blockVect!!.size - 1) {
            FAT[blockpos++] = blockpos
        }
        FAT[blockpos++] = -2 // end of sector/block flag
        // add rootstore to list of blocks writing out
        storages.add(rootStore)

        // mark position of FAT itself within blocks (==DIFAT)
        // the 1st MAXDIFATLEN DIFAT indexes goes into header
        for (i in DIFAT.indices) {
            DIFAT[i] = blockpos + 1 // each pos is decremented in method below,
            // so increment here
            FAT[blockpos++] = -3 // mark loc of FAT with special value -3
        }

        // if any FAT sectors > MAX, goes into extra blocks
        if (numFATSectors > MAXDIFATLEN) { // handle extra blocks - need extra
            // blocks to store the indexes
            for (n in MAXDIFATLEN until numFATSectors) {
                extraDIFAT[n - MAXDIFATLEN] = blockpos + 1
                FAT[blockpos++] = -3
            }
            for (i in 0 until Math.ceil((numFATSectors - MAXDIFATLEN) / (IDXBLOCKSIZE * 1.0)).toInt())
                FAT[blockpos++] = -4 // flag for MSAT or XBB
        }

        // now that all the blocks referenced by storages such as
        // workbook and document summary,
        // plus blocks referenced by the root storage,
        // transform the blockindex to blocks & input a phantom or
        // special store so can write out
        // have been indexed, store the index blocks themselves
        val idxstore = Storage()
        idxstore.setName("IDXStorage")
        val FATSectors = getIDXBlocks(FAT) // convert each int to 4-bytes
        // in sector/block format -->
        // FAT Sectors
        idxstore.blocks = FATSectors // input the FAT Sectors to a
        // non-directory store for later writing
        storages.add(idxstore) // output FAT Sectors after all other blocks

        if (numFATSectors > MAXDIFATLEN) {
            // now build and add the extra sectors necessary to store
            // the FAT
            extraDIFATStart = blockpos - Math
                    .ceil((numFATSectors - MAXDIFATLEN) / (IDXBLOCKSIZE * 1.0)).toInt()
            val xbbstore = buildExtraDIFAT(extraDIFAT, extraDIFATStart) // create
            // block(s)
            // from
            // the
            // extraDIFAT
            numExtraDIFATSectors = xbbstore!!.blockVect!!.size
            storages.add(xbbstore)
        }

        // now that blocks have been accounted for and indexed and
        // storages and all their blocks
        // (except for workbook blocks) are stored, we can build and
        // output the header block
        // build header block:

        header = LEOHeader.getPrototype(DIFAT) // input secIdChain to header
        // loc 76
        // start sector of short sector/miniFAT index chain or -2 if
        // n/a
        header!!.miniFATStart = sbidxpos
        // the block # where the storages begin
        header!!.setRootStorageStart(rootstart)
        // number of block indexes needed to describe 1st 109
        // sectors/blocks (= DIFAT)
        header!!.numFATSectors = numFATSectors
        // number of block indexes needed to describe short
        // sectors/miniFAT sectors if any
        header!!.numMiniFATSectors = numMiniFATSectors
        // number of block indexes necessary to describe extra
        // sectors if any
        header!!.numExtraDIFATSectors = numExtraDIFATSectors
        // start of extra sector block index or -2 if n/a
        header!!.extraDIFATStart = extraDIFATStart

        // setup the header info...
        if (!header!!.init())
            throw RuntimeException("LEO File Header Not Initialized") // invalid
        // WorkBook
        // File

        // return the storages
        return storages
    }

    /**
     * build Extra Storage Sector index (extra DIFAT)
     *
     * @param difatidx - indexes of blocks > 13952 (= 109 blocks described by DIFAT * 128 indexes in each block)
     * @return
     */
    private fun buildExtraDIFAT(difatidx: IntArray?, xbbpos: Int): Storage? {
        // handle Extra blocks beyond the 109*128 blocks described
        // by the regular DIFAT (109 limit)
        if (difatidx != null) {
            val outblocks = ArrayList()
            val xbytes = ByteArray(difatidx.size * 4)
            // convert the int difatidx to bytes
            for (i in difatidx.indices) {
                val idx = ByteTools.cLongToLEBytes(difatidx[i] - 1)
                System.arraycopy(idx, 0, xbytes, i * 4, 4)
            }
            var counter = 0
            var i = 0
            while (i < xbytes.size) {
                val bl = BlockFactory.getPrototypeBlock(Block.BIG)
                        .byteBuffer
                var len = xbytes.size - i

                // handle eob
                if (len > BIGBLOCK.SIZE - 4)
                    len = BIGBLOCK.SIZE - 4
                bl.position(0)
                bl.put(xbytes, i, len)

                // handle eob
                counter++
                if (len == BIGBLOCK.SIZE - 4)
                    bl.putInt(xbbpos + counter)

                val xbbBlock = BIGBLOCK()
                xbbBlock.init(bl, 0, 0)
                outblocks.add(xbbBlock)

                i += BIGBLOCK.SIZE - 4
            }
            // create new storage so can be added to end
            val xbbstore = Storage()
            xbbstore.setName("XBBStore")
            val xblx = arrayOfNulls<Block>(outblocks.size)
            outblocks.toTypedArray()
            xbbstore.blocks = xblx
            return xbbstore
        }
        return null
    }

    /**
     * Gather up all existing smallblocks from the existing storages
     * and create contiguous bigblocks from it; since smallblocks are
     * only 128 bytes long, 4 smallblocks fit into 1 bigblock
     * while building the miniFAT Sectors,also build the miniFAT index
     * which corresponds to it
     *
     * @param storages existing Storages (except workbook and root, which are handled separately)
     * @param numsbs   total number of miniFAT blocks
     * @return Storage[] - two storage containers: 1 contains smallblocks, 2 contains smallblock index
     */
    private fun buildMiniFAT(storages: List<*>, numsbs: Int): Array<Storage> {
        // once more, now gathering the sbidx's (miniFAT index)
        val sbidx = IntArray(Math.ceil(numsbs / (IDXBLOCKSIZE * 1.0)).toInt() * IDXBLOCKSIZE)
        var smallblocks: ByteArray? = ByteArray(0) // sum up all short sectors, will be
        // subsequently converted to big
        // blocks
        for (i in sbidx.indices)
            sbidx[i] = -1 // init sbidx index
        var z = 0
        for (i in storages.indices) {
            val thisStore = storages[i] as Storage
            if (thisStore.blockType == Block.SMALL) {
                // since these are static blocks we can copy original data
                thisStore.startBlock = z // set start position of miniFATs as
                // may have changed
                // keep length as that hasn't changed
                for (j in 0 until thisStore.blockVect!!.size - 1)
                    sbidx[z++] = z
                sbidx[z++] = -2 // end of sector
                smallblocks = ByteTools
                        .append(thisStore.bytes, smallblocks)
                // 20100407 KSC: this screws up subsequent usages of book
                // after writing ... see TestCreateNewName NPE
                // thisStore.setBlocks(new Block[0]);
            }
        }

        // concatenate all miniFAT sectors to build the miniFAT or
        // short sector container
        // i.e. miniFAT blocks are concatenated and stored in big
        // block(s):
        val smallBlocksToBig = BlockFactory
                .getBlocksFromByteArray(smallblocks, Block.BIG)
        val miniFATContainer = Storage()
        miniFATContainer.setName("miniStream")
        miniFATContainer.blocks = smallBlocksToBig
        miniFATContainer.actualFileSize = smallblocks!!.size
        smallblocks = null // no need anymore
        val miniFAT = Storage()
        miniFAT.setName("miniFAT")
        val sbIDX = getIDXBlocks(sbidx) // convert each int to 4-bytes in
        // sector/block format --> BBDIX
        miniFAT.blocks = sbIDX
        return arrayOf(miniFATContainer, miniFAT)
    }

    /**
     * get a handle to a Storage within this LEO file
     */
    @Throws(StorageNotFoundException::class)
    fun getStorageByName(s: String): Storage? {
        return directoryArray!!.getDirectoryByName(s)
    }

    /**
     * read LEO file information from header.
     */
    @Synchronized
    fun init(bbuf: ByteBuffer): IntArray? {
        var pos = 0
        var FATSectors = CompatibleVector() // one or more
        // sectors which
        // hold the FAT
        // (File
        // Allocation
        // Table or
        // indexes into
        // the sectors)

        bigBlocks = ArrayList()
        val len = bbuf.limit() / BIGBLOCK.SIZE
        // get ALL BIGBLOCKS (512 byte chunks of file)
        if (DEBUG)
            Logger.logInfo("\nINIT: Total Number of bigblocks:  $len")
        for (i in 0 until len) {
            val bbd = BIGBLOCK()
            bbd.init(bbuf, i, pos)
            pos += BIGBLOCK.SIZE
            this.bigBlocks!!.add(bbd)
        }

        // Encrypted workbooks can have random overages.
        // not ideal, but store this value in LEO and get from the
        // storage if its named 'EncryptedPackage'
        val encryptionStorageOverageLen = bbuf.limit() % BIGBLOCK.SIZE
        if (encryptionStorageOverageLen > 0) {
            val filepos = len * BIGBLOCK.SIZE
            if (this.encryptedXLSX) {
                bbuf.position(filepos)
                encryptionStorageOverage = ByteArray(encryptionStorageOverageLen)
                bbuf.get(encryptionStorageOverage!!, 0, encryptionStorageOverage!!.size)
            } else {
                val bbd = BIGBLOCK()
                bbd.init(bbuf, len, pos) // filepos);
                pos += encryptionStorageOverageLen
                this.bigBlocks!!.add(bbd)
            }
        }

        /***** Read in the file header  */
        // header holds directory start sector and
        // FAT/miniFAT/extraFAT or Sector Chain index info
        header = LEOHeader() // read the LEO file header rec

        // is a valid WorkBook File?
        if (!header!!.init(bbuf))
            throw InvalidFileException(
                    this.fileName + " is not a valid OLE File.")

        if (DEBUG) {
            Logger.logInfo("Header: ")
            Logger.logInfo("numbFATSectors: " + header!!.numFATSectors)
            Logger.logInfo("numMiniFATSectors: " + header!!.numMiniFATSectors)
            Logger.logInfo("numbExtraDIFATSectors: " + header!!.numExtraDIFATSectors)
            Logger.logInfo("rootstart: " + header!!.rootStartPos)
            Logger.logInfo("miniFATStart: " + header!!.miniFATStart)
        }
        val headerblock = bigBlocks!![0] as BIGBLOCK
        headerblock.initialized = true

        /***** Read in the FAT sectors - the sectors or blocks used for the FAT (==Sector Chain Index or File Allocation Table)   */
        FATSectors = fatSectors

        /***** turn the FAT Sectors into the FAT or Sector Index Chain  */
        var blx: ByteArray? = LEOFile.getBytes(FATSectors)
        var FAT: IntArray? = null
        FAT = LEOFile.readFAT(blx!!)
        blx = null // done

        if (DEBUG)
            Logger.logInfo("FAT:\n" + Arrays.toString(FAT))

        // if (DEBUG)
        // StorageTable.writeitout(FATSectors, "FAT.dat");

        /*****  Read the Directory blocks  */
        directoryArray = StorageTable()
        directoryArray!!.init(bbuf, header, bigBlocks, FAT)
        return FAT
    }

    /**
     * @return Returns the encryptionStorageOverage.
     */
    fun getEncryptionStorageOverage(): ByteArray {
        if (encryptionStorageOverage == null)
            encryptionStorageOverage = ByteArray(0)
        return encryptionStorageOverage
    }

    companion object {

        /**
         * serialVersionUID
         */
        private const val serialVersionUID = 2760792940329331096L
        val MAXDIFATLEN = 109                    // maximum
        // DIFATLEN
        // =
        // 109;
        // if
        // more
        // sectors
        // are
        // needed
        // goes
        // into
        // extraDIFAT
        val IDXBLOCKSIZE = 128                    // number
        // of
        // indexes
        // that
        // can
        // be
        // stored
        // in
        // 1
        // block
        val DEBUG = false
        var actualOutput = 0

        /**
         * Checks whether the given byte array starts with the LEO magic number.
         */
        fun checkIsLEO(data: ByteArray, count: Int): Boolean {
            if (count < LEOHeader.majick.size)
                return false
            for (idx in LEOHeader.majick.indices)
                if (data[idx] != LEOHeader.majick[idx])
                    return false
            return true
        }

        /**
         * calculate number of FAT blocks necessary to describe compound file
         */
        internal fun getNumFATSectors(storageTotal: Int): Int {
            var nFAT = storageTotal * 4
            nFAT /= BIGBLOCK.SIZE
            val realnum = (storageTotal * 4).toFloat() / BIGBLOCK.SIZE
            if (realnum - nFAT > 0 || nFAT > MAXDIFATLEN) {
                nFAT++
            }
            return nFAT
        }

        /**
         * create the idx from the actual locations
         * of the Blocks in the outblock array
         */
        internal fun initSmallBlockIndex(newidx: IntArray, b: Block) {
            var b = b
            while (b.hasNext() && !(b === b.next())) {
                val origps = b.blockIndex
                if (origps < 0) {
                    Logger.logWarn("WARNING: LEOFile Block Not In MINIFAT vector: " + b.originalIdx)
                } else {
                    val newp = (b.next() as Block).blockIndex
                    if (false)
                        Logger.logInfo("INFO: LEOFile Initializing block index: "
                                + origps + " val: " + newp + " idxlen: "
                                + newidx.size)
                    newidx[origps] = newp
                }
                b = b.next() as Block
            }
            if (b.blockIndex >= 0)
                newidx[b.blockIndex] = -2 // end of storage
        }

        /** create the idx from the actual locations
         * of the Blocks in the outblock array
         *
         * final static int initWBBlockIndex(int[] newidx, int sz) {
         * int start = 0;
         * for(int r=newidx.length-1;r>=0;r--) {
         * if(newidx[r]!=-1) {
         * start = r+1;
         * break;
         * }
         * }
         * for(int i=0;i<sz-1></sz-1>;i++) {
         * newidx[(start+i)] = start+(i+1);
         * }
         * if(start>0)
         * newidx[(start+sz-1)] = -2; // end of storage
         * else
         * newidx[(start+sz-1)] = -2; // end of storage
         * return start;
         * }
         */

        /**
         * returns an empty new FAT which
         * accounts for the size of its own blocks
         */
        internal fun getEmptyDIFAT(totblocks: Int, numFATSectors: Int): IntArray {

            // allocate space in idx for FAT recs, StorageTable and for
            // final -2
            val bbdi = IntArray(numFATSectors + totblocks + 1)
            for (x in bbdi.indices)
                bbdi[x] = (-1).toByte().toInt()
            return bbdi
        }

        /**
         * Create the array of BB Idx blocks
         */
        internal fun getIDXBlocks(bbdidx: IntArray): Array<Block>? {
            // step through and create index recs for all Blocks
            val b = ByteArray(bbdidx.size * 4)
            var bv = 0

            // can we wrap t in an nio buffered and read?

            var t = 0
            while (t < b.size/* 20100304 KSC: why? - 4 */) {
                val bs = ByteTools.cLongToLEBytes(bbdidx[bv++])
                // for(int z=0;z<4;z++)
                b[t++] = bs[0]
                b[t++] = bs[1]
                b[t++] = bs[2]
                b[t++] = bs[3]
            }
            return BlockFactory.getBlocksFromByteArray(b, Block.BIG)
        }

        /**
         * create a FileBuffer from a file, use system property to determine
         * whether to use a temp file.
         *
         * @param file  containing XLS bytes
         * @param fpath
         * @return
         */
        fun readFile(fpath: File): FileBuffer {
            var usetempfile = false
            val tmpfu = System.getProperties()["io.starter.formats.LEO.usetempfile"] as String

            if (tmpfu != null)
                usetempfile = tmpfu.equals("true", ignoreCase = true)

            return readFile(fpath, usetempfile)
        }

        /**
         * create a FileBuffer from a file, use boolean parameter to determine
         * whether to use a temp file.
         *
         * @param file    containing XLS bytes
         * @param whether to use a tempfile
         * @return
         */
        fun readFile(fpath: File, usetempfile: Boolean): FileBuffer {
            return if (usetempfile) FileBuffer.readFileUsingTemp(fpath) else FileBuffer.readFile(fpath)
        }

        /**
         * read in a WorkBook ByteBuffer from a file path.
         *
         *
         * from here on out, we're reading pointers to the bytes on disk
         *
         *
         * access data directly on disk through the ByteBuffer.
         *
         * <br></br> By default, OpenXLS will lock open WorkBook files, to close the file after parsing and work with a
         * temporary file instead, use the following setting:
         * <br></br><br></br>
         * System.getProperties().put("io.starter.formats.LEO.usetempfile", "true");
         * <br></br><br></br>
         * IMPORTANT NOTE: You will need to clean up temp files occasionally in your user directory (temp filenames will begin
         * with "ExtenXLS_".)
         * <br></br><br></br>
         */
        fun readFile(fpath: String): FileBuffer {
            return LEOFile.readFile(File(fpath))
        }

        /**
         * returns the table of int locations
         * in the file for the BIGBLOCK linked list.
         */
        fun readFAT(vect: List<*>): IntArray {
            val data = getBytes(vect)
            return readFAT(data)
        }

        /**
         * returns the table of int locations
         * in the file for the BIGBLOCK linked list.
         */
        private fun readFAT(data: ByteArray): IntArray {
            var data = data
            val bbs = IntArray(data.size / 4)
            var pos = 0
            var i = 0
            while (i < data.size) {
                bbs[pos++] = ByteTools
                        .readInt(data[i++], data[i++], data[i++], data[i++])
            }
            data = null
            return bbs
        }

        fun getBytes(outblocks: Array<Block>): ByteArray {
            val out = ByteArrayOutputStream()
            getBytes(outblocks, out)
            return out.toByteArray()
        }

        fun getByteStream(outblocks: Array<Block>): OutputStream {
            val out = ByteArrayOutputStream()
            getBytes(outblocks, out)
            return out
        }

        /**
         * Call this method for each storage, root storage first...
         *
         * @param outblocks
         * @param out
         */
        fun getBytes(outblocks: Array<Block>, out: OutputStream) {
            val blxt = ArrayList()
            for (tx in outblocks.indices)
                blxt.add(outblocks[tx])
            getBytes(blxt, out)
        }

        /**
         * Call this method for each storage, root storage first...
         *
         * @param outblocks
         * @param out
         */
        fun getBytes(outblocks: List<*>, out: OutputStream) {
            val it = outblocks.iterator()
            var i = 0
            while (it.hasNext()) {
                val boc = it.next() as Block
                try {
                    if (boc != null) {
                        // if (LEOFile.DEBUG)Logger.logInfo("INFO:
                        // LEOFile.getBytes() getting bytes from Block: "+ i++ + ":"
                        // + boc.getBlockIndex());
                        out.write(boc.bytes)
                        i++ // count how many blocks written
                    }
                } catch (a: IOException) {
                    Logger.logWarn("ERROR: gettting bytes from blocks failed: $a")
                }

            }
        }

        /**
         * return the bytes for the blocks in a vector
         */
        fun getBytes(bbvect: List<*>?): ByteArray {
            if (bbvect == null)
                return byteArrayOf()
            var b = arrayOfNulls<Block>(bbvect.size)
            b = bbvect.toTypedArray() as Array<Block>
            return LEOFile.getBytes(b)
        }

        /**
         * return the bytes for the blocks in a vector
         */
        fun getByteStream(bbvect: List<*>?): OutputStream {
            if (bbvect == null)
                return ByteArrayOutputStream()
            var b = arrayOfNulls<Block>(bbvect.size)
            b = bbvect.toTypedArray() as Array<Block>
            return LEOFile.getByteStream(b)
        }

        /**
         * get the number of Block records
         * that this Storage needs to store its byte array
         */
        fun getSizeInBlocks(sz: Int, blocksize: Int): Int {
            var size = sz / blocksize
            val realnum = sz.toFloat() / blocksize
            if (realnum - size > 0) {
                size++
            }
            return size
        }
    }

}