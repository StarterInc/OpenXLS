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

import io.starter.OpenXLS.ImageHandle
import io.starter.formats.escher.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.util.AbstractList
import java.util.ArrayList


/**
 * **MSODrawingGroup: MS Office Drawing Group (EBh)**<br></br>
 *
 *
 * These records contain only data.
 *
 *<pre>
 *
 * offset  name        size    contents
 * ---
 * 4       rgMSODrGr    var    MSO Drawing Group Data
 *
</pre> *
 *
 *
 *
 *
 * There is only one drawing group per client document (=MSOFBTDGGCONTAINER, 0xF000 ?).
 * OfficeArtDggContainer:
 * rh (8 bytes): An OfficeArtRecordHeader structure, that specifies the header for this record. The following table specifies the subfields.
 * rh.recVer			A value that MUST be 0xF.
 * rh.recInstance		A value that MUST be 0x000.
 * rh.recType			A value that MUST be 0xF000.
 * rh.recLen			 An unsigned integer specifying the number of bytes following the header that contain document-wide file records.
 * drawingGroup (variable): An OfficeArtFDGGBlock record, that specifies document-wide information about all the drawings that are saved in the file.
 * blipStore (variable): An OfficeArtBStoreContainer record, that specifies the container for all the BLIPs that are used in all the drawings in the parent document.
 * drawingPrimaryOptions (variable): An OfficeArtFOPT record, that specifies the default properties for all drawing objects that are contained in all the drawings in the parent document.
 * drawingTertiaryOptions (variable): An OfficeArtTertiaryFOPT record, that specifies the default properties for all the drawing objects that are contained in all the drawings in the parent document.
 * colorMRU (variable): An OfficeArtColorMRUContainer record, that specifies the most recently used custom colors.
 * splitColors (variable): An OfficeArtSplitMenuColorContainer record, that specifies a container for the colors that were most recently used to format shapes.
 *
 *
 *
 *
 * Drawing groups contain drawings.  	(= numDrawings)
 * Drawings in turn contain shapes that are the objects that actually mark a page. (= numShapes)
 * --Each drawing has a collection of rules that govern the shapes in the drawing
 * Shape store their properties in a property table (MSOFBTOPT record of Msodrawing)
 * The actual pictures and images are kept in a separate collection so can load and save separately
 *
 *
 * Records that are required in the MSODrawingGroup:
 * MSOFBTDGG-		Drawing Group Record- holds total # shapes saved + last or max SPID (shapeID) + number of IDclusters(FIDCLs) + total # drawings saved
 * MSOFBTOPT-		Property Table Record- Default properties of newly created shapes (can be 0'd)
 * MSOFBTBSTORECONTAINER-
 * MSOFBTBSE-		BLIP Store Entry- holds image type, id, size, index, len of blip name ...
 *
 * @see MSODrawing
 */
class MSODrawingGroup : io.starter.formats.XLS.XLSRecord() {
    // 20070914 KSC: Save drawing recs here
    val msodrawingrecs: AbstractList<*> = ArrayList()

    var dirtyflag = false

    /**
     * loop through all the Msodrawing recs and return the next valid SPID
     *
     * @return
     */
    val nextMsoSPID: Int
        get() {
            var spid = 0
            for (i in msodrawingrecs.indices)
                spid = Math.max((msodrawingrecs[i] as MSODrawing).spid, spid)
            return spid + 1
        }

    /**
     * return SpidMax
     *
     * @return
     */
    /**
     * set SpidMax
     *
     * @param spid
     */
    var spidMax = 1024
    /**
     * return the # of Id Clusters (charts related?)
     *
     * @return
     */
    var numIdClusters = 2
        internal set
    internal var numShapes = 1
    internal var numDrawings = 1

    private val imageData = ArrayList()
    private val imageType = ArrayList()  // parallel array with imageData
    private val cRef = ArrayList()    // 20071120 KSC: keep track of reference count for image data

    /* 20070813 KSC: These prototype bytes works for both Images and Charts */
    var PROTOTYPE_BYTES = byteArrayOf(15, 0, 0, -16, 82, 0, 0, 0, 0, 0, 6, -16, 24, 0, 0, 0, 2, 4, 0, 0, 2, 0, 0, 0, 2, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 2, 0, 0, 0, 51, 0, 11, -16, 18, 0, 0, 0, -65, 0, 8, 0, 8, 0, -127, 1, 9, 0, 0, 8, -64, 1, 64, 0, 0, 8, 64, 0, 30, -15, 16, 0, 0, 0, 13, 0, 0, 8, 12, 0, 0, 8, 23, 0, 0, 8, -9, 0, 0, 16)

    /**
     * The XF record can either be a style XF or a BiffRec XF.
     */
    /*These are prototype bytes for record 0x1c1 and 0x863 that seem to accompany when there is MSODrawingGroup data*/
    var PROTOTYPE_1C1 = byteArrayOf(-63, 1, 0, 0, -128, 56, 1, 0)
    var PROTOTYPE_863 = byteArrayOf(99, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 21, 0, 0, 0, 0, 0, 0, 0, -46)

    /**
     * returns the number of *unique* images in this workbook
     *
     * @return
     */
    val numImages: Int
        get() = imageData.size

    /**
     * return the number of drawing recs
     */
    val numDrawingRecs: Int
        get() = msodrawingrecs.size

    // moved from Boundsheet + renamed
    fun addMsodrawingrec(rec: MSODrawing) {
        msodrawingrecs.add(rec)
    }

    /**
     * remove linked MsoDrawing rec from this drawing group + update image references if necessary
     * NOTE THIS IS STILL EXPERIMENTAL; MUST BE TESTED WITH A VARIETY OF SCENARIOS
     */
    fun removeMsodrawingrec(rec: MSODrawing, sheet: Boundsheet, removeObjRec: Boolean) {
        val imgIdx = rec.getImageIndex() - 1
        val refCnt = this.getCRef(imgIdx)
        val wasHeader = rec.isHeader
        if (refCnt > 0)
            this.decCRef(imgIdx)
        msodrawingrecs.remove(rec)
        this.updateRecord()    // update msodg rec
        val idx = rec.recordIndex    // chart mso's have been taken out of streamer so idx will be -1
        if (idx > -1) {
            sheet.removeRecFromVec(rec)    // remove Mso rec
            if (removeObjRec)
            // also remove associated obj rec 20080804 KSC
                sheet.removeRecFromVec(idx)    // also remove linked Obj record
        }

        if (this.msodrawingrecs.size == 0) {    // no more drawing recs, delete this msodg
            sheet.removeRecFromVec(this)
            this.workBook!!.msodg = null
            // TODO: Unsure if there are other circumstances where MsodrawingSelection should be removed ... watch out for it
            // KSC: TODO: Necessary????  Appears so for delete chart ...(WHY??????)
            val b = sheet.getSheetRec(XLSConstants.MSODRAWINGSELECTION)
            if (b != null) {
                sheet.removeRecFromVec(b)
            }
        } else {
            if (wasHeader) { // we just removed the header; set 1st one to it
                var mso: MSODrawing? = null
                for (z in 0 until this.msodrawingrecs.size) {
                    mso = this.msodrawingrecs[z] as MSODrawing
                    if (mso.sheet == sheet && mso.isShape) {
                        mso.setIsHeader()    // make this one the header rec
                        break
                    }
                }
            }
            this.wkbook!!.updateMsodrawingHeaderRec(sheet)
        }
    }

    /**
     * return the Msodrawing header record for the given sheet
     *
     * @param bs
     * @return
     */
    fun getMsoHeaderRec(bs: Boundsheet): MSODrawing? {
        for (i in msodrawingrecs.indices) {
            val msd = msodrawingrecs[i] as MSODrawing    // get index of first Msodrawing rec for this sheet
            // always: the 1st msodrawing rec for the sheet contains the # information ...
            if (msd.sheet == bs) { // got it!
                if (msd.isHeader) {
                    return msd
                } //else
                //Logger.logErr("WorkBook.updateMsodrawingHeaderRec:  Header Record should be first rec in group.");
                //				break;
            }
        }
        return null
    }


    override fun init() {
        super.init()
        data = super.getData()
    }

    // Add associated recs necessary for Msodrawing ...
    fun initNewMSODrawingGroup() {
        // add new msodg rec to stream (just before sst)
        var index = streamer!!.getRecordIndex(XLSConstants.SST.toInt())
        // add unknown record that appears just before MSODrawingGroup
        val rec = XLSRecord()
        rec.opcode = 0x1C1.toShort()
        rec.setData(this.PROTOTYPE_1C1)
        streamer!!.addRecordAt(rec, index++)
        // add MSODrawingGroup
        streamer!!.addRecordAt(this, index)
        // also need msymystery record + msoselection ...
        val b = this.workBook!!.workSheets
        for (i in b.indices) {
            val z = b[i].getIndexOf(XLSConstants.PHONETIC)
            if (z == -1) {
                val p = Phonetic()
                p.setData(p.PROTOTYPE_BYTES)
                p.opcode = XLSRecord.PHONETIC
                p.setDebugLevel(this.DEBUGLEVEL)
                p.streamer = this.streamer
                b[i].insertSheetRecordAt(p, b[i].getIndexOf(XLSConstants.SELECTION) + 1)
            }
            /* truly necessary???    		if (i==0) { // msodrawingselection only for 1st sheet???????
         		Msodrawingselection msoSelection = new Msodrawingselection();
         		msoSelection.setData(msoSelection.PROTOTYPE_BYTES);
         		msoSelection.setOpcode(XLSRecord.MSODRAWINGSELECTION);
         		msoSelection.setDebugLevel(this.DEBUGLEVEL);
         		msoSelection.setStreamer(this.getStreamer());
         		b[i].insertSheetRecordAt(msoSelection, b[i].getIndexOf(Window2.class));
    		}
*/
        }
    }


    /**
     * Parse the MSODrawingGroup bytes and generate state vars:
     * imageData, imageType, cRef
     * spidMax, numIdClusters, numDrawings, numShapes
     */
    fun parse() {
        imageData.clear()
        imageType.clear()
        cRef.clear()
        //data = getBytes();
        if (data == null)
            return  // no data!

        /*
		 * This represents the MSOFBTDGGCONTAINER record (0xFOOO) which is the Drawing Group Container
		 *
		 *
		 *
		 * The MSOFBTDGGCONTAINER Contains the following records (some are optional):
		 * 			 rh.recVer			A value that MUST be 0xF.
					 rh.recInstance		A value that MUST be 0x000.
					 rh.recType			A value that MUST be 0xF000.
					 rh.recLen			variable
		 * 		MSOFBTDGG (0xF006)		Drawing Group Record, contains number of shapes, drawings and id clusters
		 * 		MSOFBTCLSID (0xF016) 	Clipboard format (optional)
		 * 		MSOFBTOPT	(0xF00B)	Property Table Record - default props of newly created shapes; only the properties that differ from
		 * 								the per-property defaults are saved.  Format is same as Msodrawing.MSOFBTOPT format
		 *		MSOFBTCOLORMRU (0xF11A)	MRU Color swatch ...
		 *		MSOFBTSPLITMENUCOLORS (0xF11E)	MRU colors of the top-level ..split menus
		 *		MSOFBTBSTORECONTAINER (0xF001)	An array of BLIP Store Entry (BSE) Records; Each shape indexes into this array for the BLIP they use
		 *		MSOFBTBSE (0xF007)		File BLIP Store Entry Recod FBSE record; Encodes type of BLIP + size + ID + ref. count + file offset ...
		 *		MSOFBTBLIP (0xF018)
		 */
        /* BLIP TYPE ENUM
         * msoblipERROR= 0
         * msoblipUNKNOWN,
         * msoblipEMF,	// enhanced meta file
         * msoblipWMF,	// windows meta file
         * msoblipPICT	// MAC pic
         * msoblipJPEG,
         * msoblipPNG,
         * msoblipDIB,
         * msoblipFirstClient=32,	// first client-defined BLIP type
         * msoblipLastClient= 255	// last ""
         */
        try {
            var buf: ByteArray
            val bis = ByteArrayInputStream(data!!)
            while (bis.available() > 0) {
                buf = ByteArray(8)
                var read = bis.read(buf, 0, 8)
                val version = 0xF and buf[0]
                val inst = 0xFF and buf[1] shl 4 or (0xF0 and buf[0] shr 4)
                val fbt = 0xFF and buf[3] shl 8 or (0xFF and buf[2])
                val len = ByteTools.readInt(buf[4], buf[5], buf[6], buf[7])

                //io.starter.toolkit.Logger.log("fbt:"+Integer.toHexString(fbt)+";len:"+len);
                if (fbt < 0xF004)
                    continue    // under 0xF005 are container recs; we just parse the atoms for needed info ...

                // parse record denoted by fbt
                buf = ByteArray(len)
                read = bis.read(buf, 0, len)

                when (fbt) {
                    MSODrawingConstants.MSOFBTDGG    //0xf006:		// MSOFBTDGG - Drawing Group Record
                    -> {
                        //  rh.recVer			A value that MUST be 0x0.
                        // 	rh.recInstance		A value that MUST be 0x000.
                        //	rh.recType			A value that MUST be 0xF006.
                        //	rh.recLen			A value that MUST be 0x00000010 + ((head.cidcl - 1) * 0x00000008)
                        // 	head (16 bytes): An OfficeArtFDGG record, that specifies document-wide information.
                        //
                        //  Rgidcl (variable): An array of OfficeArtIDCL elements, specifying file identifier clusters that are used in the drawing. The number of elements in the array is specified by (head.cidcl – 1).
                        spidMax = ByteTools.readInt(buf[0], buf[1], buf[2], buf[3])        // maximum shape ID
                        numIdClusters = ByteTools.readInt(buf[4], buf[5], buf[6], buf[7])    // number of ID clusters
                        numShapes = ByteTools.readInt(buf[8], buf[9], buf[10], buf[11])    // total number of shapes saved
                        numDrawings = ByteTools.readInt(buf[12], buf[13], buf[14], buf[15])    // total number of drawings saved
                    }

                    MSODrawingConstants.MSOFBTBSE        //0xf007:		// File BLIP Store Entry Record (FBSE)
                    ->
                        val imgBuf = getImageBytesWithBuffer(buf)
                }// the fixed part is followed by and array of ID clusters used internally for the translation of SPIDs to MSHHSPs (Shape Handles)
                // strip buffer from data image bytes
                //						numShapes--;	// 20070914 KSC: WHY?? -- (not a real shape? -jm)
            }
        } catch (e: Exception) {
            Logger.logErr("Msodrawingroup parse error.", e)
        }

    }

    private fun getImageBytesWithBuffer(buf: ByteArray): ByteArray {
        //		 Each BLIP in the BStore is serialized to a FBSE record;
        // btWin32, btMacOS, rgbUid[16] = identifier of blip, tag, size=BLIP size in stream
        val btWin32 = buf[0].toInt()
        imageType.add(Integer.valueOf(btWin32))
        /* parse header for testing purposes
		// inst= encoded type
		int pos= 1;
		int btMac= buf[pos++];
		byte[] UID= new byte[16];
		byte[] UID2= new byte[16];
		System.arraycopy(buf, pos, UID, 0, 16);
		pos+=16;
		short tag= ByteTools.readShort(buf[pos],buf[pos+1]); pos+=2;
		int size= ByteTools.readInt(buf[pos],buf[pos+1],buf[pos+2],buf[pos+3]); pos+=4;
		int cref= ByteTools.readInt(buf[pos],buf[pos+1],buf[pos+2],buf[pos+3]); pos+=4;
		int delayOffset= ByteTools.readInt(buf[pos],buf[pos+1],buf[pos+2],buf[pos+3]); pos+=4;
		byte usage= buf[pos++];
		byte nameLen= buf[pos++];
		byte unused1= buf[pos++];
		byte unused2= buf[pos++];
		if (nameLen==0 && delayOffset==0 && buf.length > 36) {
			// BLIP record follows then
			byte[] BLIPbuf = new byte[24];
			System.arraycopy(buf, pos, BLIPbuf, 0, 24);
			int version = (0xF&BLIPbuf[0]);
		    int inst = ((0xFF&BLIPbuf[1])<<4)|(0xF0&BLIPbuf[0])>>4;
		    int fbt = ((0xFF&BLIPbuf[3])<<8)|(0xFF&BLIPbuf[2]);
		    int len = ByteTools.readInt(BLIPbuf[4], BLIPbuf[5], BLIPbuf[6], BLIPbuf[7]);
		    UID2= new byte[16];
		    System.arraycopy(buf, 8, UID2, 0, 16);
		} else if (buf.length > 36){
			Logger.logWarn("Delay? " + ((delayOffset==0)?"No":"Yes") + " Namelen=" + nameLen);
		}
		*/
        val ref = ByteTools.readInt(buf[24], buf[25], buf[26], buf[27])
        cRef.add(Integer.valueOf(ref))

        val HEADERLEN = 61
        var STARTPOS = HEADERLEN
        var BYTELEN = buf.size
        if (HEADERLEN > BYTELEN) {
            BYTELEN = 0
            STARTPOS = 0
        }
        BYTELEN -= STARTPOS

        val imgBuf = ByteArray(BYTELEN)
        System.arraycopy(buf, STARTPOS, imgBuf, 0, BYTELEN)

        imageData.add(imgBuf)
        return imgBuf
    }

    /**
     * create a new MSODrawingGroup record based upon image datas defined in imageData/imageType/cRef arrays +
     * spidMax, numDrawings, numShapes, numIdClusters
     *
     *
     * The squenence of records here are:
     * F000, F006, F001, F007(xNumImages), F00B ,F11E
     * MSOFBTDGG MSOFBTBSTORECONTAINER MSOFBTBSE (x numimages)
     */
    fun updateRecord() {
        val BSE = arrayOfNulls<MsofbtBSE>(imageData.size)    //(0xF007,1,0);
        var imageBytes: ByteArray? = null

        var bos: ByteArrayOutputStream? = null
        try {
            bos = ByteArrayOutputStream()
            for (i in imageData.indices) {
                BSE[i] = MsofbtBSE(MSODrawingConstants.MSOFBTBSE, Integer.parseInt(imageType.get(i).toString()), 2)
                BSE[i].setImageData(imageData.get(i) as ByteArray)
                BSE[i].imageType = Integer.parseInt(imageType.get(i).toString())
                BSE[i].refCount = (cRef.get(i) as Int).toInt()        // 20071120 KSC: set the reference count for this image data
                bos.write(BSE[i].toByteArray())
            }
            imageBytes = bos.toByteArray()
        } catch (e: Exception) {
            Logger.logErr("Msodrawingroup createData error.", e)
        }

        val dgg = MsofbtDgg(MSODrawingConstants.MSOFBTDGG, 0, 0)
        dgg.setSpidMax(spidMax)        // 20071113 KSC
        numDrawings = getNumDrawings() // 20100324 KSC: changed from: msoRecs.size();	// 20080908 KSC
        dgg.numDrawings = numDrawings
        numShapes = getNumShapes()    // 20080904 KSC: sum up dg's shapes
        dgg.numShapes = numShapes
        // 2008003 KSC: numIdClusters is solely a function of spidMax dgg.setNumIdClusters(numIdClusters);
        val dggBytes = dgg.toByteArray()

        val OPT = MsofbtOPT(MSODrawingConstants.MSOFBTOPT, 0, 3)
        // add the apparent basic shape options
        OPT.setProperty(MSODrawingConstants.msooptfFitTextToShape, false, false, 0x80008, null)
        OPT.setProperty(MSODrawingConstants.msooptfillColor, false, false, 0x8000041, null)
        OPT.setProperty(MSODrawingConstants.msooptlineColor, false, false, 0x8000040, null)
        val OPTBytes = OPT.toByteArray()

        /* 20070915 KSC not necessary for all msodgs*/
        val SplitMenuColors = MsofbtSplitMenuColors(MSODrawingConstants.MSOFBTSPLITMENUCOLORS, 4, 0)
        val SplitMenuColorsBytes = SplitMenuColors.toByteArray()

        var totalLength = imageBytes!!.size

        // 20080910 KSC: if no images, don't input n MSOFBTBSTORE
        var BstoreContainerBytes = ByteArray(0)
        if (totalLength > 0) {
            val BstoreContainer = MsofbtBstoreContainer(MSODrawingConstants.MSOFBTBSTORECONTAINER, imageData.size, 15)
            BstoreContainer.length = totalLength
            BstoreContainerBytes = BstoreContainer.toByteArray()

            // add up the stuff
            totalLength += OPTBytes.size +
                    SplitMenuColorsBytes.size +
                    BstoreContainerBytes.size +
                    dggBytes.size
        } else {
            // add up the stuff
            totalLength += OPTBytes.size +
                    SplitMenuColorsBytes.size +
                    dggBytes.size
        }
        val dggContainer = MsofbtDggContainer(MSODrawingConstants.MSOFBTDGGCONTAINER, 0, 15)
        dggContainer.length = totalLength

        val dggContainerBytes = dggContainer.toByteArray()

        var pos = 0
        val retData = ByteArray(totalLength + dggContainerBytes.size)

        System.arraycopy(dggContainerBytes, 0, retData, pos, dggContainerBytes.size)
        pos += dggContainerBytes.size
        System.arraycopy(dggBytes, 0, retData, pos, dggBytes.size)
        pos += dggBytes.size
        System.arraycopy(BstoreContainerBytes, 0, retData, pos, BstoreContainerBytes.size)
        pos += BstoreContainerBytes.size
        // this is the BSE array
        System.arraycopy(imageBytes, 0, retData, pos, imageBytes.size)
        pos += imageBytes.size
        // default OPT
        System.arraycopy(OPTBytes, 0, retData, pos, OPTBytes.size)
        pos += OPTBytes.size
        // 20070915 KSC not truly necessary
        System.arraycopy(SplitMenuColorsBytes, 0, retData, pos, SplitMenuColorsBytes.size)
        pos += SplitMenuColorsBytes.size

        setData(retData)
    }


    /**
     * sets the underlying image bytes
     *
     * @param bts  new image bytes
     * @param bs   Boundsheet
     * @param rec  original Msodrawing rec linked to image
     * @param name original image name (used for lookups)
     * @return
     */
    fun setImageBytes(bts: ByteArray, bs: Boundsheet, rec: MSODrawing, name: String): Boolean {
        // Find original image handle - often is different than getImageIndex due to reuse, etc. of image bytes
        val trueIdx = rec.getImageIndex() - 1    // true index into imageData and cRef arrays
        if (trueIdx < 0)
            return false

        if (imageData.size <= trueIdx)
            return false

        try {
            if (getCRef(trueIdx) > 1) {    //20080802 KSC: if referenced more than 1x, add new so don't overwrite original
                // create new image handle with new bytes
                val im = ImageHandle(bts, bs)
                // Find original image handle + fill new with original info
                var index = -1
                val imgz = bs.images
                for (i in imgz!!.indices) {
                    if (imgz[i].name == name) {
                        index = i
                        break
                    }
                }
                val origIm = imgz[index]    // get original image handle
                im.name = origIm.name        // set new with original's data
                im.shapeName = origIm.shapeName
                im.imageType = origIm.imageType
                // insert new image into sheet
                bs.insertImage(im, true)
                im.position(origIm)    // position to original
                // now remove original mso rec
                removeMsodrawingrec(rec, bs, true)
                index = imageData.size - 1    // new index
            } else
            // just set the image bytes
                imageData.set(trueIdx, bts)
        } catch (ex: Exception) {
            Logger.logErr("Msodrawingroup setImageBytes failed.", ex)
            return false
        }

        updateRecord()
        parse()
        wkbook!!.initImages()
        return true
    }


    /**
     * returns the underlying image bytes
     *
     * @param index
     * @return
     */
    fun getImageBytes(index: Int): ByteArray? {
        if (index < 0)
            return null

        if (index >= imageData.size)
            return null

        var ret: ByteArray? = null
        try {
            ret = imageData.get(index)
        } catch (ex: Exception) {
            Logger.logErr("Msodrawingroup getImageBytes error.", ex)
        }

        return ret

    }

    fun getImageType(index: Int): Int {

        return Integer.parseInt(imageType.get(index).toString())
    }

    /**
     * related to number of drawing objects (= images + charts) but unclear how the count goes; may include deleted, etc .
     *
     * @return
     */
    fun getNumDrawings(): Int {
        // 20100324 KSC: this is experimental as numDrawings do not follow any obvious logic ...
        numDrawings = 0
        // 20100511 KSC: this is not correct ...
        for (i in msodrawingrecs.indices) {
            // 20100518 KSC: try this:
            if ((msodrawingrecs[i] as MSODrawing).isHeader)
                numDrawings++
            //numDrawings= Math.max(numDrawings, ((Msodrawing)msoRecs.get(i)).getDrawingId());
        }
        //		if (numDrawings==0 && msoRecs.size() > 0) numDrawings++;
        return numDrawings
    }

    /**
     * count the number of shapes in the document; shape mso's contain a msofbtSpContainer sub-record  (TODO: is this true in every case?)
     *
     * @return
     */
    fun getNumShapes(): Int {
        /**
         * NOTE: I've never found a clear algorithm for numShapes which results in
         * matching Excel results; however,
         * it appears that numShapes can be >= to Excel's value and open correctly;
         * problems occur when numShapes are less than what Excel expects.
         * So the below is basically the maximum numShapes available and appears to
         * result in templates
         */
        /*
		numShapes= 1;
		for (int i= 0; i < msoRecs.size(); i++) {
			MSODrawing mso= ((MSODrawing) msoRecs.get(i));
			if (mso.isShape)
				numShapes++;
		}
		*/
        numShapes = msodrawingrecs.size
        return numShapes
    }

    /**
     * test to see if imageData is in imageArray
     *
     * @param imgData byte[] defining image
     * @return
     */
    protected fun containsImage(imgData: ByteArray): Int {
        var z = -1
        var i = 0
        while (i < imageData.size && z < 0) {
            if (java.util.Arrays.equals(imgData, imageData.get(0) as ByteArray))
                z = i
            i++
        }
        return z
    }

    /**
     * return the index into the imageData array for the specified image (via byte lookup)
     *
     * @param imgData image bytes
     * @return index into imageData array
     */
    fun findImage(imgData: ByteArray): Int {
        return containsImage(imgData)
    }

    /**
     * if imageData doesn't exist, add to array
     * otherwise just inc ref count
     *
     * @param imgData             byte[] defining image
     * @param imgType             type of image
     * @param bAddUnconditionally add new even if already referenced (used in setting image bytes)
     * @return index to image
     */
    fun addImage(imgData: ByteArray, imgType: Int, bAddUnconditionally: Boolean): Int {
        var n = -1
        // 20080908 KSC: done automatically numShapes++;	// 20080208 KSC: if add unconditionally, add even if imageData already exists
        if (bAddUnconditionally || (n = containsImage(imgData)) == -1) { // 20071120 KSC: it's a unique image
            imageData.add(imgData)
            imageType.add(Integer.valueOf(imgType))
            cRef.add(Integer.valueOf(1))
            n = imageData.size
        } else {    // 20071120 KSC: If not a unqiue image, just update ref count
            incCRef(n)
            n++
        }
        return n
    }

    fun clear() {
        numShapes = 0
        imageData.clear()
        imageType.clear()
    }

    /**
     * If large, MSODrawingGroup will span multiple records; merge data
     *
     * @param rec next MSODrawingGroup record in stream
     */
    fun mergeRecords(rec: MSODrawingGroup) {
        // Merge and remove continues
        if (rec.hasContinues()) {
            rec.mergeAndRemoveContinues()        // now that data is merged, get rid of continues ...
        }
        val prevData = this.bytes
        val newData = rec.bytes
        var totalData = ByteArray(newData!!.size)
        if (prevData != null) { // a simple append of the data together
            totalData = ByteArray(prevData.size + newData.size)
            System.arraycopy(prevData, 0, totalData, 0, prevData.size)
            System.arraycopy(newData, 0, totalData, prevData.size, newData.size)
        } else
            totalData = newData
        this.setData(totalData)
    }

    /**
     * show pertinent information for record
     */
    override fun toString(): String {
        val sb = StringBuffer()
        sb.append("MSODrawingGroup: numShapes=$numShapes numDrawings=$numDrawings numIdCluster=$numIdClusters spidMax=$spidMax")
        sb.append("\nNumber of drawing records=" + msodrawingrecs.size)
        if (data != null)
            sb.append(" Length of data=" + data!!.size)
        return sb.toString()
    }

    // continue handling
    fun mergeAndRemoveContinues() {
        if (!this.isContinueMerged && this.hasContinues()) super.mergeContinues()
        if (this.hasContinues()) {
            // now that data is merged, get rid of continues ...
            val it = this.continues!!.iterator()
            while (it.hasNext()) {
                val ci = it.next() as Continue
                this.streamer!!.removeRecord(ci) // remove existing continues from stream
            }
            super.removeContinues()
            this.continues = null
        }
    }

    /**
     * increment reference count for specific image data
     *
     * @param idx
     */
    protected fun incCRef(idx: Int) {
        if (idx >= 0 && idx < cRef.size) {
            val cr = (cRef.get(idx) as Int).toInt() + 1
            cRef.removeAt(idx)
            cRef.add(idx, Integer.valueOf(cr))
        } //else  20071126 KSC: it's OK, can have - indexes ...
        //Logger.logErr("Index error encountered when updating Reference Count");
    }

    /**
     * return the reference count for the specific image
     *
     * @param idx
     */
    protected fun getCRef(idx: Int): Int {
        return if (idx >= 0 && idx < cRef.size) {
            (cRef.get(idx) as Int).toInt()
        } else -1
        //Logger.logErr("MSODrawingGroup: error encountered when returning Reference Count");
    }

    /**
     * decrement the reference count for the specific image
     *
     * @param idx
     */
    protected fun decCRef(idx: Int) {
        if (idx >= 0 && idx < cRef.size) {
            val cr = (cRef.get(idx) as Int).toInt() - 1
            cRef.removeAt(idx)
            cRef.add(idx, Integer.valueOf(cr))
        } else
            Logger.logErr("MSODrawingGroup: error encountered when decrementing Reference Count")
    }

    /**
     * add a new Drawing Record based on existing drawing record
     * i.e. from CopyWorkSheet ...
     *
     * @param spidMax
     * @param rec
     */
    fun addDrawingRecord(spidMax: Int, rec: MSODrawing) {
        this.numDrawings++
        incCRef(rec.imageIndex - 1)  // increment cRef
        this.spidMax = spidMax
        this.updateRecord()        // given all information, generate appropriate bytes
    }

    /**
     * Must ensure that oridinal drawing Id for each drawing record is correct
     * Plus ensure SPID's are correct
     *
     *
     * Not default prestreaming as we need these values when we assemble sheet recs
     */
    fun prestream() {
        var j = 0
        if (dirtyflag) {
            for (i in msodrawingrecs.indices) {
                val mso = msodrawingrecs[i] as MSODrawing
                if (mso.isHeader) {
                    mso.setDrawingId(++j)
                }
            }
        }
    }

    /**
     * clear out object references in prep for closing workbook
     */
    override fun close() {
        super.close()
        for (i in msodrawingrecs.indices) {
            val m = msodrawingrecs[i] as MSODrawing
            m.close()
        }
        msodrawingrecs.clear()
    }

    companion object {

        val prototype: XLSRecord?
            get() {
                val grp = MSODrawingGroup()
                grp.opcode = XLSConstants.MSODRAWINGGROUP
                grp.setData(grp.PROTOTYPE_BYTES)
                grp.init()
                return grp
            }

        /**
         *
         */
        private val serialVersionUID = 2378100973014157878L
    }
}