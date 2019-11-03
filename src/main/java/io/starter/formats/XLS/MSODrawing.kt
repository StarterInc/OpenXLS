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

import io.starter.OpenXLS.ColHandle
import io.starter.OpenXLS.RowHandle
import io.starter.formats.escher.*
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream


/**
 * **Msodrawing: MS Office Drawing (ECh)**<br></br>
 *
 *
 * These records contain only data.
 *
 *<pre>
 *
 * offset  name        size    contents
 * ---
 * 4       rgMSODr     var    MSO Drawing Data
 *
</pre> *
 *
 *
 * The Msodrawing record represents the MSOFBTDGCONTAINER (0xF002) in the Drawing Layer (Escher) format, and contains all per-sheet
 * types of info, including the shapes themselves.  (A shape=the elemental object that composes a drawing.  All graphical figures on a
 * drawing are shapes).  With few exceptions, shapes are stored hierachically according to how they've been grouped thru the use of
 * the Draw/Group command).
 * Each Msodrawing record contains several sub-records or atoms (atoms are records that are kept inside container records; container
 * records keep atoms and other container records organized).
 * There are several such records that are important to us (these are always present - except for MSOFBTDG in subsequent recs):
 * MSOFBTDG-				Basic Drawing Info-	#shapes in this drawing; last SPID given to an SP in this Drawing Group
 * MSOFBTSPGRCONTAINER-		Patriarch shape	Container - always first MSOFBTSPGRCONTAINER in the drawing container
 * MSOFBTSPCONTAINER-		Shape Container
 * MSOFBTSP-				Shape Atom Record- SPID= shape id + a set of flags
 * MSOFBTOPT-				Property Table Record Associated with Shape Rec- holds image Name, index + many other properties
 * MSOFBTCLIENTANCHOR- 	Client Anchor rec- holds size/bounds info
 * MSOFBTCLIENTDATA-		Host-specific client data record
 *
 *
 * There are many other records or atoms that are optional and that we will omit for now.
 *
 *
 * There appears to be 1 msodrawing record per image (there is also one OBJ record per Msodrawing record).
 * The first or header msodrawing record contains MSODBTDG for # shapes
 * There is one MsodrawingGroup record per file.  This MsodrawingGroup record also contains sub-records that hold # shapes
 *
 *
 * Occasionally, when there is a lot of image data, there can be 2 MSODRAWINGGROUP objects -- continue recs will follow
 * the second MSODG only.
 *
 *
 * SPIDs are unique per drawing group, and are parceled out by the drawing group to individual drawings in blocks of 1024
 *
 * @see MSODrawingGroup
 */
// TODO: MSOFBTCLIENTANCHOR may be substituted for MSOFBTANCHOR (clipboard), MSOFBTCHILDANCHOR (if shape is a child of a group shape)
class MSODrawing : io.starter.formats.XLS.XLSRecord() {

    var PROTOTYPE_BYTES = byteArrayOf(15, 0, 4, -16, 92, 0, 0, 0, -78, 4, 10, -16, 8, 0, 0, 0, 2, 4, 0, 0, 0, 10, 0, 0, 35, 0, 11, -16, 34, 0, 0, 0, 4, 65, 2, 0, 0, 0, 5, -63, 22, 0, 0, 0, 66, 0, 108, 0, 117, 0, 101, 0, 32, 0, 104, 0, 105, 0, 108, 0, 108, 0, 115, 0, 0, 0, 0, 0, 16, -16, 18, 0, 0, 0, 2, 0, 2, 0, 0, 0, 11, 0, 0, 0, 6, 0, 0, 0, 22, 0, 75, 0, 0, 0, 17, -16, 0, 0, 0, 0)

    internal var imageIndex = -1
    /**
     * returns true if this drawing is active i.e. not deleted
     * <br></br>Note this is experimental
     *
     * @return
     */
    // TODO: also report falses if height or width is 0??
    var isActive = false
        internal set        // true if this image is Active (not deleted) - NOTE: setting Active algorithm is not definitively proven
    var name: String? = ""
        internal set
    internal var shapeName: String? = ""
    internal var clientAnchorFlag: Short = 0            // MSOFBTCLIENTANCHOR 1st two bytes - * 0 = Move and size with Cells, 2 = Move but don't size with cells, 3 = Don't move or size with cells. */
    internal var bounds: ShortArray? = ShortArray(8)    // MSOFBTCLIENTANCHOR-
    /**
     * public method that returns saved height value;
     * used when calculated method is incorrect due to row height changes ...
     *
     * @return
     */
    var height: Short = 0
        internal set
    /**
     * public methd that returns saved width value;
     * useful when calculated method is incorrect due to column width changes ...
     *
     * @return
     */
    var originalWidth: Short = 0
        internal set    // save original height and width so that if underlying row height(s) or column width(s) change, can still set dimensions correctly ...
    /**
     * retrieve the OPT rec for specific option setting
     *
     * @return
     */
    var optRec: MsofbtOPT? = null
        private set    // 20091209 KSC: save MsofbtOPT records for later updating ...
    private var secondaryoptrec: MsofbtOPT? = null
    private var tertiaryoptrec: MsofbtOPT? = null    // apparently can have secondary and tertiary obt recs depending upon version in which it was saved ...
    internal var shapeType: Short = 0                // shape type for this drawing record

    private var SPIDSEED = 1024
    /**
     * @return whether "this is the header drawing object for the sheet"
     */
    var isHeader = false
        internal set            // whether "this is the 1st Msodrawing rec in the sheet" - contains several header records
    internal var SPID = 0                        // Shape ID
    /**
     * returns the Shape Container Length for this drawing record
     * used when setting total size for the header drawing record
     *
     * @return
     */
    var spContainerLength = 0
        internal set            // this shape's container length
    internal var isShape = true        // false for Mso's which do not contain an SPCONTAINER sub-record (can be attached textbox or solver container); not included in number of shapes count
    // only applicable to header records ***
    internal var numShapes = 1            // TODO: how do we know the value for a new image????
    internal var lastSPID = SPIDSEED    // lastSPID is stored at book level so that we can track max SPIDs - useful when images have been deleted and SPIDs are not in order ...
    internal var otherSPCONTAINERLENGTH = 0    // sum of other SPCONTAINERLENGTHs from other Msodrawing recs
    /**
     * returns the solver container length for this drawing record
     * used when setting the total size for the header drawing record
     *
     * @return
     */
    var solverContainerLength = 0
        internal set    // Solver Container length
    internal var drawingId = 0                // ordinal # of this drawing record in the workbook

    /**
     * return coordinates as {X, Y, W, H} in pixels
     *
     * @param coords
     */
    /**
     * set coordinates as {X, Y, W, H} in pixels
     *
     * @param coords
     */
    //*** WHY need -4 ?????????
    //*** WHY need -2 ?????????
    /* testsing new way above		short x= (short) Math.round(getX()*(XCONVERSION);	// convert excel units to pixels
		short y= (short) Math.round(getY()*PIXELCONVERSION);	// convert points to pixels
		short w= (short) Math.round(getWidth()*WCONVERSION);	// convert excel units to pixels
		short h= (short) Math.round((calcHeight()-8)*PIXELCONVERSION);	// convert points to pixels
*/// TODO should use ABOVE CALC!!!! see getCoords ******************
    // convert pixels to excel units
    // convert pixels to points
    // convert pixels to excel units
    // convert pixels to points
    var coords: ShortArray
        get() {
            val x = Math.round((x * 256 / ColHandle.COL_UNITS_TO_PIXELS).toFloat()).toShort()
            val y = Math.round((y * 20 / (RowHandle.ROW_HEIGHT_DIVISOR - 4)).toFloat()).toShort()
            val w = Math.round((width * 256 / ColHandle.COL_UNITS_TO_PIXELS).toFloat()).toShort()
            val h = Math.round((calcHeight() * 20 / (RowHandle.ROW_HEIGHT_DIVISOR - 2)).toFloat()).toShort()
            return shortArrayOf(x, y, w, h)
        }
        set(coords) {
            val x = Math.round((coords[0] * ColHandle.COL_UNITS_TO_PIXELS / 256).toFloat()).toShort()
            val y = Math.round((coords[1] * (RowHandle.ROW_HEIGHT_DIVISOR - 4) / 20).toFloat()).toShort()
            val w = Math.round((coords[2] * ColHandle.COL_UNITS_TO_PIXELS / 256).toFloat()).toShort()
            val h = Math.round((coords[3] * (RowHandle.ROW_HEIGHT_DIVISOR - 2) / 20).toFloat()).toShort()
            setX(x.toInt())
            setY(y.toInt())
            setWidth(w.toInt())
            setHeight(h.toInt())
        }

    val col: Int
        get() = bounds!![MSODrawing.COL].toInt()

    val col1: Int
        get() = bounds!![MSODrawing.COL1].toInt()

    val row0: Int
        get() = bounds!![MSODrawing.ROW].toInt()

    var row1: Int
        get() = bounds!![MSODrawing.ROW1].toInt()
        set(row) {
            bounds[MSODrawing.ROW1] = row.toShort()
            updateClientAnchorRecord(bounds)
        }

    /**
     * get X value of upper left corner
     * units are in excel units
     *
     * @return short x value
     */
    val x: Short
        get() {
            val col = bounds!![MSODrawing.COL].toInt()
            val colOff = bounds!![MSODrawing.COLOFFSET] / 1024.0
            var x = 0.0
            for (i in 0 until col) {
                x += getColWidth(i).toDouble()
            }
            x += colOff * getColWidth(col)
            return Math.round(x).toShort()
        }

    /**
     * returns the offset within the column in pixels
     *
     * @return
     */
    val colOffset: Short
        get() {
            val colOff = bounds!![MSODrawing.COLOFFSET] / 1024.0
            val col = bounds!![MSODrawing.COL].toInt()
            val x = colOff * getColWidth(col)
            return (Math.round(x) * XCONVERSION).toShort()
        }

    /**
     * return the y position of this object in points
     *
     * @return
     */
    val y: Short
        get() {
            val row = bounds!![MSODrawing.ROW].toInt()
            var y = 0.0
            for (i in 0 until row) {
                y += getRowHeight(i)
            }
            val rowOff = bounds!![MSODrawing.ROWOFFSET] / 256.0
            y += getRowHeight(row) * rowOff
            return Math.round(y).toShort()
        }

    /**
     * calculate width based upon col#'s, coloffsets and col widths
     * units are in excel column units
     *
     * @return short width
     */
    /*
		bounds[0]= column # of top left position (0-based) of the shape
		bounds[1]= x offset within the top-left column
		bounds[2]= row # for top left corner
		bounds[3]= y offset within the top-left corner
		bounds[4]= column # of the bottom right corner of the shape
		bounds[5]= x offset within the cell  for the bottom-right corner
		bounds[6]= row # for bottom-right corner of the shape
		bounds[7]= y offset within the cell for the bottom-right corner
		*///correct????
    val width: Short
        get() {

            val col = bounds!![MSODrawing.COL].toInt()
            val colOff = bounds!![MSODrawing.COLOFFSET] / 1024.0
            val col1 = bounds!![MSODrawing.COL1].toInt()
            val colOff1 = bounds!![MSODrawing.COLOFFSET1] / 1024.0
            var w = getColWidth(col) - getColWidth(col) * colOff
            for (i in col + 1 until col1) {
                w += getColWidth(i).toDouble()
            }
            if (col1 > col)
                w += getColWidth(col1) * colOff1
            else
                w = getColWidth(col1) * (colOff1 - colOff)

            return Math.round(w).toShort()
        }

    /* col/row and offset methods */
    var colAndOffset: ShortArray
        get() = shortArrayOf(bounds!![COL], bounds!![COLOFFSET])
        set(b) {
            bounds[COL] = b[0]
            bounds[COLOFFSET] = b[1]
            updateClientAnchorRecord(bounds)
        }

    var rowAndOffset: ShortArray
        get() = shortArrayOf(bounds!![ROW], bounds!![ROWOFFSET])
        set(b) {
            bounds[ROW] = b[0]
            bounds[ROWOFFSET] = b[1]
            updateClientAnchorRecord(bounds)
        }
    /* end position methods */


    /**
     * @return
     */
    /**
     * @param phonetic
     */
    var mystery: Phonetic? = null

    /**
     * return the SPID for this drawing record
     */
    /**
     * set the SPID for this drawing record
     * used upon copyworksheet ...
     *
     * @param spid
     */
    var spid: Int
        get() = SPID
        set(spid) {
            SPID = spid
            updateSPID()
        }

    /**
     * create a new msodrawing record with the desired SPID, imageName, shapeName and imageIndex
     * bounds should also be set?
     * create correct record bytes
     *
     * @param spid
     * @param imageName
     * @param shapeName
     * @param imageIndex
     * @return
     */
    fun createRecord(spid: Int, imageName: String?, shapeName: String?, imageIndex: Int): ByteArray {
        this.name = imageName
        this.shapeName = shapeName
        this.imageIndex = imageIndex

        var retData: ByteArray
        // Order of Msodrawing required records:
        /*		MSOFBTDG
         * 		MSOFBTSPGRCONTAINER
         * 			MSOFBTSPCONTAINER
         * 			MSOFBTSP
         * 			MSOFBTOPT
         * 			MSOFBTCLIENTANCHOR (MSOFBTCHILDANCHOR, MSOFBTANCHOR)
         * 			MSOFBTCLIENTDATA
         *
         * 		NOTE that every container has a length field that = the sum of the length of all the atoms (records) it contains as well as the length
         *      of its header
         */
        // key sub-records to update:
        // MSOFBTSP, MSOFBTOPT (image index, shape name, image name), CLIENTANCHOR if present - bounds
        // plus container which must be calculated from it's sub-records or atoms
        this.SPID = spid
        //this.SPID = SPIDSEED + imageIndex;// algorithm is incorrect for instances when images are deleted or out of order ...
        //    	Following are present in all msodrawing: MSOFBTSPCONTAINER MSOFBTSP MSOFBTOPT MSOFBTCLIENTANCHOR MSOFBTCLIENTDATA  -- not true! can have CHILDANCHOR, ANCHOR ...
        //Shape Atom; shape type must be msosptPictureFrame = 75
        val msofbtSp1 = MsofbtSp(MSODrawingConstants.MSOFBTSP, shapeType.toInt(), 2)
        msofbtSp1.setId(SPID)
        msofbtSp1.setGrfPersistence(2560)               //flag= hasSp type + has anchor -- usual for shape-type msoFbtSp's
        val msofbtSp1Bytes = msofbtSp1.toByteArray()


        // OPT= picture options
        optRec = MsofbtOPT(MSODrawingConstants.MSOFBTOPT, 0, 3)    //version is always 3, inst is current count of properties.
        if (imageIndex != -1)
            optRec!!.imageIndex = imageIndex
        if (imageName != null && imageName != "")
            optRec!!.imageName = imageName
        if (shapeName != null && shapeName != "")
            optRec!!.shapeName = shapeName

        val msofbtOPT1Bytes = optRec!!.toByteArray()

        // Client Anchor==Bounds
        val msofbtClientAnchor1 = MsofbtClientAnchor(MSODrawingConstants.MSOFBTCLIENTANCHOR, 0, 0)
        msofbtClientAnchor1.setBounds(bounds)
        val msofbtClientAnchor1Bytes = msofbtClientAnchor1.toByteArray()

        val msofbtClientData1 = MsofbtClientData(MSODrawingConstants.MSOFBTCLIENTDATA, 0, 0)  //This is an empty record
        val msofbtClientData1Bytes = msofbtClientData1.toByteArray()

        spContainerLength = msofbtSp1Bytes.size + msofbtOPT1Bytes.size + msofbtClientAnchor1Bytes.size + msofbtClientData1Bytes.size

        // 20100412 KSC: must count "oddball" msofbtClientTextBox record that follows this Mso's obj record ...
        if (shapeType.toInt() == MSODrawingConstants.msosptTextBox) {
            spContainerLength += 8
        }
        val msofbtSpContainer1 = MsofbtSpContainer(MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15)
        msofbtSpContainer1.length = spContainerLength
        val msofbtSpContainer1Bytes = msofbtSpContainer1.toByteArray()
        spContainerLength += +msofbtSpContainer1Bytes.size    // include this rec
        retData = ByteArray(spContainerLength)

        var pos = 0
        System.arraycopy(msofbtSpContainer1Bytes, 0, retData, pos, msofbtSpContainer1Bytes.size)
        pos += msofbtSpContainer1Bytes.size
        System.arraycopy(msofbtSp1Bytes, 0, retData, pos, msofbtSp1Bytes.size)
        pos += msofbtSp1Bytes.size
        System.arraycopy(msofbtOPT1Bytes, 0, retData, pos, msofbtOPT1Bytes.size)
        pos += msofbtOPT1Bytes.size
        System.arraycopy(msofbtClientAnchor1Bytes, 0, retData, pos, msofbtClientAnchor1Bytes.size)
        pos += msofbtClientAnchor1Bytes.size
        System.arraycopy(msofbtClientData1Bytes, 0, retData, pos, msofbtClientData1Bytes.size)
        pos += msofbtClientData1Bytes.size    // 20100420 KSC: empty client data record- necessary???

        if (isHeader) {    //This is only present in the first msodrawing per sheet
            if (lastSPID < SPID)
                lastSPID = SPID        // TODO: Shouldn't assume to be SPID
            var totalSPRECORDS = 0

            // Header also contains Shape Id Seed SP record
            val msofbtSp = MsofbtSp(MSODrawingConstants.MSOFBTSP, MSODrawingConstants.msosptMin, 2)  //1st shape rec is of type MSOSPTMIN
            msofbtSp.setId(SPIDSEED)                      // SPMIN==SPIDSEED
            msofbtSp.setGrfPersistence(5)                    // flag= fPatriarch,
            val msofbtSpBytes = msofbtSp.toByteArray()

            val msofbtSpgr = MsofbtSpgr(MSODrawingConstants.MSOFBTSPGR, 0, 1)
            msofbtSpgr.setRect(0, 0, 0, 0)
            val msofbtSpgrBytes = msofbtSpgr.toByteArray()
            totalSPRECORDS = msofbtSpgrBytes.size + msofbtSpBytes.size

            val msofbtSpContainer = MsofbtSpContainer(MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15)
            msofbtSpContainer.length = totalSPRECORDS
            val msofbtSpContainerBytes = msofbtSpContainer.toByteArray()
            totalSPRECORDS += msofbtSpContainerBytes.size

            spContainerLength += totalSPRECORDS
            val msofbtSpgrContainer = MsofbtSpgrContainer(MSODrawingConstants.MSOFBTSPGRCONTAINER, 0, 15)
            msofbtSpgrContainer.length = spContainerLength + otherSPCONTAINERLENGTH
            val msofbtSpgrContainerBytes = msofbtSpgrContainer.toByteArray()

            val msofbtDg = MsofbtDg(MSODrawingConstants.MSOFBTDG, drawingId, 0)
            msofbtDg.setNumShapes(numShapes)                //Number of images and drawings.
            msofbtDg.setLastSPID(lastSPID)                    //lastSPID
            val msofbtDgBytes = msofbtDg.toByteArray()

            val msofbtDgContainer = MsofbtDgContainer(MSODrawingConstants.MSOFBTDGCONTAINER, 0, 15)
            msofbtDgContainer.length = HEADERRECLENGTH + spContainerLength + otherSPCONTAINERLENGTH
            val msofbtDgContainerBytes = msofbtDgContainer.toByteArray()

            val headerRec = ByteArray(80 + retData.size) //+retData.length];

            pos = 0
            System.arraycopy(msofbtDgContainerBytes, 0, headerRec, pos, msofbtDgContainerBytes.size)
            pos += msofbtDgContainerBytes.size
            System.arraycopy(msofbtDgBytes, 0, headerRec, pos, msofbtDgBytes.size)
            pos += msofbtDgBytes.size
            System.arraycopy(msofbtSpgrContainerBytes, 0, headerRec, pos, msofbtSpgrContainerBytes.size)
            pos += msofbtSpgrContainerBytes.size
            System.arraycopy(msofbtSpContainerBytes, 0, headerRec, pos, msofbtSpContainerBytes.size)
            pos += msofbtSpContainerBytes.size
            System.arraycopy(msofbtSpgrBytes, 0, headerRec, pos, msofbtSpgrBytes.size)
            pos += msofbtSpgrBytes.size
            System.arraycopy(msofbtSpBytes, 0, headerRec, pos, msofbtSpBytes.size)
            pos += msofbtSpBytes.size

            System.arraycopy(retData, 0, headerRec, pos, retData.size)
            retData = headerRec

        }

        this.setData(retData)
        this.length = data!!.size
        return retData
    }


    /**
     * parse the data contained in this drawing record
     */
    override fun init() {
        // *************************************************************************************************************************************
        // 20070910 KSC: parse MSO record + atom ids (records keep atoms and other containers; atoms contain info and are kept inside containers)
        // common record header for both:  ver, inst, fbt, len; fbt deterimes record type (0xF000 to 0xFFFF)
        // for a specific record, inst differentiates atoms
        // for atoms, ver= version; for records, ver= 0xFFFF
        // for atoms, len= of atom excluding header; for records; sum of len of atoms contained within it

        /*	an Msodrawing record may contain the following records and atoms:
		 * 	MSOFBTDG					// drawing record: count + MSOSPID seed
        	MSOFBTREGROUPITEMS			// Mappings to reconstitute groups (for regrouping)
        	MSOFBTCOLORSCHEME
        	MSOFBTSPGRCONTAINER			// Group Shape Container
            MSOFBTSPCONTAINER			// Shape Container
        		MSOFBTSPGR				// Group-shape-specific info  (i.e. shapes that are groups) optional
        	**	MSOFBTSP				// A Shape atom record **
		 	**	MSOFBTOPT				// The Property Table for a shape ** - image index + name ...
            	MSOFBTTEXTBOX			// if the shape has text
            	MSOFBTCLIENTTEXTBOX		// for clipboard stream
            	MSOFBTANCHOR			// Anchor or location fo a shape (if streamed to a clipboard) optional
            	MSOFBTCHILDANCHOR		//   " ", if shape is a child of a group shape optional
            **  MSOFBTCLIENTANCHOR		// Client Anchor/Bounds **
            	MSOFBTCLIENTDATA		// content is determined by host optional
            	MSOFBTOLEOBJECT			// optional
            	MSOFBTDELETEDPSPL		// optional
		 */
        // *************************************************************************************************************************************
        super.init()

        val bis = ByteArrayInputStream(super.getData()!!)
        var version: Int
        var inst: Int
        var fbt: Int
        var len: Int
        spContainerLength = 0        // this shape container length
        otherSPCONTAINERLENGTH = 0    // if header, all other SPCONTAINERLENGTHS -- calc from DGCONTAINERLENGTH + SPCONTAINERLENGTHS

        var SPGRCONTAINERLENGTH = 0        // group shape container length
        var SPCONTAINERATOMS = 0        // atoms or sub-records which make up the shape container
        var DGCONTAINERLENGTH = 0        // drawing container length
        var DGCONTAINERATOMS = 0        // atoms or sub-records which, along with SPGRCONTAINER + SOLVERCONTAINER(s), make up the DGCONTAINERLENGHT
        var SOLVERCONTAINERATOMS = 0    // atoms or sub-records which make up the SOLVERCONTAINERLENGTH
        var hasUndoInfo = false        // true if a non-header Mso contains an SPGRCONTAINER - documentation states:   Shapes that have been deleted but that could be brought back via Undo.
        while (bis.available() > 0) {
            var dat = ByteArray(8)
            bis.read(dat, 0, 8)
            version = 0xF and dat[0]    // 15 for containers, version for atoms
            inst = 0xFF and dat[1] shl 4 or (0xF0 and dat[0] shr 4)
            fbt = 0xFF and dat[3] shl 8 or (0xFF and dat[2])    // record type id
            len = ByteTools.readInt(dat[4], dat[5], dat[6], dat[7])    // for atoms, record length - header length (=8), if container, refers to sum of lengths of atoms inside it, incl. record headers
            if (version == 15) {// do not parse containers
                // MSOFBTSPGRCONTAINER:		// Shape Group Container, contains a variable number of shapes (=msofbtSpContainer) + other groups 0xF003
                // MSOFBTSPCONTAINER:		// Shape Container 0xF004
                // may have several SPCONTAINERs, 1 for background shape, several for deleted shapes ...
                // possible containers
                // DGCONTAINER= DG, REGROUPITEMS, ColorSCHEME, SPGR, SPCONTAINER
                // SPGRCONTAINER= SPCONTAINER(s)
                // SPCONTAINER= SPGR, SP, OPT, TEXTBOX, ANCHOR, CHILDANCHOR, CLIENTANCHOR, CLIENTDATA, OLEOBJECT, DeletedPSPL
                // SOLVERCONTAINER= ConnetorRule, AlignRule, ArcRule, ClientRule, CalloutRule,
                if (fbt == MSODrawingConstants.MSOFBTDGCONTAINER) {
                    isHeader = true
                    otherSPCONTAINERLENGTH = len    //-HEADERRECLENGTH;
                    DGCONTAINERLENGTH = len
                } else if (fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER) {//patriarch shape, with all non-background non-deleted shapes in it - may have more than 1 subrecord
                    // A group is a collection of other shapes. The contained shapes are placed in the coordinate system of the group.
                    // The group container contains a variable number of shapes (msofbtSpContainer) and other groups (msofbtSpgrContainer, for nested groups).
                    // The group itself is a shape, and always appears as the first msofbtSpContainer in the group container.
                    if (SPGRCONTAINERLENGTH == 0)
                    // only add 1st container length - others are deleted shapes ...
                        SPGRCONTAINERLENGTH = len
                    if (!isHeader)
                    // then this is a grouped shape, must add the header length as is apparent in existing container lengths (see below)
                        hasUndoInfo = true
                } else if (fbt == MSODrawingConstants.MSOFBTSPCONTAINER) {
                    spContainerLength += len + 8 //	  add 8 for record header
                    if (isHeader)
                    // keep track of total "other sp container length" - necessary to calculate SPGRCONTAINERLENGTH + DGCONTAINERLENGTH
                        otherSPCONTAINERLENGTH -= len + 8
                    isShape = true    // any mso that contains a normal "spcontainer" is a shape
                } else if (fbt == MSODrawingConstants.MSOFBTSOLVERCONTAINER) {    // solver container: rules governing shapes
                    isShape = false
                    solverContainerLength = len + 8    // added to dgcontainerlength
                } else {
                    Logger.logInfo("MSODrawing.init: unknown container encountered: $fbt")
                }
                continue
            }
            // parse atoms or sub-records (distinct from container records above)
            dat = ByteArray(len)
            bis.read(dat, 0, len)
            when (fbt) {
                MSODrawingConstants.MSOFBTSP            // A shape atom rec (inst= shape type) rec= shape ID + group of flags
                -> {
                    val fGroup: Boolean
                    val fChild: Boolean
                    val fPatriarch: Boolean
                    val fDeleted: Boolean
                    val fOleShape: Boolean
                    val fHaveMaster: Boolean
                    val fFlipH: Boolean
                    val fFlipV: Boolean
                    val fConnector: Boolean
                    val fHaveAnchor: Boolean
                    val fBackground: Boolean
                    val fHaveSpt: Boolean
                    val flag: Int
                    SPID = ByteTools.readInt(dat[0], dat[1], dat[2], dat[3])
                    flag = ByteTools.readInt(dat[4], dat[5], dat[6], dat[7])
                    // parse out flag:
                    fGroup = flag and 0x1 == 0x1        // it's a group shape
                    fChild = flag and 0x2 == 0x2        // it's not a top-level shape
                    fPatriarch = flag and 0x4 == 0x4    // topmost group shape ** 1 per drawing *8
                    fDeleted = flag and 0x8 == 0x8        // had been deleted
                    // 4
                    fOleShape = flag and 0x10 == 0x10    // it's an OLE shape
                    fHaveMaster = flag and 0x20 == 0x20    // it has a master prop
                    fFlipH = flag and 0x40 == 0x40        // it's flipped horizontally
                    fFlipV = flag and 0x80 == 0x80        // it's flipped vertically
                    // 8
                    fConnector = flag and 0x100 == 0x100    // it's a connector type
                    fHaveAnchor = flag and 0x200 == 0x200    // it's an anchor type
                    fBackground = flag and 0x400 == 0x400    // it's a background shape
                    fHaveSpt = flag and 0x800 == 0x800    // it has a shape-type property
                    // FYI: there are normally two msofbtsp records for each drawing
                    // the first (flag==5) defines fGroup + fPatriarch,
                    // with inst==0, apparently shape type to define SPIDSEED
                    // the second msofbtsp record (flag=2560)defines the SPID and contains flags:
                    // fHaveAnchor, fHaveSpt and inst==shape type
                    //NOTE: setting Active algorithm is not definitively proven;
                    // if it ISN'T active, why isn't the fDeleted flag set???
                    isActive = isActive || fPatriarch    // if we have fPatriarch, it's active ...
                    if (fHaveSpt)
                        shapeType = inst.toShort()    // save shape type
                    if (inst == 0)
                    //== shape type
                        SPIDSEED = SPID        // seed+imageIndex= SPID
                    SPCONTAINERATOMS += len + 8
                }

                MSODrawingConstants.MSOFBTCLIENTANCHOR        // Anchor or location for a shape
                -> {
                    // NOT SO! sheetIndex = ByteTools.readShort(buf[0],buf[1]);
                    /**
                     * bounds[0]= column # of top left position (0-based) of the shape
                     * bounds[1]= x offset within the top-left column
                     * bounds[2]= row # for top left corner
                     * bounds[3]= y offset within the top-left corner
                     * bounds[4]= column # of the bottom right corner of the shape
                     * bounds[5]= x offset within the cell  for the bottom-right corner
                     * bounds[6]= row # for bottom-right corner of the shape
                     * bounds[7]= y offset within the cell for the bottom-right corner
                     */
                    clientAnchorFlag = ByteTools.readShort(dat[0].toInt(), dat[1].toInt())
                    bounds[COL] = ByteTools.readShort(dat[2].toInt(), dat[3].toInt())
                    bounds[COLOFFSET] = ByteTools.readShort(dat[4].toInt(), dat[5].toInt())
                    bounds[ROW] = ByteTools.readShort(dat[6].toInt(), dat[7].toInt())
                    bounds[ROWOFFSET] = ByteTools.readShort(dat[8].toInt(), dat[9].toInt())
                    bounds[COL1] = ByteTools.readShort(dat[10].toInt(), dat[11].toInt())
                    bounds[COLOFFSET1] = ByteTools.readShort(dat[12].toInt(), dat[13].toInt())
                    bounds[ROW1] = ByteTools.readShort(dat[14].toInt(), dat[15].toInt())
                    bounds[ROWOFFSET1] = ByteTools.readShort(dat[16].toInt(), dat[17].toInt())
                    SPCONTAINERATOMS += len + 8
                }

                MSODrawingConstants.MSOFBTOPT    // property table atom - for 97 and earlier versions
                -> {
                    // 20091209 KSC: save MsoFbtOpt record for later use in updating, if necessary
                    optRec = MsofbtOPT(fbt, inst, version)
                    optRec!!.setData(dat)    // sets and parses msoFbtOpt data, including imagename, shapename and imageindex
                    name = optRec!!.imageName
                    shapeName = optRec!!.shapeName
                    imageIndex = optRec!!.imageIndex
                    SPCONTAINERATOMS += len + 8
                }

                // 20100519 KSC: later versions can have secondary and tertiary opt blocks
                MSODrawingConstants.MSOFBTSECONDARYOPT        // movie (id= 274)
                -> {
                    secondaryoptrec = MsofbtOPT(fbt, inst, version)
                    secondaryoptrec!!.setData(dat)    // sets and parses msoFbtOpt data
                    SPCONTAINERATOMS += len + 8
                }

                MSODrawingConstants.MSOFBTTERTIARYOPT    // for Office versions 2000, XP, 2003, and 2007
                -> {
                    tertiaryoptrec = MsofbtOPT(fbt, inst, version)
                    tertiaryoptrec!!.setData(dat)    // sets and parses msoFbtOpt data
                    SPCONTAINERATOMS += len + 8
                }

                MSODrawingConstants.MSOFBTDG    // Drawing Record: ID, num shapes + Last SPID for this DG
                -> {
                    drawingId = inst
                    numShapes = ByteTools.readInt(dat[0], dat[1], dat[2], dat[3])  // number of shapes in this drawing
                    lastSPID = ByteTools.readInt(dat[4], dat[5], dat[6], dat[7])
                    DGCONTAINERATOMS += len + 8
                    otherSPCONTAINERLENGTH -= len + 8 // these atoms are part of DGCONTAINER, not SPCONTAINER - adjust accordingly
                }
                MSODrawingConstants.MSOFBTCOLORSCHEME, MSODrawingConstants.MSOFBTREGROUPITEMS -> {
                    DGCONTAINERATOMS += len + 8
                    otherSPCONTAINERLENGTH -= len + 8
                }

                MSODrawingConstants.MSOFBTCLIENTTEXTBOX // msofbtClientTextbox sub-record, contains only this one atom no containers ...
                    , MSODrawingConstants.MSOFBTTEXTBOX // msofbtClientTextbox sub-record, contains only this one atom no containers ...
                -> isShape = false    // is treated differently, isn't counted in numShapes calcs

                MSODrawingConstants.MSOFBTCHILDANCHOR    //  used for all shapes that belong to a group. The content of the record is simply a RECT in the coordinate system of the parent group shape
                    ,
                    //  If the shape is saved to a clipboard:
                MSODrawingConstants.MSOFBTSPGR    // shapes that ARE groups, not shapes that are IN groups; The group shape record defines the coordinate system of the shape
                    ,
                    //  If the shape is a child of a group shape:
                MSODrawingConstants.MSOFBTANCHOR        // used for top-level shapes when the shape streamed to the clipboard. The content of the record is simply a RECT with a coordinate system of 100,000 units per inch and origin in the top-left of the drawing
                    , MSODrawingConstants.MSOFBTCLIENTDATA, MSODrawingConstants.MSOFBTDELETEDPSPL -> SPCONTAINERATOMS += len + 8

                MSODrawingConstants.MSOFBTCONNECTORRULE, MSODrawingConstants.MSOFBTALIGNRULE, MSODrawingConstants.MSOFBTARCRULE, MSODrawingConstants.MSOFBTCLIENTRULE, MSODrawingConstants.MSOFBTCALLOUTRULE -> SOLVERCONTAINERATOMS += len + 8

                else -> Logger.logInfo("MSODrawing.init:  unknown subrecord encountered: $fbt")
            }
        }
/* //DEBUGGING:  THESE CONTAINER LENGTH CALCULATIONS PASS FOR ALL thus far MSO's ENCOUNTERED
		boolean diff= false;
		if (isHeader()) {
			int d0= (DGCONTAINERLENGTH-(DGCONTAINERATOMS+SPGRCONTAINERLENGTH+SOLVERCONTAINERLENGTH+8));
			int d1= (SPGRCONTAINERLENGTH+8-(SPCONTAINERLENGTH+otherSPCONTAINERLENGTH));
			if (d0+d1!=0) {
				if (DGCONTAINERLENGTH!=(DGCONTAINERATOMS+SPGRCONTAINERLENGTH+SOLVERCONTAINERLENGTH+8)) {  // this may not be 100% since must account for OTHER record's SOLVERCONTAINER LENGTHS
					io.starter.toolkit.Logger.log("DGCONTAINERLENGTH DIFF: " + (DGCONTAINERLENGTH-(DGCONTAINERATOMS+SPGRCONTAINERLENGTH+SOLVERCONTAINERLENGTH+8)));
					diff= true;
				}
				// ******* sum of SPCONTAINERS ***************
				if (SPGRCONTAINERLENGTH+8!=SPCONTAINERLENGTH+otherSPCONTAINERLENGTH) {
					io.starter.toolkit.Logger.log("SPGRCONTAINERLENGTH DIFF: " + (SPGRCONTAINERLENGTH+8-(SPCONTAINERLENGTH+otherSPCONTAINERLENGTH)));
					diff= true;
				}
			}
		}
		// one or two header lengths (8 bytes each) have been added to SPCONTAINERLENGTH
		// adjust here:
		int headerlens= 0;
		if (isHeader())
			headerlens= 16;
		else if (isShape) 	// non-shapes don't have SPGRCONTAINERS
			headerlens= 8;
		if (isHeader() && SPCONTAINERLENGTH==48) // only one SPCONTAINER, decrement by 1 header length
			headerlens= 8;
		if (optrec!=null && optrec.hasTextId()) // shape has an attached text box; must add 8 for following CLIENTTEXTBOX (the !isShape Mso which follows ...)
			SPCONTAINERATOMS+=8;
		if (SPCONTAINERLENGTH-headerlens!=SPCONTAINERATOMS) {
			System.err.println("SPCONTAINERLEN IS OFF: " + (SPCONTAINERLENGTH-headerlens-SPCONTAINERATOMS));
			System.err.println(this.toString());
			System.err.println(this.debugOutput());
			diff= true;
		}
		/**/
        if (hasUndoInfo)
spContainerLength += 8        // Shapes that have been deleted but that could be brought back via Undo. -- must add to sp container length for total container length calc (see UpdateHeader)
}

/**
 * update the existing mso with the appropriate basic mso data
 * <br></br>NOTE: To set other mso data, see setShapeType, setIsHeader ...
 *
 * @param spid       Unique Shape Id
 * @param imageName  String image name or null
 * @param shapeName  String shape name or null
 * @param imageIndex int index into the image byte store (for a picture-type) or -1 for null
 * @param bounds     short[8] the position of this Mso given in rows, cols and offsets
 */
     fun updateRecord(spid:Int, imageName:String?, shapeName:String?, imageIndex:Int, bounds:ShortArray) {
this.SPID = spid
this.name = imageName
this.shapeName = shapeName
this.imageIndex = imageIndex
this.bounds = bounds
updateRecord()
}

/**
 * rebuild record bytes for this Msodrawing
 * aside from updating significant atoms such as MSOFBTOPT, it also recalculates container lengths
 *
 * @param spid
 * @return
 */
     fun updateRecord() {
 /*// debug: check algorithm:
	io.starter.toolkit.Logger.log(this.toString());
	io.starter.toolkit.Logger.log(this.debugOutput());
	int origSP= SPCONTAINERLENGTH;
	int origDG= 0;
	int origSPGR= 0;
/**/
        var spcontainer1atoms = ByteArray(0)
var spcontainer2atoms = ByteArray(0)    // header specific
var dgcontaineratoms = ByteArray(0)    // header specific

var hasUndoInfo = false        // true if a non-header Mso contains an SPGRCONTAINER - documentation states:   Shapes that have been deleted but that could be brought back via Undo.
 // reset state vars
        spContainerLength = 0        // this shape container length

 // first pass, update significant atoms and sum up container lengths
        val fbt:Int
val len:Int
super.getData()
val bis = ByteArrayInputStream(data!!)
while (bis.available() > 0)
{
val header = ByteArray(8)
bis.read(header, 0, 8)
fbt = ((0xFF and header[3]) shl 8) or (0xFF and header[2])
len = ByteTools.readInt(header[4], header[5], header[6], header[7])
if ((0xF and header[0]) == 15)
{    // 15 for containers, version for atoms
 // most containers, just skip; some, however, are necessaary for container length calcs (see below
                if (fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER)
{
 //if (origSPGR==0) origSPGR= len;
                    if (!isHeader)
 // then this is a grouped shape, must add the header length as is apparent in existing container lengths (see below)
                        hasUndoInfo = true
}
else if (fbt == MSODrawingConstants.MSOFBTSPCONTAINER)
{
isShape = true    // any mso that contains a normal "spcontainer" is a shape
}
else if (fbt == MSODrawingConstants.MSOFBTSOLVERCONTAINER)
{    // solver container: rules governing shapes
 // TODO: is there EVER a reason to update a SOLVERCONTAINER RECORD???
                    // STRUCTURE:
                    // MSOFBTSOLVERCONTAINER 61445 15/#/# (15/0/0 for an empty solvercontainer)
                    // then 0 or more rules:
                    // MSOFBTCONNECTORRULE 61458 1/0/24 ... etc
                    solverContainerLength = len + 8    // added to dgcontainerlength
isShape = false
 // testing- remove when done
                }
else if (fbt == MSODrawingConstants.MSOFBTDGCONTAINER)
;//origDG= len;
}
else
{
 // parse atoms or sub-records (distinct from container records above)
                var data = ByteArray(len)
bis.read(data, 0, len)
when (fbt) {
MSODrawingConstants.MSOFBTDG    // update drawing record atom
 -> {
System.arraycopy(ByteTools.cLongToLEBytes(this.numShapes), 0, data, 0, 4)
System.arraycopy(ByteTools.cLongToLEBytes(this.lastSPID), 0, data, 4, 4)
len = data.size
data = ByteTools.append(data, header)
dgcontaineratoms = ByteTools.append(data, dgcontaineratoms)
}

MSODrawingConstants.MSOFBTSP            // A shape atom rec (inst= shape type) rec= shape ID + group of flags
 -> {
 // TODO: necessary to ever update the SPIDSEED?
                        val flag = ByteTools.readInt(data[4], data[5], data[6], data[7])
val fHaveSpt = (flag and 0x800) == 0x800    // it has a shape-type property
if (flag != 5)
{    // if it's not the SPIDseed, update SPID
System.arraycopy(ByteTools.cLongToLEBytes(SPID), 0, data, 0, 4)
}
if (fHaveSpt)
{    // shape type is contained within inst var
header[0] = ((0xF and 2) or (0xF0 and (shapeType shl 4))).toByte()
header[1] = ((0x00000FF0 and shapeType) shr 4).toByte()
}
data = ByteTools.append(data, header)
if (flag != 5)
 // if it's not the header
                            spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
else
spcontainer2atoms = ByteTools.append(data, spcontainer2atoms)
}

MSODrawingConstants.MSOFBTCLIENTANCHOR        // Anchor or location for a shape
 -> {
 // udpate bounds
                        System.arraycopy(ByteTools.shortToLEBytes(clientAnchorFlag), 0, data, 0, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![0]), 0, data, 2, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![1]), 0, data, 4, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![2]), 0, data, 6, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![3]), 0, data, 8, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![4]), 0, data, 10, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![5]), 0, data, 12, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![6]), 0, data, 14, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![7]), 0, data, 16, 2)
data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTOPT    // property table atom - for 97 and earlier versions
 -> {
 // OPT= picture options
                        if (optRec == null)
 // if do not have a MsoFbtOPT record, then create new -- shouldn't get here as should always have an existing record
                            optRec = MsofbtOPT(MSODrawingConstants.MSOFBTOPT, 0, 3)    //version is always 3, inst is current count of properties.
if (imageIndex != optRec!!.getImageIndex())
optRec!!.setImageIndex(imageIndex)
if (name == null || name != optRec!!.getImageName())
optRec!!.setImageName(name)
if (shapeName == null || shapeName != optRec!!.getShapeName())
{
optRec!!.setShapeName(shapeName)
}
data = optRec!!.toByteArray()
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTTERTIARYOPT    // for Office versions 2000, XP, 2003, and 2007
, MSODrawingConstants.MSOFBTSECONDARYOPT        // movie (id= 274)
 -> {
data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTCOLORSCHEME, MSODrawingConstants.MSOFBTREGROUPITEMS -> {
data = ByteTools.append(data, header)
dgcontaineratoms = ByteTools.append(data, dgcontaineratoms)
}

MSODrawingConstants.MSOFBTCLIENTTEXTBOX // msofbtClientTextbox sub-record, contains only this one atom no containers ...
, MSODrawingConstants.MSOFBTTEXTBOX // msofbtClientTextbox sub-record, contains only this one atom no containers ...
 -> {
 // DON"T ADD LEN **************
                        isShape = false
data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTSPGR    // shapes that ARE groups, not shapes that are IN groups; The group shape record defines the coordinate system of the shape
 -> {
data = ByteTools.append(data, header)
spcontainer2atoms = ByteTools.append(data, spcontainer2atoms)
}

MSODrawingConstants.MSOFBTCHILDANCHOR    //  used for all shapes that belong to a group. The content of the record is simply a RECT in the coordinate system of the parent group shape
, MSODrawingConstants.MSOFBTANCHOR        // used for top-level shapes when the shape streamed to the clipboard. The content of the record is simply a RECT with a coordinate system of 100,000 units per inch and origin in the top-left of the drawing
, MSODrawingConstants.MSOFBTCLIENTDATA, MSODrawingConstants.MSOFBTDELETEDPSPL -> {
data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTCONNECTORRULE, MSODrawingConstants.MSOFBTALIGNRULE, MSODrawingConstants.MSOFBTARCRULE, MSODrawingConstants.MSOFBTCLIENTRULE, MSODrawingConstants.MSOFBTCALLOUTRULE -> {}

else -> {
Logger.logInfo("MSODrawing.updateRecord:  unknown subrecord encountered: " + fbt)
data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}
}// SOLVERCONTAINERATOMS+=len+8;
 // TODO: is there EVER a reason to update a SOLVERCONTAINER RECORD???
 //Logger.logInfo("MSODrawing.updateRecord:  encountered solver container atom");
}
}
 // container lengths:
        // are these adjustments necessary??
 /*
	  	if (optrec!=null && optrec.hasTextId()) // shape has an attached text box; must add 8 for following CLIENTTEXTBOX (the !isShape Mso which follows ...)
			SPCONTAINERATOMS+=8;
*/
        spContainerLength = spcontainer1atoms.size
var additionalSP = 0
if (hasUndoInfo)
additionalSP = 8        // Shapes that have been deleted but that could be brought back via Undo. -- must add to sp container length for total container length calc (see UpdateHeader)
if (shapeType.toInt() == MSODrawingConstants.msosptTextBox)
additionalSP = 8        // account for attached text mso - which has no SPCONTAINER so must include in "controlling" rec
 // 2nd pass:  now have the important container lengths and their updated
        // subrecords/atoms, create resulting byte array

        // build main spcontainer - valid for all records
        val spcontainer1 = MsofbtSpContainer(MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15)
spcontainer1.length = spContainerLength + additionalSP
val container = spcontainer1.toByteArray()
spContainerLength += container.size    // include this header length

 /*// debugging
    	if (!bIsHeader && SPCONTAINERLENGTH!=origSP)
    		Logger.logErr("SPCONTAINERLENTH IS OFF: " + (SPCONTAINERLENGTH-origSP));
    	 */
        var retData = ByteArray(spContainerLength)
System.arraycopy(container, 0, retData, 0, container.size)
System.arraycopy(spcontainer1atoms, 0, retData, container.size, spcontainer1atoms.size)

spContainerLength += additionalSP// necessary when summing up container lengths -- see WorkBook.updateMsoHeaderRecord
if (isHeader)
{
 // SPCONTAINER
            // sp2 -- SPIDSEED
            // spgr if necessary
            val spcontainer2 = MsofbtSpContainer(MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15)
spcontainer2.length = spcontainer2atoms.size
val container2 = spcontainer2.toByteArray()
spContainerLength += spcontainer2atoms.size + container2.size    // include this header length

 /*// debugging
        	if (SPCONTAINERLENGTH!=origSP)
        		Logger.logErr("SPCONTAINERLENTH IS OFF: " + (SPCONTAINERLENGTH-origSP));
			/**/
            // SPGRCONTAINER
            val spgrcontainerlen = spContainerLength + otherSPCONTAINERLENGTH - 8
 /*// debugging
    	  	if (spgrcontainerlen!=origSPGR)
    	  		Logger.logErr("SPGRCONTAINERLENTH IS OFF: " + (spgrcontainerlen-origSPGR));
    	  	/**/
            val msofbtSpgrContainer = MsofbtSpgrContainer(MSODrawingConstants.MSOFBTSPGRCONTAINER, 0, 15)
msofbtSpgrContainer.length = spgrcontainerlen
val spgrcontainer = msofbtSpgrContainer.toByteArray()

 // DGCONTAINER
            val dgcontainerlen = (dgcontaineratoms.size + spgrcontainerlen + solverContainerLength + 8)        // drawing container length
 /*// debugging
    	  	if (dgcontainerlen!=origDG)
    	  		Logger.logErr("DGCONTAINERLENTH IS OFF: " + (dgcontainerlen-origDG));
    	  	/**/
            val msofbtDgContainer = MsofbtDgContainer(MSODrawingConstants.MSOFBTDGCONTAINER, 0, 15)
msofbtDgContainer.length = dgcontainerlen    //HEADERRECLENGTH + SPCONTAINERLENGTH + otherSPCONTAINERLENGTH);
val dgcontainer = msofbtDgContainer.toByteArray()

val header = ByteArray(HEADERRECLENGTH + spContainerLength - additionalSP + dgcontainer.size) //+retData.length];

var pos = 0
System.arraycopy(dgcontainer, 0, header, pos, dgcontainer.size)
pos += dgcontainer.size
System.arraycopy(dgcontaineratoms, 0, header, pos, dgcontaineratoms.size)
pos += dgcontaineratoms.size
System.arraycopy(spgrcontainer, 0, header, pos, spgrcontainer.size)
pos += spgrcontainer.size
System.arraycopy(container2, 0, header, pos, container2.size)
pos += container2.size
System.arraycopy(spcontainer2atoms, 0, header, pos, spcontainer2atoms.size)
pos += spcontainer2atoms.size

System.arraycopy(retData, 0, header, pos, retData.size)
retData = header
}
this.setData(retData)
this.length = data!!.size


 // testing
    	/*
    	io.starter.toolkit.Logger.log(this.toString());
    	io.starter.toolkit.Logger.log(this.debugOutput());
		/**/
    }

/**
 * add the set of subrecords necessary to define a Mso header record
 * <br></br>used when removing images, charts, etc. and have removed a previous header record
 */
     fun addHeader() {
 /*// testing
    	System.err.println(this.toString());
    	System.err.println(this.debugOutput());
    	/**/
        isHeader = true
if (lastSPID < SPID)
lastSPID = SPID        // TODO: Shouldn't assume to be SPID

val msofbtSp = MsofbtSp(MSODrawingConstants.MSOFBTSP, MSODrawingConstants.msosptMin, 2)  //1st shape rec is of type MSOSPTMIN
msofbtSp.setId(SPIDSEED)                      // SPMIN==SPIDSEED
msofbtSp.setGrfPersistence(5)                    // flag= fPatriarch,
val msofbtSpBytes = msofbtSp.toByteArray()
spContainerLength += msofbtSpBytes.size

val msofbtSpgr = MsofbtSpgr(MSODrawingConstants.MSOFBTSPGR, 0, 1)
msofbtSpgr.setRect(0, 0, 0, 0)
val msofbtSpgrBytes = msofbtSpgr.toByteArray()
spContainerLength += msofbtSpgrBytes.size

 // SPCONTAINER
        val msofbtSpContainer = MsofbtSpContainer(MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15)
msofbtSpContainer.length = msofbtSpgrBytes.size + msofbtSpBytes.size
val msofbtSpContainerBytes = msofbtSpContainer.toByteArray()
spContainerLength += 8    // account for the SPCONTAINER header length;

 // SPGRCONTAINER
        val msofbtSpgrContainer = MsofbtSpgrContainer(MSODrawingConstants.MSOFBTSPGRCONTAINER, 0, 15)
msofbtSpgrContainer.length = spContainerLength + otherSPCONTAINERLENGTH
val msofbtSpgrContainerBytes = msofbtSpgrContainer.toByteArray()

val msofbtDg = MsofbtDg(MSODrawingConstants.MSOFBTDG, drawingId, 0)
msofbtDg.setNumShapes(numShapes)                //Number of images and drawings.
msofbtDg.setLastSPID(lastSPID)                    //lastSPID
val msofbtDgBytes = msofbtDg.toByteArray()

 // DGCONTAINER
        val msofbtDgContainer = MsofbtDgContainer(MSODrawingConstants.MSOFBTDGCONTAINER, 0, 15)
msofbtDgContainer.length = HEADERRECLENGTH + spContainerLength + otherSPCONTAINERLENGTH
val msofbtDgContainerBytes = msofbtDgContainer.toByteArray()


val headerRec = ByteArray(this.getData()!!.size + 80)  // below records take 80 bytes

var pos = 0
System.arraycopy(msofbtDgContainerBytes, 0, headerRec, pos, msofbtDgContainerBytes.size)
pos += msofbtDgContainerBytes.size// 8
System.arraycopy(msofbtDgBytes, 0, headerRec, pos, msofbtDgBytes.size)
pos += msofbtDgBytes.size
System.arraycopy(msofbtSpgrContainerBytes, 0, headerRec, pos, msofbtSpgrContainerBytes.size)
pos += msofbtSpgrContainerBytes.size// 8
System.arraycopy(msofbtSpContainerBytes, 0, headerRec, pos, msofbtSpContainerBytes.size)
pos += msofbtSpContainerBytes.size// 8
System.arraycopy(msofbtSpgrBytes, 0, headerRec, pos, msofbtSpgrBytes.size)
pos += msofbtSpgrBytes.size
System.arraycopy(msofbtSpBytes, 0, headerRec, pos, msofbtSpBytes.size)
pos += msofbtSpBytes.size

System.arraycopy(this.getData()!!, 0, headerRec, pos, this.getData()!!.size)
setData(headerRec)

 /*// testing
    	System.err.println(this.toString());
    	System.err.println(this.debugOutput());
    	/**/
    }

/**
 * remove the set of subrecords necessary to define a MSO header record
 * <br></br>used when removing images, charts, etc. and have removed a previous header record
 */
     fun removeHeader() {
 /*// testing
    	System.err.println(this.toString());
    	System.err.println(this.debugOutput());
    	*/
        var spcontainer1atoms = ByteArray(0)
 // reset state vars
        spContainerLength = 0        // this shape container length

 // first pass, update significant atoms and sum up container lengths
        val fbt:Int
val len:Int
super.getData()
val bis = ByteArrayInputStream(data!!)
while (bis.available() > 0)
{
val header = ByteArray(8)
bis.read(header, 0, 8)
fbt = ((0xFF and header[3]) shl 8) or (0xFF and header[2])
len = ByteTools.readInt(header[4], header[5], header[6], header[7])
if ((0xF and header[0]) != 15)
{    // skip containers
 // parse atoms or sub-records (distinct from container records above)
                var data = ByteArray(len)
bis.read(data, 0, len)
when (fbt) {
MSODrawingConstants.MSOFBTDG    // update drawing record atom
 -> {}

MSODrawingConstants.MSOFBTSP            // A shape atom rec (inst= shape type) rec= shape ID + group of flags
 -> {
 // TODO: necessary to ever update the SPIDSEED?
                        val flag = ByteTools.readInt(data[4], data[5], data[6], data[7])
if (flag != 5)
{    // if it's not the SPIDseed, update SPID
System.arraycopy(ByteTools.cLongToLEBytes(SPID), 0, data, 0, 4)
}
data = ByteTools.append(data, header)
if (flag != 5)
 // if it's not the header
                            spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTCLIENTANCHOR        // Anchor or location for a shape
 -> {
 // udpate bounds
                        System.arraycopy(ByteTools.shortToLEBytes(clientAnchorFlag), 0, data, 0, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![0]), 0, data, 2, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![1]), 0, data, 4, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![2]), 0, data, 6, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![3]), 0, data, 8, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![4]), 0, data, 10, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![5]), 0, data, 12, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![6]), 0, data, 14, 2)
System.arraycopy(ByteTools.shortToLEBytes(bounds!![7]), 0, data, 16, 2)
data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTOPT    // property table atom - for 97 and earlier versions
 -> {
 // OPT= picture options
                        if (optRec == null)
 // if do not have a MsoFbtOPT record, then create new -- shouldn't get here as should always have an existing record
                            optRec = MsofbtOPT(MSODrawingConstants.MSOFBTOPT, 0, 3)    //version is always 3, inst is current count of properties.
if (imageIndex != -1 && imageIndex != optRec!!.getImageIndex())
optRec!!.setImageIndex(imageIndex)
if (name != null && name != "" && name != optRec!!.getImageName())
optRec!!.setImageName(name)
if (shapeName != null && shapeName != "" && shapeName != optRec!!.getShapeName())
{
optRec!!.setShapeName(shapeName)
}
data = optRec!!.toByteArray()
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTTERTIARYOPT    // for Office versions 2000, XP, 2003, and 2007
, MSODrawingConstants.MSOFBTSECONDARYOPT        // movie (id= 274)
 -> {
data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTCOLORSCHEME, MSODrawingConstants.MSOFBTREGROUPITEMS -> {}

MSODrawingConstants.MSOFBTCLIENTTEXTBOX // msofbtClientTextbox sub-record, contains only this one atom no containers ...
, MSODrawingConstants.MSOFBTTEXTBOX // msofbtClientTextbox sub-record, contains only this one atom no containers ...
 -> {
 // DON"T ADD LEN
                        data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTSPGR    // shapes that ARE groups, not shapes that are IN groups; The group shape record defines the coordinate system of the shape
 -> {}

MSODrawingConstants.MSOFBTCHILDANCHOR    //  used for all shapes that belong to a group. The content of the record is simply a RECT in the coordinate system of the parent group shape
, MSODrawingConstants.MSOFBTANCHOR        // used for top-level shapes when the shape streamed to the clipboard. The content of the record is simply a RECT with a coordinate system of 100,000 units per inch and origin in the top-left of the drawing
, MSODrawingConstants.MSOFBTCLIENTDATA, MSODrawingConstants.MSOFBTDELETEDPSPL -> {
data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}

MSODrawingConstants.MSOFBTCONNECTORRULE, MSODrawingConstants.MSOFBTALIGNRULE, MSODrawingConstants.MSOFBTARCRULE, MSODrawingConstants.MSOFBTCLIENTRULE, MSODrawingConstants.MSOFBTCALLOUTRULE -> {}

else -> {
Logger.logInfo("MSODrawing.removeHeader:  unknown subrecord encountered: " + fbt)
data = ByteTools.append(data, header)
spcontainer1atoms = ByteTools.append(data, spcontainer1atoms)
}
}// SOLVERCONTAINERATOMS+=len+8;
 //don't add to len
 // TODO: HANDLE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
}
}
spContainerLength = spcontainer1atoms.size
 // 2nd pass:  now have the important container lengths and their updated
        // subrecords/atoms, create resulting byte array

        // build main spcontainer - valid for all records
        val spcontainer1 = MsofbtSpContainer(MSODrawingConstants.MSOFBTSPCONTAINER, 0, 15)
spcontainer1.length = spContainerLength
val container = spcontainer1.toByteArray()
spContainerLength += container.size    // include this header length

 //if (!bIsHeader && SPCONTAINERLENGTH!=origSP)
        //Logger.logErr("SPCONTAINERLENTH IS OFF: " + (SPCONTAINERLENGTH-origSP));

        val retData = ByteArray(spContainerLength)
System.arraycopy(container, 0, retData, 0, container.size)
System.arraycopy(spcontainer1atoms, 0, retData, container.size, spcontainer1atoms.size)

setData(retData)
 /*// testing
    	System.err.println(this.toString());
    	System.err.println(this.debugOutput());
    	/**/
    }

/**
 * update the header records with new container lengths
 *
 * @param otherlength
 */
     fun updateHeader(otherSPContainers:Int, otherContainers:Int, numShapes:Int, lastSPID:Int) {
if (!isHeader)
{
Logger.logErr("Msodrawing.updateHeader is only applicable for the header drawing object")
return
}
 /*
		// dgcontainerlength= 	dg(+8) + regroupitems(+8) + spgrcontainer + solvercontainer(s) + colorscheme(+8)
		// spgrcontainerlength= 	sum of spcontainers i.e this spcontainerlength + otherspcontainers
	  	 */
        this.numShapes = numShapes
 // this.lastSPID= SPIDSEED + nImages; algorithm is wrong when book contains deleted images, etc.
        this.lastSPID = lastSPID
val fbt:Int
val len:Int
 // the two container lengths we are concerned about:
        //SPGRCONTAINERLENGTH
        otherSPCONTAINERLENGTH = otherSPContainers
var spgrcontainerlength = spContainerLength + otherSPCONTAINERLENGTH
 // DGCONTAINERLENGTH= spgrcontainerlength + all the atoms contained within the DG CONTAINER
        var dgcontainerlength = otherContainers + spgrcontainerlength
var DGATOMS = 0
var hasSPGRCONTAINER = false
 /* // debugging container lengths on update
		boolean diff= false;
		int origSPGRL= 0;
		int origDGL= 0;
		/**/
        // count dgcontainer atorms=MSOFBTDG, MSOFBTREGROUPITEMS, MSOFBTCOLORSCHEME (MSOFBTSPGRCONTAINER is added separately)
        super.getData()
var bis = ByteArrayInputStream(data!!)
while (bis.available() > 0)
{
var buf = ByteArray(8)
bis.read(buf, 0, 8)
fbt = ((0xFF and buf[3]) shl 8) or (0xFF and buf[2])
len = ByteTools.readInt(buf[4], buf[5], buf[6], buf[7])
 // 1st pass count dgcontainer atoms
            if (fbt == MSODrawingConstants.MSOFBTDGCONTAINER)
{
 // skip contianers
                //origDGL= len;	// debugging
            }
else if (fbt == MSODrawingConstants.MSOFBTSPCONTAINER)
{
 // skip containers
            }
else if (fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER)
{
if (hasSPGRCONTAINER)
{    // already has the 1 required SPGRCONTAINER, if more than 1, multiple groups
spgrcontainerlength += 8
}
DGATOMS += 8    // just count header length here
 //if (!hasSPGRCONTAINER)	origSPGRL= len;	// debugging
                hasSPGRCONTAINER = true
}
else if (fbt == MSODrawingConstants.MSOFBTOPT)
{
break    // nothing important after this
}
else
{
if (fbt == MSODrawingConstants.MSOFBTREGROUPITEMS || fbt == MSODrawingConstants.MSOFBTDG || fbt == MSODrawingConstants.MSOFBTCOLORSCHEME)
DGATOMS += (len + 8)
buf = ByteArray(len)
bis.read(buf, 0, len)
}
}
dgcontainerlength += DGATOMS
/** debugging  */
		/*if (origSPGRL!=spgrcontainerlength || origDGL!=dgcontainerlength) {
			io.starter.toolkit.Logger.log(this.toString());
			io.starter.toolkit.Logger.log(this.debugOutput());
			io.starter.toolkit.Logger.log("ORIGDG=" + origDGL + " ORIGSPL=" + origSPGRL + " DIFF: " + (origDGL-dgcontainerlength) + "-" + (origSPGRL-spgrcontainerlength));
			diff= true;
		}
		/**/

        super.getData()
bis = ByteArrayInputStream(data!!)
while (bis.available() > 0)
{
var buf = ByteArray(8)
bis.read(buf, 0, 8)
fbt = ((0xFF and buf[3]) shl 8) or (0xFF and buf[2])
len = ByteTools.readInt(buf[4], buf[5], buf[6], buf[7])

if (fbt == MSODrawingConstants.MSOFBTDGCONTAINER)
{
System.arraycopy(ByteTools.cLongToLEBytes(dgcontainerlength), 0, data!!, data!!.size - bis.available() - 4, 4)
}
else if (fbt == MSODrawingConstants.MSOFBTDG)
{  // Drawing Record: count + MSOSPID seed
buf = ByteArray(len)
bis.read(buf, 0, len)
val newrec = ByteArray(8)
System.arraycopy(ByteTools.cLongToLEBytes(this.numShapes), 0, newrec, 0, 4)
System.arraycopy(ByteTools.cLongToLEBytes(this.lastSPID), 0, newrec, 4, 4)
if (len == newrec.size)
{// should!!!
 // update Msodrawing data ...
                    System.arraycopy(newrec, 0, data!!, data!!.size - bis.available() - newrec.size, newrec.size)
}
else
Logger.logErr("UpdateClientAnchorRecord: New Array Size=" + newrec.size)
}
else if (fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER)
{    // sum of all spcontainers on the sheet
System.arraycopy(ByteTools.cLongToLEBytes(spgrcontainerlength), 0, data!!, data!!.size - bis.available() - 4, 4)
 /*if (diff) {	// debugging container lengths
					io.starter.toolkit.Logger.log(this.toString());
					io.starter.toolkit.Logger.log(this.debugOutput());
				}/**/
                return
}
else
{    // skip atoms
buf = ByteArray(len)
bis.read(buf, 0, len)
}
}
}

/**
 * update just the image index portion of this record
 * useful when you don't need to rebuild entire record (and thus possibly lose information)
 *
 * @param idx
 */
     fun updateImageIndex(idx:Int) {
val version:Int
val inst:Int
val fbt:Int
val len:Int
super.getData()
val bis = ByteArrayInputStream(data!!)
while (bis.available() > 0)
{
var buf = ByteArray(8)
bis.read(buf, 0, 8)
inst = ((0xFF and buf[1]) shl 4) or ((0xF0 and buf[0]) shr 4)
fbt = ((0xFF and buf[3]) shl 8) or (0xFF and buf[2])
len = ByteTools.readInt(buf[4], buf[5], buf[6], buf[7])
if (fbt < 0xF005)
{ // ignore containers
continue
}
buf = ByteArray(len)
bis.read(buf, 0, len)    // position at next sub-rec
if (fbt == MSODrawingConstants.MSOFBTOPT)
{
 // Find the property that governs image index and update the bytes
                val propertyId:Int
val n = inst                // number of properties to parse
var pos = 0                    // pointer to current property in data/property table
for (i in 0 until n)
{
propertyId = ((0x3F and buf[pos + 1]) shl 8) or (0xFF and buf[pos])
if (propertyId == MSODrawingConstants.msooptpib)
{// blip to display = image index
 // testing int dtx = ByteTools.readInt(dat[pos+2],dat[pos+3],dat[pos+4],dat[pos+5]);
                        val insertPosition = data!!.size - bis.available() - len + pos + 2
System.arraycopy(ByteTools.cLongToLEBytes(idx), 0, data!!, insertPosition, 4)
imageIndex = idx
return
}
pos += 6
}
}
}
}


/**
 * removes the header-specific portion of this Msodrawing record, if any
 * useful when adding Msodrawng recs from other sources such as adding charts
 * only one Msodrawing header record is allowed
 */
     fun makeNonHeader() {
if (!isHeader)
return     // nothing to do!
isHeader = false
this.removeHeader()
}

/**
 * set whether "this is the header drawing object for the sheet"
 *
 * @param b
 */
     fun setIsHeader() {
if (!isHeader)
{
this.addHeader()
}
}

/**
 * @return explicitly-set shape name (= name in NamedRange box, I believe)
 */
     fun getShapeName():String? {
return shapeName
}


/**
 * @return true if this image has a border
 */
    // TODO: detecting a border may be more complicated than this
    // TODO: ability to set a border
     fun hasBorder():Boolean {
if (optRec != null)
return optRec!!.hasBorder()
return false
}

/**
 * if has a border, return the border line width
 *
 * @return
 */
     fun borderLineWidth():Int {
if (optRec != null)
return optRec!!.borderLineWidth
return -1
}

 fun setImageIndex(value:Int) {
imageIndex = value
this.updateRecord()
}

 fun setImageName(name:String) {
if (name != this.name)
{
this.name = name
updateRecord()
}
}

/**
 * set the ordinal # for this drawing record
 *
 * @param id
 */
     fun setDrawingId(id:Int) {
drawingId = id
updateDGRecord()
}

/**
 * return the ordinal # for this record
 *
 * @return
 */
     fun getDrawingId():Int {
return drawingId
}

/**
 * allow setting of "Named Range" name
 *
 * @param name
 */
     fun setShapeName(name:String) {
if (name != shapeName)
{
shapeName = name
updateRecord()
}
}

 fun getImageIndex():Int {
return imageIndex
}

 fun setBounds(b:ShortArray) {
bounds = b.clone()
height = calcHeight()    // 20090831 KSC
updateClientAnchorRecord(bounds)
}


/**
 * set coordinates as {X, Y, W, H} in 7
 *
 * @param coords
 */
     fun setBoundsInPixels(coords:ShortArray) {
 // 20090505 KSC: convert
        val x = Math.round(coords[0].toFloat()).toShort()    // convert pixels to excel units
val y = Math.round(coords[1].toFloat()).toShort()    // convert pixels to points
val w = Math.round(coords[2].toFloat()).toShort()    // convert pixels to excel units
val h = Math.round(coords[3].toFloat()).toShort()    // convert pixels to points
 /*		setX(coords[0]);
		setY(coords[1]);
		setWidth(coords[2]);
		setHeight(coords[3]);
		*/
        setX(x.toInt())
setY(y.toInt())
setWidth(w.toInt())
setHeight(h.toInt())
}

/**
 * returns the bounds of this object
 * bounds are relative and based upon rows, columns and offsets within
 */
     fun getBounds():ShortArray? {
return bounds
}

 fun setRow(row:Int) {
bounds[MSODrawing.ROW] = row.toShort()
updateClientAnchorRecord(bounds)
}

/**
 * set the x position of this object
 * units are in excel units
 *
 * @param x
 */
     fun setX(x:Int) {
var z = 0
var col:Short = 0
var colOffset:Short = 0
var i:Short = 0
while (i < XLSConstants.MAXCOLS && z < x)
{
val w = getColWidth(i.toInt())
if (z + w < x)
z += w
else
{
col = i
colOffset = Math.round(1024 * (((x - z).toDouble()) / w.toDouble())).toShort()
z = x
}
i++
}
bounds[MSODrawing.COL] = col
bounds[MSODrawing.COLOFFSET] = colOffset
updateClientAnchorRecord(bounds)
}

/**
 * set the y position of this object
 * units are in points
 *
 * @param y
 */
     fun setY(y:Int) {
var z = 0.0
var row:Short = 0
var rowOffset:Short = 0
var i:Short = 0
while (z < y)
{
val h = getRowHeight(i.toInt())
if (z + h < y)
z += h
else
{
row = i
rowOffset = Math.round((256 * ((y - z) / getRowHeight(i.toInt())))).toShort()
z = y.toDouble()
}
i++
}
bounds[MSODrawing.ROW] = row
bounds[MSODrawing.ROWOFFSET] = rowOffset
updateClientAnchorRecord(bounds)
}

/**
 * calculate height based upon row #s, row offsets and row heights
 * units are in points
 *
 * @return short row height
 */
    private fun calcHeight():Short {
val row = bounds!![MSODrawing.ROW].toInt()
val row1 = bounds!![MSODrawing.ROW1].toInt()
val rowOff = bounds!![MSODrawing.ROWOFFSET] / 256.0
val rowOff1 = bounds!![MSODrawing.ROWOFFSET1] / 256.0
var y = getRowHeight(row) - getRowHeight(row) * rowOff
for (i in row + 1 until row1)
{
y += getRowHeight(i)
}
if (row1 > row)
y += getRowHeight(row1) * rowOff1
else
y = getRowHeight(row1) * (rowOff1 - rowOff)
return Math.round(y).toShort()
}

/**
 * set the width of this object
 * units are in excel units
 *
 * @param w
 */
     fun setWidth(w:Int) {
val col = bounds!![MSODrawing.COL].toInt()
val colOff = bounds!![MSODrawing.COLOFFSET] / 1024.0
var col1 = col
var colOff1 = 0
var z = getColWidth(col) - (getColWidth(col) * colOff).toInt()
if (z >= w)
{    // 20100322 KSC: was > w, so the == case never was handled, see TestImages.insertImageOffsetDisappearance bug
col1 = col
colOff1 = Math.round((1024 * (w / getColWidth(col).toDouble()))).toShort() + bounds!![MSODrawing.COLOFFSET]
}
var i = col + 1
while (i < XLSConstants.MAXCOLS && z < w)
{
val cw = getColWidth(i)
if (z + cw < w)
z += cw
else
{
col1 = i
colOff1 = (1024 * (((w - z).toDouble()) / cw.toDouble())).toShort().toInt()
z = w
}
i++
}
bounds[MSODrawing.COL1] = col1.toShort()
bounds[MSODrawing.COLOFFSET1] = colOff1.toShort()
updateClientAnchorRecord(bounds)
originalWidth = originalWidth    // 20071024 KSC: if change width, record
}

/**
 * set the height of this object
 * units are in points
 *
 * @param h
 */
     fun setHeight(h:Int) {
val row = bounds!![MSODrawing.ROW].toInt()
val rowOff = bounds!![MSODrawing.ROWOFFSET] / 256.0
var row1 = row
var rowOff1 = 0
var rh = getRowHeight(row)
var y = rh - (rh * rowOff)    // distance from start position to end of row
if (y > h)
{
rowOff1 = (256 * (h / rh)).toShort() + bounds!![MSODrawing.ROWOFFSET]
}
var i = row + 1
while (y < h)
{
rh = getRowHeight(i)
if (y + rh < h)
y += rh
else
{        // height is met; see what offset into row i is necessary
row1 = i
rowOff1 = Math.round((256 * ((h - y) / rh))).toShort().toInt()
y = h.toDouble()    // exit loop
}
i++
}
bounds[MSODrawing.ROW1] = row1.toShort()
bounds[MSODrawing.ROWOFFSET1] = rowOff1.toShort()
updateClientAnchorRecord(bounds)
height = calcHeight()    // 20071024 KSC: if change height, record
}

/**
 * return the column width in excel units
 *
 * @param col
 * @return
 */
    private fun getColWidth(col:Int):Int {
 // MSO column width is in Excel units= 0-255 characters
        // Excel Units are 1/256 of default font width (10 pt Arial)
        var w = Colinfo.DEFAULT_COLWIDTH.toDouble()
try
{
val co = this.sheet!!.getColInfo(col)
if (co != null)
w = co!!.getColWidth().toDouble()
}
catch (e:Exception) {    // exception if no col defined
}

return Math.round(w / 256).toInt()        // 20090505 KSC: try this
}

/**
 * return the row height in points
 *
 * @param row
 * @return
 */
    private fun getRowHeight(row:Int):Double {
 // MSO row height is measured in points, 0-409
        // in Arial 10 pt, standard row height is 12.75 points
        // .75 points/pixel
        // row height is in twips= 1/20 of a point,1 pt= 1/72 inch
        var h = 255
try
{
val r = this.sheet!!.getRowByNumber(row)
if (r != null)
h = r!!.rowHeight
else
 // no row defined - use default row height 20100504 KSC
                return this.sheet!!.defaultRowHeight // default
}
catch (e:Exception) {    // exception if no row defined // no row defined - use default row height 20100504 KSC
return this.sheet!!.defaultRowHeight // default
}

return (h / 20.0)        // 20090506 KSC: it's in twips 1/20 of a point
}

/**
 * @return Returns the numShapes.
 */
     fun getNumShapes():Int {
return numShapes
}

/**
 * set the number of shapes for this drawing rec
 *
 * @param n
 */
     fun setNumShapes(n:Int) {
numShapes = n
updateDGRecord()
}

/**
 * return the lastSPID (only valid for the header msodrawing record
 * necessary to track so that newly added images have the appropriate SPID
 *
 * @return
 */
     fun getlastSPID():Int {
return this.lastSPID
}

/**
 * update the bytes of the DG record
 */
    private fun updateDGRecord() {
val fbt:Int
val len:Int

super.getData()
val bis = ByteArrayInputStream(data!!)
while (bis.available() > 0)
{
var buf = ByteArray(8)
bis.read(buf, 0, 8)
fbt = ((0xFF and buf[3]) shl 8) or (0xFF and buf[2])
len = ByteTools.readInt(buf[4], buf[5], buf[6], buf[7])
if (fbt < 0xF005)
{ // ignore containers
continue
}
buf = ByteArray(len)
bis.read(buf, 0, len)    // position at next sub-rec
if (fbt == MSODrawingConstants.MSOFBTDG)
{  // Drawing Record: count + MSOSPID seed
data[8] = (drawingId * 16).toByte()    //= the 1st byte of the header portion of the DG record 20080902 KSC
val newrec = ByteArray(8)
System.arraycopy(ByteTools.cLongToLEBytes(numShapes), 0, newrec, 0, 4)
System.arraycopy(ByteTools.cLongToLEBytes(this.lastSPID), 0, newrec, 4, 4)
if (len == newrec.size)
{// should!!!
 // update Msodrawing data ...
                    System.arraycopy(newrec, 0, data!!, data!!.size - bis.available() - newrec.size, newrec.size)
}
else
Logger.logErr("UpdateClientAnchorRecord: New Array Size=" + newrec.size)
return
}
}
}

/**
 * update the bytes of the CLIENTANCHOR record
 *
 * @param bounds bounds[0]= column # of top left position (0-based) of the shape
 * bounds[1]= x offset within the top-left column
 * bounds[2]= row # for top left corner
 * bounds[3]= y offset within the top-left corner
 * bounds[4]= column # of the bottom right corner of the shape
 * bounds[5]= x offset within the cell  for the bottom-right corner
 * bounds[6]= row # for bottom-right corner of the shape
 * bounds[7]= y offset within the cell for the bottom-right corner
 */
    private fun updateClientAnchorRecord(bounds:ShortArray?) {
val fbt:Int
val len:Int

super.getData()
val bis = ByteArrayInputStream(data!!)
while (bis.available() > 0)
{
var buf = ByteArray(8)
bis.read(buf, 0, 8)
fbt = ((0xFF and buf[3]) shl 8) or (0xFF and buf[2])
len = ByteTools.readInt(buf[4], buf[5], buf[6], buf[7])
if (fbt < 0xF005)
{ // ignore containers
continue
}
buf = ByteArray(len)
bis.read(buf, 0, len)    // position at next sub-rec
if (fbt == MSODrawingConstants.MSOFBTCLIENTANCHOR)
{        // Anchor or location fo a shape
 // udpate bounds
                val bos = ByteArrayOutputStream()
try
{
bos.write(ByteTools.shortToLEBytes(clientAnchorFlag))
bos.write(ByteTools.shortToLEBytes(bounds!![0]))
bos.write(ByteTools.shortToLEBytes(bounds!![1]))
bos.write(ByteTools.shortToLEBytes(bounds!![2]))
bos.write(ByteTools.shortToLEBytes(bounds!![3]))
bos.write(ByteTools.shortToLEBytes(bounds!![4]))
bos.write(ByteTools.shortToLEBytes(bounds!![5]))
bos.write(ByteTools.shortToLEBytes(bounds!![6]))
bos.write(ByteTools.shortToLEBytes(bounds!![7]))
}
catch (e:Exception) {}

val newrec = bos.toByteArray()
if (buf.size == newrec.size)
{// should!!!
 // update Msodrawing data ...
                    System.arraycopy(newrec, 0, data!!, data!!.size - bis.available() - newrec.size, newrec.size)
}
else
Logger.logErr("UpdateClientAnchorRecord: New Array Size=" + newrec.size)
return
}
}
}

/**
 * sets a specific OPT subrecord
 *
 * @param propertyId   int property id see Msoconstants
 * @param isBid        true if this is a BLIP id
 * @param isComplex    true if has complexBytes
 * @param dtx          if not isComplex, the value; if isComplex, length
 * @param complexBytes complex bytes if isComplex
 */
     fun setOPTSubRecord(propertyId:Int, isBid:Boolean, isComplex:Boolean, dtx:Int, complexBytes:ByteArray) {
val optrec = this.optRec
 // TODO: store optrec length instead of calculating each time
        val origlen = optrec!!.toByteArray().size
optrec!!.setProperty(MSODrawingConstants.msooptGroupShapeProperties, isBid, isComplex, dtx, complexBytes)
updateRecord()
if (origlen != optrec!!.toByteArray().size)
{    // must update header
this.workBook!!.updateMsodrawingHeaderRec(this.sheet)
}
}

/**
 * set the LastSPID for this drawing record, if it's a header-type record
 *
 * @param spid
 */
     fun setLastSPID(spid:Int) {
lastSPID = spid
updateDGRecord()
}


 fun setShapeType(shapeType:Int) {
this.shapeType = shapeType.toShort()
}

 fun getShapeType():Int {
return shapeType.toInt()
}

/**
 * change the SPID for this record
 *
 * @param spid
 */
    private fun updateSPID() {
val fbt:Int
val len:Int
val inst:Int
super.getData()
val bis = ByteArrayInputStream(data!!)
while (bis.available() > 0)
{
var buf = ByteArray(8)
bis.read(buf, 0, 8)
inst = ((0xFF and buf[1]) shl 4) or ((0xF0 and buf[0]) shr 4)
fbt = ((0xFF and buf[3]) shl 8) or (0xFF and buf[2])
len = ByteTools.readInt(buf[4], buf[5], buf[6], buf[7])
if (fbt >= 0xF005)
{ // ignore containers
buf = ByteArray(len)
bis.read(buf, 0, len)

if (fbt == MSODrawingConstants.MSOFBTSP)
{
 //byte[] dat = new byte[len];
                    //bis.read(dat,0,len);
                    val flag = ByteTools.readInt(buf[4], buf[5], buf[6], buf[7])
 //					if (flag!=5) {	// if it's not the SPIDseed
                    System.arraycopy(ByteTools.cLongToLEBytes(SPID), 0, data!!, data!!.size - bis.available() - len, 4)
 //SPID = ByteTools.readInt(dat[0],dat[1],dat[2],dat[3]);
 //						return;
 //					}
                }
}
}
}


/**
 * update the header records with new container lengths
 *
 * @param otherlength
 */
    public override fun toString():String {
val sb = StringBuffer()
sb.append("Msodrawing: image=" + this.name + ".\t" + this.shapeName + "\timageIndex=" + imageIndex + " sheet=" + (if ((this.sheet != null)) this.sheet!!.sheetName else "none"))
sb.append(" ID= " + drawingId + " SPID=" + SPID)
sb.append("\tShapeType= " + shapeType)
if (isHeader)
sb.append("\tNumber of Shapes=" + numShapes + " Last SPID=" + lastSPID + " SPCL=" + spContainerLength + " otherLen=" + otherSPCONTAINERLENGTH)
else
sb.append(" SPCL=" + spContainerLength)
if (!isShape)
sb.append(" NOT A SHAPE")
return sb.toString()
}

/**
 * 20081106 KSC: when set sheet, record original height and width
 * as dependent upon row heights ...
 */
    public override fun setSheet(bs:Sheet?) {
super.setSheet(bs)
 // 20081106 Moved from parse as sheet must be set before calcHeight call
        // Record original height and width in case of row height/col width changes
        height = calcHeight()
originalWidth = originalWidth
}

/**
 * create records sub-records necessary to define an AutoFilter drop-down symbol
 */
     fun createDropDownListStyle(col:Int) {
 // mso record which - we hope - has the specific options necessary to define the dropdown box
        val optrec = this.optRec
optrec!!.inst = 0 // clear out
optrec!!.setData(byteArrayOf())
optrec!!.setProperty(MSODrawingConstants.msooptfLockAgainstGrouping, false, false, 17039620, null)
optrec!!.setProperty(MSODrawingConstants.msooptfFitTextToShape, false, false, 524296, null)
optrec!!.setProperty(MSODrawingConstants.msooptfNoLineDrawDash, false, false, 524288, null)
optrec!!.setProperty(MSODrawingConstants.msooptGroupShapeProperties, false, false, 131072, null)
this.setShapeType(MSODrawingConstants.msosptHostControl)    // shape type for these drop-downs
this.updateRecord(++this.wkbook!!.lastSPID, null, null, -1, shortArrayOf(col.toShort(), 0, 0, 0, (col + 1).toShort(), 0, 1, 0))  // generate msoDrawing using correct values moved from above
 // KSC: keep for testing
        //col++;
        //this.updateRecord(++this.wkbook.lastSPID, null, null, -1, new short[] {(short)col, 0, 0, 0, (short)(col+1), 272, 1, 0});  // generate msoDrawing using correct values moved from above
    }

/**
 * create the sub-records necessary to define a Comment (Note) Box
 */
     fun createCommentBox(row:Int, col:Int) {
val optrec = this.optRec
optrec!!.inst = 0 // clear out
optrec!!.setData(byteArrayOf())
 //47752196
        // try a random text id ... instead of 48678864
        val id = java.util.Random().nextInt()
optrec!!.setProperty(MSODrawingConstants.msofbtlTxid, false, false, id, null)
optrec!!.setProperty(MSODrawingConstants.msofbttxdir, false, false, 2, null)
optrec!!.setProperty(MSODrawingConstants.msooptfFitTextToShape, false, false, 524296, null)
optrec!!.setProperty(344, false, false, 0, null)    // no info on this subrecord
optrec!!.setProperty(MSODrawingConstants.msooptfillColor, false, false, 134217808, null)
optrec!!.setProperty(MSODrawingConstants.msooptfillBackColor, false, false, 134217808, null)
optrec!!.setProperty(MSODrawingConstants.msooptfNoFillHitTest, false, false, 1048592, null)
optrec!!.setProperty(MSODrawingConstants.msooptfshadowColor, false, false, 0, null)
optrec!!.setProperty(MSODrawingConstants.msooptfShadowObscured, false, false, 196611, null)
 // this, strangely, controls note hidden or show- when it's 131074, it's hidden, when it's 131072, it's shown
        optrec!!.setProperty(MSODrawingConstants.msooptGroupShapeProperties, false, false, 131074, null)
this.setShapeType(MSODrawingConstants.msosptTextBox)    // shape type for text boxes
 // position of text box - garnered from Excel examples
        // [1, 240, 0, 30, 3, 496, 4, 196]	A1
        // [4, 240, 2, 105, 6, 496, 7, 15]  D4
        this.updateRecord(++this.wkbook!!.lastSPID, null, null, -1, shortArrayOf((col + 1).toShort(), 240, row.toShort(), 30, (col + 3).toShort(), 496, (row + 4).toShort(), 196))  // generate msoDrawing using correct values moved from above
}
 /* notes from another attempt:  does this match ours?
_store_mso_opt_comment {
    my $self        = shift;
    my $type        = 0xF00B;
    my $version     = 3;
    my $instance    = 9;
    my $data        = '';
    my $length      = 54;
    my $spid        = $_[0];
    my $visible     = $_[1];
    my $colour      = $_[2] || 0x50;
    $data    = pack "V",  $spid;
    $data   .= pack "H*", '0000BF00080008005801000000008101' ;
    $data   .= pack "C",  $colour;
    $data   .= pack "H*", '000008830150000008BF011000110001' .
                          '02000000003F0203000300BF03';
    $data   .= pack "v",  $visible;
    $data   .= pack "H*", '0A00';
     *
     */


    /**
 * test debug output - FOR INTERNAL USE ONLY
 *
 * @return
 */
     fun debugOutput():String {
val bis = ByteArrayInputStream(super.getData()!!)
val version:Int
val inst:Int
val fbt:Int
val len:Int
val log = StringBuffer()
try
{
while (bis.available() > 0)
{
var dat = ByteArray(8)
bis.read(dat, 0, 8)
version = (0xF and dat[0])
inst = ((0xFF and dat[1]) shl 4) or ((0xF0 and dat[0]) shr 4)
fbt = ((0xFF and dat[3]) shl 8) or (0xFF and dat[2])
len = ByteTools.readInt(dat[4], dat[5], dat[6], dat[7])

log.append(fbt + " " + version + "/" + inst + "/" + len)

 //		    if (fbt <= 0xF005) { // ignore containers
                if (version == 15)
{ // it's a container - no parsing
 // MSOFBTSPGRCONTAINER:		// Shape Group Container, contains a variable number of shapes (=msofbtSpContainer) + other groups 0xF003
                    // MSOFBTSPCONTAINER:		// Shape Container 0xF004
                    if (fbt == MSODrawingConstants.MSOFBTDGCONTAINER)
{
log.append("\tMSOFBTDGCONTAINER")
}
else if (fbt == MSODrawingConstants.MSOFBTSPGRCONTAINER)
{
log.append("\tMSOFBTSPGRCONTAINER")
}
else if (fbt == MSODrawingConstants.MSOFBTSPCONTAINER)
{
log.append("\tMSOFBTSPCONTAINER")
}
else if (fbt == MSODrawingConstants.MSOFBTSOLVERCONTAINER)
{
log.append("\tMSOFBTSOLVERCONTAINER")
}
else
{
log.append("\tUNKNOWN CONTAINER")
}
log.append("\r\n")
continue
}

dat = ByteArray(len)
bis.read(dat, 0, len)
when (fbt) {
MSODrawingConstants.MSOFBTCALLOUTRULE -> log.append("\tMSOFBTCALLOUTRULE")

MSODrawingConstants.MSOFBTDELETEDPSPL -> log.append("\tMSOFBTDELETEDPSPL")

MSODrawingConstants.MSOFBTREGROUPITEMS -> log.append("\tMSOFBTREGROUPITEMS")

MSODrawingConstants.MSOFBTSP            // A shape atom rec (inst= shape type) rec= shape ID + group of flags
 -> {
log.append("\tMSOFBTSP")
val flag:Int
flag = ByteTools.readInt(dat[4], dat[5], dat[6], dat[7])
if ((flag and 0x800) == 0x800)
 // it has a shape-type property
                            log.append("\tshapeType=" + inst)
if (inst == 0)
 //== shape type
                            log.append("\tSPIDSEED=" + ByteTools.readInt(dat[0], dat[1], dat[2], dat[3]))
else
log.append("\tSPID=" + +ByteTools.readInt(dat[0], dat[1], dat[2], dat[3]))
log.append("\tflag=" + flag)
}

MSODrawingConstants.MSOFBTCLIENTANCHOR        // Anchor or location fo a shape
 -> {
log.append("\tMSOFBTCLIENTANCHOR")
log.append("\t[")
log.append(ByteTools.readShort(dat[2].toInt(), dat[3].toInt()) + ",")
log.append(ByteTools.readShort(dat[4].toInt(), dat[5].toInt()) + ",")
log.append(ByteTools.readShort(dat[6].toInt(), dat[7].toInt()) + ",")
log.append(ByteTools.readShort(dat[8].toInt(), dat[9].toInt()) + ",")
log.append(ByteTools.readShort(dat[10].toInt(), dat[11].toInt()) + ",")
log.append(ByteTools.readShort(dat[12].toInt(), dat[13].toInt()) + ",")
log.append(ByteTools.readShort(dat[14].toInt(), dat[15].toInt()) + ",")
log.append(ByteTools.readShort(dat[16].toInt(), dat[17].toInt()).toInt())
log.append("]")
}

MSODrawingConstants.MSOFBTOPT    // property table atom
 -> {
log.append("\tMSOFBTOPT")
 //MsofbtOPT optrec= new MsofbtOPT(fbt, inst, version);
                        //optrec.setData(dat);	// sets and parses msoFbtOpt data, including imagename, shapename and imageindex
                        log.append(optRec!!.debugOutput())
}

MSODrawingConstants.MSOFBTSECONDARYOPT    // property table atom  block 2
 -> {
log.append("\tMSOFBTSECONDARYOPT")
val secondaryoptrec = MsofbtOPT(fbt, inst, version)
secondaryoptrec.setData(dat)    // sets and parses msoFbtOpt data, including imagename, shapename and imageindex
log.append(secondaryoptrec.debugOutput())
}

MSODrawingConstants.MSOFBTTERTIARYOPT    // property table atom  block 3
 -> {
log.append("\tMSOFBTTERTIARYOPT")
val tertiaryoptrec = MsofbtOPT(fbt, inst, version)
tertiaryoptrec.setData(dat)    // sets and parses msoFbtOpt data, including imagename, shapename and imageindex
log.append(tertiaryoptrec.debugOutput())
}

MSODrawingConstants.MSOFBTDG    // Drawing Record: ID, num shapes + Last SPID for this DG
 -> {
log.append("\tMSOFBTDG")
log.append("\tID=" + inst)
log.append("\tns=" + ByteTools.readInt(dat[0], dat[1], dat[2], dat[3]))  // number of shapes in this drawing
log.append("\tllastSPID=" + ByteTools.readInt(dat[4], dat[5], dat[6], dat[7]))
}

MSODrawingConstants.MSOFBTCLIENTTEXTBOX -> log.append("\tMSOFBTCLIENTTEXTBOX")

MSODrawingConstants.MSOFBTSPGR -> log.append("\tMSOFBTSPGR")

MSODrawingConstants.MSOFBTCLIENTDATA -> log.append("\tMSOFBTCLIENTDATA")

MSODrawingConstants.MSOFBTSOLVERCONTAINER -> log.append("\tMSOFBTSOLVERCONTAINER")

MSODrawingConstants.MSOFBTCHILDANCHOR -> log.append("\tMSOFBTCHILDANCHOR")

MSODrawingConstants.MSOFBTCONNECTORRULE -> log.append("\tMSOFBTCONNECTORRULE")

else    //MSOFBTCONNECTORRULE
 -> log.append("\tUNKNOWN ATOM")
}
log.append("\r\n")
}
}
catch (e:Exception) {
log.append("\r\nEXCEPTION: " + e.toString())
}

log.append("**")
return log.toString()
}

public override fun close() {
super.close()
this.bounds = null
optRec = null    // 20091209 KSC: save MsofbtOPT records for later updating ...
secondaryoptrec = null
tertiaryoptrec = null    // apparently can have secondary and tertiary obt recs depending upon version in which it was saved ...
}

companion object {
private val serialVersionUID = 8275831369787287975L
internal val HEADERRECLENGTH = 24            //

 // correct way of setting/getting shape (image and chart) bounds
    /* position methods */
     val COL = 0
 val COLOFFSET = 1
 val ROW = 2
 val ROWOFFSET = 3
 val COL1 = 4
 val COLOFFSET1 = 5
 val ROW1 = 6
 val ROWOFFSET1 = 7
 val OFFSETMAX = 1023

 // conversions:  approx .136 excel units/pixel, 7.5 pixels/excel unit for X/W/Columns (really 6.4????)
    //               approx .75 pixels/points (Y,H/Rows) (really .60???)
    /**
 * Untested Info from Excel:
 * width/height in pixels = (w/h field - 8) * DPI of the display device / 72
 * DPI for Windows is 96 (Mac= 72)
 * thus conversion factor= 1.3333 for Windows devices
 * NOTES: seems correct for y/h, not for x/w
 * return coordinates as {X, Y, W, H} in pixels
 *
 * @return
 */
    // these are approximate conversion factors to convert excel units to pixels ...
     val XCONVERSION = 10.0        // convert excel units to pixels, garnered from actual comparisions rather than any equation (since it's based upon the normal font, it cannot be 100% for all calculations ...
 val WCONVERSION = 8.8    // ""
 val PIXELCONVERSION = 1.333    // see above


 // prototype bytes contain an image index and an image name - remove here
 // TODO: make prototype bytes correct
 //        mso.imageIndex= -1;
 //    	mso.imageName= "";
 //        mso.optrec.setImageName("");
 //        mso.optrec.setImageIndex(-1);
 //        mso.updateRecord();
 val prototype:XLSRecord?
get() {
val mso = MSODrawing()
mso.opcode = XLSConstants.MSODRAWING
mso.setData(mso.PROTOTYPE_BYTES)
mso.init()
return mso
}

/**
 * creates a msofbtClientTextbox, a very strange mso record containing only one sub-record or atom;
 * associated data=the text in the textbox, in a host-defined format i.e. text from a NOTE
 * <br></br>
 * NOTE: this must be flagged as incomplete so that it is not counted in number of shapes and other calcs ...
 * <br></br>
 * NOTE: knowledge of this record is very sketchy ...
 *
 * @return
 */
    // only contains 1 sub-record:  0xF00D or 61453= msofbtTextBox==shape has attached text
 val textBoxPrototype:XLSRecord
get() {
val mso = MSODrawing()
mso.opcode = XLSConstants.MSODRAWING
mso.setData(byteArrayOf(0, 0, 13, -16, 0, 0, 0, 0))
mso.init()
return mso
}
}

}