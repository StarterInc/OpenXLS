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
package io.starter.OpenXLS

import io.starter.formats.OOXML.SpPr
import io.starter.formats.OOXML.TwoCellAnchor
import io.starter.formats.XLS.Boundsheet
import io.starter.formats.XLS.MSODrawing
import io.starter.formats.XLS.MSODrawingConstants
import io.starter.toolkit.Logger
import org.json.JSONException
import org.json.JSONObject

import javax.imageio.ImageIO
import javax.imageio.ImageReader
import javax.imageio.stream.ImageInputStream
import java.awt.image.BufferedImage
import java.io.*

//OOXML-specific structures


/**
 * The ImageHandle provides access to an Image embedded in a spreadsheet.<br></br>
 * <br></br>
 * Use the ImageHandle to work with images in spreadsheet.<br></br>
 * <br></br>  <br></br>
 * With an ImageHandle you can:
 * <br></br><br></br>
 * <blockquote>
 * insert images into your spreadsheet
 * set the position of the image
 * set the width and height of the image
 * write spreadsheet image files to any outputstream
 *
</blockquote> *
 * <br></br>
 * <br></br>
 *
 * @see io.starter.OpenXLS.WorkBookHandle
 *
 * @see io.starter.OpenXLS.WorkSheetHandle
 */
class ImageHandle : Serializable {
    /**
     * returns width divided by height for the aspect ratio
     *
     * @return the aspect ratio for the image
     */
    //(convert to in)
    //(convert to in)
    val aspectRatio: Double
        get() {
            val height: Double
            val width: Double
            var aspectRatio = 0.0
            height = getHeight() / 122.27
            width = getWidth() / 57.06
            aspectRatio = width / height
            return aspectRatio
        }
    private var imageBytes: ByteArray? = null
    /**
     * returns the WorkSheet this image is contained in
     *
     * @return Returns the sheet.
     */
    var sheet: Boundsheet? = null
        private set

    /**
     * return the imageName (= original file name, I believe)
     *
     * @return
     */
    var imageName = " "
        private set
    /**
     * returns the explicitly set shape name (set by entering text in the named range field in Excel)
     *
     * @return the shape name
     */
    /**
     * allow setting of image name as seen in the Named Range Box
     *
     * @param name
     */
    // only set name/update record if names have changed
    // update is done in setShapeName
    // 20100202 KSC: must update header if change mso's - Claritas image insert regression bug (testImages.testInsertImageCorruption)
    var shapeName: String? = ""
        set(name) {
            if (name != null)
                field = name
            else
                field = ""
            if (this.msodrawing != null && this.shapeName != this.msodrawing!!.shapeName) {
                this.msodrawing!!.shapeName = this.shapeName!!
                this.msodrawing!!.workBook!!.updateMsodrawingHeaderRec(this.sheet)
            }
        }
    private var height: Short = 0
    private var width: Short = 0
    private var x: Short = 0
    private var y: Short = 0
    /**
     * @return Returns the image type.
     */
    /**
     * @param image_type The image type to set.
     */
    var imageType = -1
    private val DEBUGLEVEL = 0 // 	eventually will set!

    // OOXML-specific
    /**
     * return the OOXML shape property for this image
     *
     * @return
     */
    /**
     * define the OOXML shape property for this image from an existing spPr element
     */
    //    	imagesp.setNS("xdr");
    var spPr: SpPr? = null
    private var editMovement: String? = null

    var msodrawing: MSODrawing? = null
        private set    //20070924 KSC: link to actual msodrawing rec that describes this image


    /**
     * Returns the image type
     *
     * @return
     */
    val type: String
        get() {
            when (imageType) {
                io.starter.formats.XLS.MSODrawingConstants.IMAGE_TYPE_GIF -> return "gif"
                io.starter.formats.XLS.MSODrawingConstants.IMAGE_TYPE_JPG -> return "jpeg"
                io.starter.formats.XLS.MSODrawingConstants.IMAGE_TYPE_PNG -> return "png"
                io.starter.formats.XLS.MSODrawingConstants.IMAGE_TYPE_EMF -> return "emf"
                else -> return "undefined"
            }
        }

    /**
     * Returns the image mime type
     *
     * @return
     */
    val mimeType: String
        get() = "image/$imageType"

    /**
     * returns true if this drawing is active i.e. not deleted
     * <br></br>Note this is experimental
     *
     * @return
     */
    // default
    val isActive: Boolean
        get() = if (this.msodrawing != null) this.msodrawing!!.isActive else true


    /*
     * 20070924 KSC: image bounds are set via column/row # + offsets within the respective cell
     * added many position methods that access msodrawing for actual positions
     */
    /* position methods */

    /**
     * return the image bounds
     * images bounds are as follows:
     * bounds[0]= column # of top left position (0-based) of the shape
     * bounds[1]= x offset within the top-left column	(0-1023)
     * bounds[2]= row # for top left corner
     * bounds[3]= y offset within the top-left corner (0-1023)
     * bounds[4]= column # of the bottom right corner of the shape
     * bounds[5]= x offset within the cell  for the bottom-right corner (0-1023)
     * bounds[6]= row # for bottom-right corner of the shape
     * bounds[7]= y offset within the cell for the bottom-right corner (0-1023)
     */
    // short is not big enough to handle MAXROWS... correct? -jm

    /**
     * sets the image bounds
     * images bounds are as follows:
     * bounds[0]= column # of top left position (0-based) of the shape
     * bounds[1]= x offset within the top-left column (0-1023)
     * bounds[2]= row # for top left corner
     * bounds[3]= y offset within the top-left corner	(0-1023)
     * bounds[4]= column # of the bottom right corner of the shape
     * bounds[5]= x offset within the cell  for the bottom-right corner (0-1023)
     * bounds[6]= row # for bottom-right corner of the shape
     * bounds[7]= y offset within the cell for the bottom-right corner (0-1023)
     */
    var bounds: ShortArray?
        get() = msodrawing!!.bounds
        set(bounds) {
            msodrawing!!.bounds = bounds!!
        }

    /**
     * return the image bounds in x, y, width, height format in pixels
     *
     * @return
     */
    val coords: ShortArray
        get() = if (msodrawing != null)
            msodrawing!!.coords
        else
            shortArrayOf(x, y, width, height)

    /**
     * return the topmost row of the image
     *
     * @return
     */
    /**
     * set the topmost row of the image
     *
     * @param row
     */
    var row: Int
        get() = msodrawing!!.row0
        set(row) = msodrawing!!.setRow(row)

    /**
     * return the lower row of the image
     *
     * @return
     */
    /**
     * set the lower row of the image
     *
     * @param row
     */
    var row1: Int
        get() = msodrawing!!.row1
        set(row) {
            msodrawing!!.row1 = row
        }

    /**
     * return the leftmost column of the image
     */
    val col: Int
        get() = msodrawing!!.col

    /**
     * return the rightmost column of the image
     */
    val col1: Int
        get() = msodrawing!!.col1

    val originalWidth: Int
        get() = msodrawing!!.originalWidth.toInt()

    /**
     * get the position of the top left corner of the image in the sheet
     */

    val rowAndOffset: ShortArray
        get() = msodrawing!!.rowAndOffset

    /**
     * get the position of the top left corner of the image in the sheet
     */
    val colAndOffset: ShortArray
        get() = msodrawing!!.colAndOffset

    /**
     * Internal method that converts the image bounds appropriate for saving
     * by Excel in the MsofbtOPT record.
     */
    val imageIndex: Int
        get() = msodrawing!!.imageIndex

    /**
     * @return Returns the name (either explicitly set name or image name).
     */
    /**
     * @param name The name to set.
     */
    // 20071025 KSC: use Explicitly set name = shape name, if present.  Otherwise, use imageName
    // NOTE: updating mso code causes corruption in certain Infoteria files- must look at
    // update is done in setImageName
    // must update header if change mso's - Claritas image insert regression bug (testImages.testInsertImageCorruption)
    var name: String
        get() = if (this.shapeName != "") this.shapeName else this.imageName
        set(name) {
            this.imageName = name
            if (this.msodrawing != null && this.msodrawing!!.name != name) {
                this.msodrawing!!.setImageName(name)
                this.msodrawing!!.workBook!!.updateMsodrawingHeaderRec(this.sheet)
            }
        }

    /**
     * Get a JSON representation of the format
     *
     * @param cr
     * @return
     */
    // short[] coords =  { x, y, width, height };
    // ch.put("width", width); // for some reason COORDS wrong for width
    val json: JSONObject
        get() {
            val ch = JSONObject()
            try {
                ch.put("name", this.imageName)
                val coords = this.coords

                ch.put("x", coords[ImageHandle.X].toInt())
                ch.put("y", coords[ImageHandle.Y].toInt())

                ch.put("width", coords[ImageHandle.WIDTH].toInt())

                ch.put("height", coords[ImageHandle.HEIGHT].toInt())
                ch.put("type", this.type)

            } catch (e: JSONException) {
                Logger.logErr("Error getting imageHandle JSON: $e")
            }

            return ch
        }

    fun setMsgdrawing(rec: MSODrawing) {
        msodrawing = rec
    }

    /**
     * Constructor which takes image file bytes and inserts into
     * specific sheet
     *
     * @param imageBytes
     * @param sheet
     */
    constructor(imagebytestream: InputStream, _sheet: WorkSheetHandle) : this(imagebytestream, _sheet.mysheet) {}

    /**
     * Constructor  which takes image file bytes and associates it with the
     * specified boundsheet
     *
     * @param imagebytestream
     * @param bs
     */
    constructor(imagebytestream: InputStream, bs: Boundsheet?) {
        sheet = bs
        try {
            imageBytes = ByteArray(imagebytestream.available())
            imagebytestream.read(imageBytes!!)
        } catch (ex: Exception) {
            System.err.print("Failed to create new ImageHandle in sheet" + bs!!.sheetName + " from InputStream:" + ex.toString())
        }

        initialize()
    }

    /**
     * Constructor which takes image file bytes and inserts into
     * specific sheet
     *
     * @param imageBytes
     * @param sheet
     */
    constructor(_imageBytes: ByteArray, sheet: Boundsheet) {
        this.sheet = sheet
        imageBytes = _imageBytes

        initialize()

    }

    /**
     * Override equals so that Sheet cannot contain dupes.
     *
     * @see java.lang.Object.equals
     */
    override fun equals(another: Any?): Boolean {
        return another!!.toString() == this.toString()
    }

    private fun initialize() {
        // after instantiating the image you should have
        // access to the width and height bounds
        var imageFormat: String? = ""
        try {
            val bis = ByteArrayInputStream(imageBytes!!)

            imageFormat = getImageFormat(bis)
            if (imageFormat == null) {
                // 20080128 KSC: Occurs when image format= EMF ????? We cannot interpret, will most likely crash file
                if (DEBUGLEVEL > WorkBook.DEBUG_LOW)
                    Logger.logErr("ImageHandle.initialize: Unrecognized Image Format")
                return
            }
            if (!(imageFormat.equals("jpeg", ignoreCase = true) || imageFormat.equals("png", ignoreCase = true))) {
                bis.reset()
                imageBytes = convertData(bis)
                imageType = MSODrawingConstants.IMAGE_TYPE_PNG
            } else {
                if (imageFormat.equals("jpeg", ignoreCase = true))
                    imageType = MSODrawingConstants.IMAGE_TYPE_JPG
                else
                    this.imageType = MSODrawingConstants.IMAGE_TYPE_PNG

            }

            try {
                var bi: BufferedImage? = ImageIO.read(ByteArrayInputStream(imageBytes!!))
                width = bi!!.width.toShort()
                height = bi.height.toShort()
                bi = null
            } catch (ex: IOException) {
                Logger.logWarn("Java ImageIO could not decode image bytes for $imageName:$ex")
                if (false) {    // 20081028 KSC: don't overwrite original image bytes
                    val imgname = "failed_image_read_" + this.imageName + "." + imageFormat
                    val outimg = FileOutputStream(imgname)
                    outimg.write(imageBytes!!)
                    outimg.flush()
                    outimg.close()

                }
            }

            // 20070924 KSC: bounds[2] = width;
            // "" bounds[3] = height;
            this.imageName = "UnnamedImage"
        } catch (e: Exception) {
            Logger.logWarn("Problem creating ImageHandle:$e Please see BugTrack article: http://extentech.com/uimodules/docs/docs_detail.jsp?meme_id=1431&showall=true")
            if (false) { // debug image parse probs
                val imgname = "failed_image_read_" + this.imageName + "." + imageFormat
                try {
                    val outimg = FileOutputStream(imgname)
                    outimg.write(imageBytes!!)
                    outimg.flush()
                    outimg.close()
                } catch (ex: Exception) {
                }

            }
        }

    }

    /**
     * update the underlying image record
     */
    @Throws(Exception::class)
    fun update() {
        if (this.msodrawing == null)
            throw Exception("ImageHandle.Update: Image Not initialzed")
        this.msodrawing!!.updateRecord()    //this.thisMsodrawing.getSPID());
        this.sheet!!.workBook!!.updateMsodrawingHeaderRec(this.sheet)
    }

    /**
     * converts image data into byte array used by
     *
     *
     * Jan 22, 2010
     *
     * @param imagebytestream
     * @return
     */
    fun convertData(imagebytestream: InputStream): ByteArray? {
        try {
            val bos = ByteArrayOutputStream()

            val bi = ImageIO.read(imagebytestream)

            ImageIO.write(bi, "png", bos)

            return bos.toByteArray()
        } catch (e: Exception) {
            Logger.logErr("ImageHandle.convertData: $e")
            return null
        }

    }

    /**
     * returns the format name of the image data
     *
     *
     * Jan 22, 2010
     *
     * @param imagebytestream
     * @return
     */
    fun getImageFormat(imagebytestream: InputStream): String? {

        try {
            // Create an image input stream on the image
            val iis = ImageIO.createImageInputStream(imagebytestream)

            // Find all image readers that recognize the image format
            val iter = ImageIO.getImageReaders(iis)
            if (!iter.hasNext()) {

                return null
            }

            // Use the first reader
            val reader = iter.next() as ImageReader

            // Close stream
            iis.close()

            // Return the format name
            return reader.formatName
        } catch (e: IOException) {
        }

        // The image could not be read
        return null
    }

    /**
     * set the image x, w, width and height in pixels
     *
     * @param x
     * @param y
     * @param w
     * @param h
     */
    fun setCoords(x: Int, y: Int, w: Int, h: Int) {
        if (msodrawing != null)
            msodrawing!!.coords = shortArrayOf(x.toShort(), y.toShort(), w.toShort(), h.toShort())
        else {// save for later
            this.x = x.toShort()
            this.y = y.toShort()
            this.width = w.toShort()
            this.height = h.toShort()
        }
    }

    /**
     * set the image upper x coordinate in pixels
     *
     * @param x
     */
    fun setX(x: Int) {
        if (msodrawing != null)
            msodrawing!!.setX(Math.round(x / 6.4).toShort().toInt()) // 20090506 KSC: convert pixels to excel units
        else
        // save for later
            this.x = x.toShort()
    }

    /**
     * set the image upper y coordinate in pixels
     *
     * @param y
     */
    fun setY(y: Int) {
        if (msodrawing != null)
            msodrawing!!.setY(Math.round(y * 0.60).toShort().toInt())    //convert pixels to points
        else
        // save for later
            this.y = y.toShort()
    }

    /**
     * return the width of the image in pixels
     *
     * @return
     */
    fun getWidth(): Short {
        return Math.round(msodrawing!!.width * 6.4).toShort() // 20090506 KSC: Convert excel units to pixels
    }

    /**
     * set the width of the image in pixels
     *
     * @param w
     */
    fun setWidth(w: Int) {
        msodrawing!!.setWidth(Math.round(w / 6.4).toShort().toInt()) // 20090506 KSC: convert pixels to excel units
    }

    /**
     * return the height of the image in pixels
     *
     * @return
     */
    fun getHeight(): Short {
        return Math.round(msodrawing!!.height / 0.60).toShort() // 20090506 KSC: Convert points to pixels
    }

    /**
     * set the height of the image in pixels
     *
     * @param h
     */
    fun setHeight(h: Int) {
        msodrawing!!.setHeight(Math.round(h * 0.60).toShort().toInt()) // 20090506 KSC: convert pixels to points
    }

    /**
     * return the upper x coordinate of the image in pixels
     *
     * @return
     */
    fun getX(): Short {
        return Math.round(msodrawing!!.x * 6.4).toShort()  // 20090506 KSC: convert excel units to pixels
    }

    /**
     * return the upper y coordinate of the image in pixels
     *
     * @return
     */
    fun getY(): Short {
        return Math.round(msodrawing!!.y * 0.60).toShort()  // 20090506 KSC: convert points to pixels
    }

    /**
     * write the image bytes to an outputstream such as a file
     *
     * @param out
     */
    //Modified by Bikash
    @Throws(IOException::class)
    fun write(out: OutputStream) {
        // TODO: write the image bytes out
        out.write(this.imageBytes!!)
    }

    /**
     * Need to figure out a unique identifier for these -- if there
     * is one in the BIFF8 information, that would be best
     *
     * @see java.lang.Object.toString
     */
    override fun toString(): String {
        return name    // 20071026 KSC: imageName;
    }


    /**
     * removes this Image from the WorkBook.
     *
     * @return whether the removal was a success
     */
    fun remove(): Boolean {
        this.sheet!!.removeImage(this)
        // blow out the image rec
        this.msodrawing!!.remove(true)
        return true
    }

    /**
     * @return Returns the imageBytes.
     */
    fun getImageBytes(): ByteArray? {
        return imageBytes
    }

    /**
     * sets the underlying image bytes to the bytes from a new image
     *
     *
     * essentially swapping the image bytes for another, while retaining
     * all other image properties (borders, etc.)
     *
     * @param imageBytes The imageBytes to set.
     */
    fun setImageBytes(imageBytes: ByteArray) {
        this.imageBytes = imageBytes
        sheet!!.workBook!!.msoDrawingGroup!!.setImageBytes(this.imageBytes, sheet, this.msodrawing!!, this.name)
    }

    /**
     * set ImageHandle position based on another
     *
     * @param im source ImageHandle
     * @return
     */
    fun position(im: ImageHandle) {
        /* one way, just set x and y and keep original w and h
        //      set x and y, keep original width and height
        short[] origcoords= im.getCoords();
        short[] coords= this.getCoords();
        coords[0]= origcoords[0];
        coords[1]= origcoords[1];
        this.setCoords(coords[0], coords[1], coords[2], coords[3]);
        */
        /* other way, set with all original coordinates */
        this.bounds = im.bounds
    }

    /**
     * return the XML representation of this image
     *
     * @return String
     */
    fun getXML(rId: Int): String {
        val sb = StringBuffer()
        val bounds = this.bounds
        val EMU = 1270    // 1 pt= 1270 EMUs		-- for consistency with OOXML

        sb.append("<twoCellAnchor editAs=\"oneCell\">")
        sb.append("\r\n")
        // top left coords
        sb.append("<from>")
        sb.append("<col>" + bounds!![0] + "</col>")
        sb.append("<colOff>" + bounds[1] * EMU + "</colOff>")
        sb.append("<row>" + bounds[2] + "</row>")
        sb.append("<rowOff>" + bounds[3] * EMU + "</rowOff>")
        sb.append("</from>")
        sb.append("\r\n")
        // bottom right coords
        sb.append("<to>")
        sb.append("<col>" + bounds[4] + "</col>")
        sb.append("<colOff>" + bounds[5] * EMU + "</colOff>")
        sb.append("<row>" + bounds[6] + "</row>")
        sb.append("<rowOff>" + bounds[7] * EMU + "</rowOff>")
        sb.append("</to>")
        sb.append("\r\n")
        // Picture details - req. child elements= nvPicPr (non-visual picture properties), blipFill (links to image), spPr (shape properties)
        sb.append("<pic>")
        sb.append("\r\n")
        sb.append("<nvPicPr>")
        sb.append("<cNvPr id=\"" + this.msodrawing!!.spid + "\"")
        sb.append(" name=\"" + this.imageName + "\"")
        sb.append(" descr=\"" + this.shapeName + "\"/>")
        sb.append("<cNvPicPr>")
        sb.append("<picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>")
        sb.append("</cNvPicPr>")
        sb.append("</nvPicPr>")
        sb.append("\r\n")
        // Picture relationship Id and relationship to the package
        sb.append("<blipFill>")
        sb.append("<blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId$rId\"/>")
        //<a:srcRect/>	//If the picture is cropped, these details are stored in the <srcRect/> element
        sb.append("<stretch><fillRect/></stretch>")
        sb.append("</blipFill>")
        sb.append("\r\n")
        // shape properties
        sb.append("<spPr>"/* bwMode=\"auto\">"*/)
        sb.append("<xfrm>")
        val x = this.getX() * EMU
        val y = this.getY() * EMU
        val cx = this.getWidth() * EMU
        val cy = this.getHeight() * EMU
        sb.append("<off x=\"$x\" y=\"$y\"/>")        // offsets= location
        sb.append("<ext cx=\"$cx\" cy=\"$cy\"/>")    // extents= size of bounding box enclosing pic in EMUs
        sb.append("</xfrm>")
        sb.append("<prstGeom prst=\"rect\">")        // preset geometry, appears necessary for images ...
        sb.append("<avLst/>")
        sb.append("</prstGeom>")
        sb.append("<noFill/>")
        sb.append("</spPr>")
        sb.append("</pic>")
        sb.append("\r\n")
        sb.append("<clientData/>")
        sb.append("\r\n")
        sb.append("</twoCellAnchor>")
        return sb.toString()
    }

    /**
     * return the (00)XML (or DrawingML) representation of this image
     *
     * @param rId relationship id for image file
     * @return String
     */
    fun getOOXML(rId: Int): String {
        val t = TwoCellAnchor(this.editMovement)
        t.setAsImage(rId, this.imageName, this.shapeName, this.msodrawing!!.spid, this.spPr)
        t.bounds = TwoCellAnchor.convertBoundsFromBIFF8(this.sheet, this.bounds!!)    // adjust BIFF8 bounds to OOXML units
        return t.ooxml
        // missing in <xdr:pic>
        // <xdr:nvPicPr><xdr:cNvPicPr><a:picLocks noChangeArrowheads="1">
        // <xdr:blipFill><a:blip cstate="print">
        // <xdr:blipFill><a:srcRect/><a:stretch><a:fillRect/></a:stretch>


        /*
    	StringBuffer sb= new StringBuffer();
    	int[] bounds= twoCellAnchor.convertBoundsFromBIFF8(this.getSheet(), this.getBounds());	// adjust BIFF8 bounds to OOXML units
        final int EMU= 1270;	// 1 pt= 1270 EMUs

        // 20081008 KSC:  Added namespaces for Excel7 Use **********************
    	// TODO: create a twoCellAnchor from this ImageHandle and use twoCellAnchor.getOOXML
    	sb.append("<xdr:twoCellAnchor");
    		if (editMovement!=null) sb.append(" editAs=\"" + editMovement + "\"");	// how to resize or move upon editing
    		sb.append(">\r\n");
    		// top left coords
    		sb.append("<xdr:from>");
    			sb.append("<xdr:col>" + bounds[0] + "</xdr:col>");	// 1-based column
    			sb.append("<xdr:colOff>" + bounds[1]+ "</xdr:colOff>");
    			sb.append("<xdr:row>" + bounds[2] + "</xdr:row>");
    			sb.append("<xdr:rowOff>" + bounds[3] + "</xdr:rowOff>");
    		sb.append("</xdr:from>");	sb.append("\r\n");
    		// bottom right coords
    		sb.append("<xdr:to>");
    			sb.append("<xdr:col>" + bounds[4] + "</xdr:col>");		// 1-based column
    			sb.append("<xdr:colOff>" + bounds[5] + "</xdr:colOff>");
    			sb.append("<xdr:row>" + bounds[6] + "</xdr:row>");
    			sb.append("<xdr:rowOff>" + bounds[7] + "</xdr:rowOff>");
    		sb.append("</xdr:to>");		sb.append("\r\n");
    		// Picture details - req. child elements= nvPicPr (non-visual picture properties), blipFill (links to image), spPr (shape properties)
    		sb.append("<xdr:pic>");		sb.append("\r\n");
    			sb.append("<xdr:nvPicPr>");
    	        	sb.append("<xdr:cNvPr id=\"" + this.getMsodrawing().getSPID() + "\"");
    	        	sb.append(" name=\"" + this.getImageName() + "\"");
    	        	sb.append(" descr=\"" + this.getShapeName() + "\"/>");
    	        	sb.append("<xdr:cNvPicPr>");
    	        	sb.append("<a:picLocks noChangeAspect=\"1\"  noChangeArrowheads=\"1\"/>");
    	        	sb.append("</xdr:cNvPicPr>");
    	        sb.append("</xdr:nvPicPr>");	sb.append("\r\n");
    	        // Picture relationship Id and relationship to the package
    	        sb.append("<xdr:blipFill>");
    	        	sb.append("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId" + rId + "\"/>");
    	        	sb.append("<a:srcRect/>");			//If the picture is cropped, these details are stored in the <srcRect/> element
    	        	sb.append("<a:stretch><a:fillRect/></a:stretch>");
    	        sb.append("</xdr:blipFill>");	sb.append("\r\n");
    	        // shape properties
    	        if (imagesp!=null)
    	        	sb.append(imagesp.getOOXML());
    	        else { // default basic
	    	        sb.append("<xdr:spPr bwMode=\"auto\">");
    	        	sb.append("<a:xfrm>");
	    	        int x=this.getX()*EMU, y=this.getY()*EMU, cx= this.getWidth()*EMU, cy= this.getHeight()*EMU;
	    	        sb.append("<a:off x=\"" + x + "\" y=\"" + y + "\"/>");		// offsets= location
	    	        sb.append("<a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/>");	// extents= size of bounding box enclosing pic in EMUs
	    	        sb.append("</a:xfrm>");
	    	        sb.append("<a:prstGeom prst=\"rect\">");		// preset geometry, appears necessary for images ...
	    	        sb.append("<a:avLst/>");
	    	        sb.append("</a:prstGeom>");
	    	        sb.append("</xdr:spPr>");	sb.append("\r\n");
    	        }
    	    sb.append("</xdr:pic>");	sb.append("\r\n");
    	    sb.append("<xdr:clientData/>");	sb.append("\r\n");
    	sb.append("</xdr:twoCellAnchor>");
    	return sb.toString();
*/
    }

    /**
     * specify how to resize or move upon edit OOXML specific
     *
     * @param editMovement
     */
    fun setEditMovement(editMovement: String) {
        this.editMovement = editMovement
    }

    companion object {

        /**
         *
         */
        private const val serialVersionUID = 3177017738178634238L

        // coordinates
        val X = 0
        val Y = 1
        val WIDTH = 2
        val HEIGHT = 3
    }


}