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
package io.starter.formats.escher

import io.starter.OpenXLS.FormatHandle
import io.starter.formats.XLS.FormatConstants
import io.starter.formats.XLS.MSODrawingConstants
import io.starter.formats.XLS.XLSConstants
import io.starter.toolkit.ByteTools
import io.starter.toolkit.Logger

import java.util.Arrays
import java.util.LinkedHashMap

//0xf00b

/**
 * shape properties table
 */
class MsofbtOPT(fbt: Int, inst: Int, version: Int) : EscherRecord(fbt, inst, version) {
    internal var recordData = ByteArray(0)
    internal var bBackground: Boolean = false
    internal var bActive: Boolean = false
    internal var bPrint: Boolean = false
    internal var imageIndex = -1
    var fillColor: java.awt.Color? = null
        internal set
    var fillType = 0
        internal set
    internal var imageName: String? = ""
    internal var shapeName: String? = ""
    internal var alternateText = ""
    internal var lineprops: IntArray? = null            // Line properties -- weight, color, style ...
    internal var hasTextId = false    // true if OPT contains msofbtlTxid - necessary to calculate container lengths correctly - see MSODrawing.updateRecord
    internal var props = LinkedHashMap()        // parsed property table note: properties are ordered via property set;

    val borderLineWidth: Int
        get() = if (lineprops != null) lineprops!![LINEPROPTS_WEIGHT] else -1

    fun setImageIndex(value: Int) {
        imageIndex = value
        if (imageIndex > -1)
            setProperty(MSODrawingConstants.msooptpib, true, false, value, null)
        else
        // remove property
            props.remove(Integer.valueOf(MSODrawingConstants.msooptpib))

    }

    /**
     * returns true if this OPT subrecord contains an msofbtlTxid entry-
     * necessary to calculate container lengths correctly - see MSODrawing.updateRecord
     *
     * @return
     */
    fun hasTextId(): Boolean {
        return hasTextId
    }

    /**
     * set the image name for this shape
     * msooptpibName
     *
     * @param name
     */
    fun setImageName(name: String) {
        try {
            imageName = name
            if (imageName == null || imageName == "") {
                // remove property
                props.remove(Integer.valueOf(MSODrawingConstants.msooptpibName))
                return
            }
            var imageNameBytes = name.toByteArray(charset(XLSConstants.UNICODEENCODING))
            val newbytes = ByteArray(imageNameBytes.size)
            System.arraycopy(imageNameBytes, 0, newbytes, 0, imageNameBytes.size)
            imageNameBytes = newbytes
            setProperty(MSODrawingConstants.msooptpibName, true, true, imageNameBytes.size, imageNameBytes)
        } catch (e: Exception) {
            Logger.logErr("Msofbt.setImageName failed.", e)
        }

    }

    /**
     * set the shape name atom in this OPT record
     *
     * @param name
     */
    fun setShapeName(name: String) {
        try {
            shapeName = name
            if (shapeName == null || shapeName == "") {
                // remove property
                props.remove(Integer.valueOf(MSODrawingConstants.msooptwzName))
                return
            }
            var shapeNameBytes = name.toByteArray(charset(XLSConstants.UNICODEENCODING))
            val newbytes = ByteArray(shapeNameBytes.size)
            System.arraycopy(shapeNameBytes, 0, newbytes, 0, shapeNameBytes.size)
            shapeNameBytes = newbytes
            setProperty(MSODrawingConstants.msooptwzName, true, true, shapeNameBytes.size, shapeNameBytes)
        } catch (e: Exception) {
            Logger.logErr("Msofbt.setShapeName failed.", e)
        }

    }

    /**
     * generate the recordData from the stored props hashmap if anything has changed
     */
    override fun getData(): ByteArray {
        if (isDirty) {    // regenerate recordData as contents have changed
            val tmp = ByteArray(inst * 6)        // basic property table
            var complexData = ByteArray(0)    // extra complex data, if any, after basic property table
            var pos = 0
            // try to extract properties in order
            val keys = java.util.ArrayList(props.keys)
            // try in numerical order-- appears to MOSTLY be the case ...
            val k = keys.toTypedArray()
            Arrays.sort(k)

            // write out properties in (numerical) order
            for (i in k.indices) {
                val propId = k[i] as Int
                val o = props.get(propId) as Array<Any>
                val isComplex = (o[0] as Boolean).booleanValue()
                val isBid = (o[2] as Boolean).booleanValue()
                var flag = 0
                if (isComplex) flag = flag or 0x80
                if (isBid) flag = flag or 0x40
                val dtx: Int
                if (!isComplex)
                    dtx = (o[1] as Int).toInt()    // non-complex data is just an integer
                else {
                    dtx = (o[1] as ByteArray).size + 2        // stored data is a byte array; get length + 2
                    complexData = ByteTools.append(o[1] as ByteArray, complexData)
                    complexData = ByteTools.append(byteArrayOf(0, 0), complexData)
                }
                // the basic part of the property table
                tmp[pos++] = (0xFF and propId).toByte()
                tmp[pos++] = (flag or (0x3F00 and propId shr 8)).toByte()
                val dtxBytes = ByteTools.cLongToLEBytes(dtx)
                System.arraycopy(dtxBytes, 0, tmp, pos, 4)
                pos += 4
            }
            recordData = ByteArray(tmp.size + complexData.size)
            System.arraycopy(tmp, 0, recordData, 0, tmp.size)
            // after the basic property table (PropID, IsBID, IsCOMPEX, dtx), store the complex data
            System.arraycopy(complexData, 0, recordData, tmp.size, complexData.size)
            isDirty = false
        }
        this.length = recordData.size
        return recordData
    }

    fun setData(b: ByteArray) {
        recordData = b
        props.clear()
        imageIndex = -1
        imageName = ""
        shapeName = ""
        alternateText = ""
        parseData()
    }

    /**
     * given property table bytes, parse into props hashmap
     */
    private fun parseData() {
        /*
         * First part of an OPT record is an array of FOPTEs (propertyId, fBid, fComplex, data)
         * If fComplex is set, the actual data (Unicode strings, arrays, etc.) is stored AFTER the last FOPTE (sorted by property id???);
         * the length of the complex data is stored in the data field.
         * if fComplex is not set, the meaining of the data field is dependent upon the propertyId
         * if fBid is set and fComplex is not set, the data = a BLIP id (= an index into the BLIP store)
         * The number of FOPTES is the inst field read above
         */
        var propertyId: Int
        var fBid: Int
        var fComplex: Int
        //int n= inst;				// number of properties to parse
        var pos = 0                    // pointer to current property in data/property table
        if (inst == 0 && recordData.size > 0) {    // called from GelFrame ...
            val dat = ByteArray(8)    // read header
            System.arraycopy(recordData, 0, dat, 0, 8)
            version = 0xF and dat[0]    // 15 for containers, version for atoms
            inst = 0xFF and dat[1] shl 4 or (0xF0 and dat[0] shr 4)
            fbt = 0xFF and dat[3] shl 8 or (0xFF and dat[2])    // record type id==0xF00B
            pos = 8    // skip header
        }
        for (i in 0 until inst) {
            propertyId = 0x3F and recordData[pos + 1] shl 8 or (0xFF and recordData[pos])    // 14 bits
            fBid = 0x40 and recordData[pos + 1] shr 6        // specifies whether the value in the dtx field is a BLIP identifier- only valid if fComplex= FALSE
            fComplex = 0x80 and recordData[pos + 1] shr 7    // if complex property, value is length.  Data is parsed after.
            val dtx = ByteTools.readInt(recordData[pos + 2], recordData[pos + 3], recordData[pos + 4], recordData[pos + 5])
            // TODO: if property number is of type bool/long/msoarray/... parse accordingly ..
            if (propertyId == MSODrawingConstants.msooptpib)
            // blip to display
                imageIndex = dtx
            else if (propertyId == MSODrawingConstants.msooptFillType) {
                fillType = dtx
            } else if (propertyId == MSODrawingConstants.msooptfillColor)
                fillColor = setFillColor(dtx)
            else if (propertyId == MSODrawingConstants.msooptfBackground)
                bBackground = dtx != 0
            else if (propertyId == MSODrawingConstants.msooptGroupShapeProperties) {
                //bPrint= (dtx!=0); // NOT TRUE!! TODO: parse real GroupShapeProperties (many)
            } else if (propertyId == MSODrawingConstants.msooptpictureActive)
                bActive = dtx != 0
            else if (propertyId == MSODrawingConstants.msooptlineWidth) {    // appears that this controls display of line
                if (lineprops == null) lineprops = IntArray(3)
                lineprops[LINEPROPTS_WEIGHT] = dtx
            } else if (propertyId == MSODrawingConstants.msooptlineColor) {    // appears that this is always present, even if no line
                if (lineprops == null) lineprops = IntArray(3)
                lineprops[LINEPROPTS_COLOR] = dtx
            } else if (propertyId == MSODrawingConstants.msooptLineStyle) {
                if (lineprops == null) lineprops = IntArray(3)
                lineprops[LINEPROPTS_STYLE] = dtx
            } else if (propertyId == MSODrawingConstants.msofbtlTxid) {
                hasTextId = true
            } // msooptFillWidth
            props.put(Integer.valueOf(propertyId), arrayOf(java.lang.Boolean.valueOf(fComplex != 0), Integer.valueOf(dtx), java.lang.Boolean.valueOf(fBid != 0)))
            pos += 6
        }

        // now parse complex data after all "tightly packed" properties have been parsed.  Order of data is original order
        val ii = props.keys.iterator()
        while (ii.hasNext()) {
            val propId = ii.next() as Int
            val o = props.get(propId) as Array<Any>    // Object[]:  0= isComplex, 1= dtx (value or len of complex data -- filled in below), 2= isBid
            if ((o[0] as Boolean).booleanValue()) {
                val len = (o[1] as Int).toInt()
                if (len >= 2) {
                    // apparently each record is delimited by a double byte 0 -- so decrement by 2 here and increment pos by 2 below
                    val complexdata = ByteArray(len - 2)    // retrieve complex data at end of record
                    System.arraycopy(recordData, pos, complexdata, 0, complexdata.size)    // get property data after main property table
                    props.put(propId, arrayOf(o[0], complexdata, o[2]))    //store complex data for later retrieval
                    if (propId == MSODrawingConstants.msooptpibName) { // = image name
                        try {
                            imageName = String(complexdata, XLSConstants.UNICODEENCODING)
                        } catch (e: Exception) {
                            imageName = "Unnamed"
                        }

                    } else if (propId == MSODrawingConstants.msooptwzName) { // = shape name
                        try {
                            shapeName = String(complexdata, XLSConstants.UNICODEENCODING)
                        } catch (e: Exception) {
                        }

                    } else if (propId == MSODrawingConstants.msooptwzDescription) { // = Alternate Text
                        try {
                            alternateText = String(complexdata, XLSConstants.UNICODEENCODING)
                        } catch (e: Exception) {
                        }

                    } else if (propId == MSODrawingConstants.msooptFillBlipName) {    // = the comment, file name, or the full URL that is used as a fill
                        try {
                            val fillName = String(complexdata, XLSConstants.UNICODEENCODING)
                        } catch (e: Exception) {
                        }

                    }

                    pos += complexdata.size + 2
                }
            }
        }
    }

    /**
     * @param propId       msofbtopt property ide see Msoconstants
     * @param isBid        value is a BLIP id - only valid if isComplex is false
     * @param isComplex    if false, dtx is used; if true, complexBytes are used and dtx=length
     * @param dtx          if not iscomplex, the value of property id; if iscomplex, length of complex data following the property table
     * @param complexBytes if iscomplex, holds value of complex property e.g. shape name
     */
    fun setProperty(propId: Int, isBid: Boolean, isComplex: Boolean, dtx: Int, complexBytes: ByteArray?) {
        // a general order of common properties is (via property id):
        /*
         * 127
         * 267
         * 261
         * 262
         * 128
         * 133
         * 139
         * 191
         * 385
         * 447
         * 448
         * 459
         * 511
         */
        if (isComplex)
        // complexBytes shouldn't be null
            props.put(Integer.valueOf(propId), arrayOf<Any>(java.lang.Boolean.valueOf(isComplex), complexBytes, java.lang.Boolean.valueOf(isBid)))
        else
            props.put(Integer.valueOf(propId), arrayOf(java.lang.Boolean.valueOf(isComplex), Integer.valueOf(dtx), java.lang.Boolean.valueOf(isBid)))

        this.inst = props.size
        isDirty = true    // flag to regenerate recordData
    }

    fun hasBorder(): Boolean {
        return lineprops != null && lineprops!![LINEPROPTS_WEIGHT] > 1
    }

    fun getImageIndex(): Int {
        return imageIndex
    }

    fun getImageName(): String? {
        return imageName
    }

    fun getShapeName(): String? {
        return shapeName
    }

    /**
     * Debug Output -- For Internal Use Only
     */
    fun debugOutput(): String {
        var propertyId: Int
        val log = StringBuffer()
        /*		java.util.ArrayList keys= new java.util.ArrayList(props.keySet());
		int n= keys.size();
		for (int i= 0; i < n; i++) {
			log.append("\r\n");
			propertyId = ((Integer)keys.get(i)).intValue(); //   (0x3F&recordData[pos+1])<<8|(0xFF&recordData[pos]);
			Object[] o= (Object[]) props.get(keys.get(i));
			boolean isComplex= ((Boolean)o[0]).booleanValue();
			boolean isBid= ((Boolean)o[2]).booleanValue();
			if (isComplex) fComplex=1;
			if (isBid)	fBid= 1;
			int dtx;
			if (!isComplex)
				dtx= ((Integer)o[1]).intValue();	// non-complex data is just an integer
			else
				dtx= ((byte[])o[1]).length + 2;		// stored data is a byte array; get length + 2
			//fBid =  ((0x40&recordData[pos+1])>>6);		// value is a BLIP ID - only valid if fComplex= FALSE
			//fComplex = ((0x80&recordData[pos+1])>>7);	// if complex property, value is length.  Data is parsed after.
			//int dtx = ByteTools.readInt(recordData[pos+2],recordData[pos+3],recordData[pos+4],recordData[pos+5]);
			log.append("\t\t" + propertyId + "/" + fBid + "/" + fComplex + "/" + dtx);
			//pos+=6;
		}
*/
        val n = inst                // number of properties to parse
        // pointer to current property in data/property table
        var fBid = 0
        var fComplex = 0
        var end = recordData.size
        var pos = 0
        while (pos < end) {
            propertyId = 0x3F and recordData[pos + 1] shl 8 or (0xFF and recordData[pos])
            fBid = 0x40 and recordData[pos + 1] shr 6
            fComplex = 0x80 and recordData[pos + 1] shr 7    // if complex property, value is length.  Data is parsed after.
            val dtx = ByteTools.readInt(recordData[pos + 2], recordData[pos + 3], recordData[pos + 4], recordData[pos + 5])
            if (fComplex != 0)
                end -= dtx
            log.append("\t\t$propertyId/$fBid/$fComplex/$dtx")
            pos += 6
        }
        return log.toString()
    }

    /**
     * Interpret an OfficeArtCOLORREF (used in fillColor, lineColor and many other opts)
     *
     * @param clrStructure <br></br> More information
     * The OfficeArtCOLORREF structure specifies a color. The high 8 bits MAY be set to 0xFF, in which case the color MUST be ignored.
     * The color properties that are specified in the following table have a set of extended-color properties. The color property specifies the main color.
     * The colorExt and colorExtMod properties specify the extended colors that can be used to define the main color more precisely.
     * If neither extended-color property is set, the main color property contains the full color definition.
     * Otherwise, the colorExt property specifies the base color, and the colorExtMod property specifies a tint or shade modification that is applied to the colorExt property.
     * In this case, the main color property contains the flattened RGB color that is computed by applying the specified tint or shade modification to the specified base color.
     *
     *
     *
     *
     * A - unused1 (1 bit): A bit that is undefined and MUST be ignored.
     *
     *
     * B - unused2 (1 bit): A bit that is undefined and MUST be ignored.
     *
     *
     * C - unused3 (1 bit): A bit that is undefined and MUST be ignored.
     *
     *
     * D - fSysIndex (1 bit): A bit that specifies whether the system color scheme will be used to determine the color.
     * A value of 0x1 specifies that green and red will be treated as an unsigned 16-bit index into the system color table. Values less than 0x00F0 map directly to system colors.
     * For more information, see [MSDN-GetSysColor] (below)
     * The following table specifies values that have special meaning.
     * Value		Meaning
     * 0x00F0		Use the fill color of the shape.
     * 0x00F1		If the shape contains a line, use the line color of the shape. Otherwise, use the fill color.
     * 0x00F2		Use the line color of the shape.
     * 0x00F3		Use the shadow color of the shape.
     * 0x00F4	    Use the current, or last-used, color.
     * 0x00F5	    Use the fill background color of the shape.
     * 0x00F6	    Use the line background color of the shape.
     * 0x00F7	    If the shape contains a fill, use the fill color of the shape. Otherwise, use the line color.
     * The following table specifies values that indicate special procedural properties that are used to modify the color components of another color.
     * These values are combined with those in the preceding table or with a user-specified color. The first six values are mutually exclusive.
     * Value		Meaning
     * 0x0100	    Darken the color by the value that is specified in the blue field. A blue value of 0xFF specifies that the color is to be left unchanged, whereas a blue value of 0x00 specifies that the color is to be completely darkened.
     * 0x0200	    Lighten the color by the value that is specified in the blue field. A blue value of 0xFF specifies that the color is to be left unchanged, whereas a blue value of 0x00 specifies that the color is to be completely lightened.
     * 0x0300  	Add a gray level RGB value. The blue field contains the gray level to add:    NewColor = SourceColor + gray
     * 0x0400	    Subtract a gray level RGB value. The blue field contains the gray level to subtract:	NewColor = SourceColor - gray
     * 0x0500		Reverse-subtract a gray level RGB value. The blue field contains the gray level from which to subtract:	    NewColor = gray - SourceColor
     * 0x0600	    If the color component being modified is less than the parameter contained in the blue field, set it to the minimum intensity.
     * If the color component being modified is greater than or equal to the parameter, set it to the maximum intensity.
     * 0x2000	    After making other modifications, invert the color.
     * 0x4000	    After making other modifications, invert the color by toggling just the high bit of each color channel.
     * 0x8000      Before making other modifications, convert the color to grayscale.
     * E - fSchemeIndex (1 bit): A bit that specifies whether the current application-defined color scheme will be used to determine the color.
     * A value of 0x1 specifies that red will be treated as an index into the current color scheme table. If this value is 0x1, green and blue MUST be 0x00.
     * F - fSystemRGB (1 bit): A bit that specifies whether the color is a standard RGB color. The following table specifies the meaning of each value for this field.
     * Value		Meaning
     * 0x0			The RGB color MAY use halftone dithering to display.
     * 0x1		    The color MUST be a solid color.
     * G - fPaletteRGB (1 bit): A bit that specifies whether the current palette will be used to determine the color.
     * A value of 0x1 specifies that red, green, and blue contain an RGB value that will be matched in the current color palette. This color MUST be solid.
     * H - fPaletteIndex (1 bit): A bit that specifies whether the current palette will be used to determine the color.
     * A value of 0x1 specifies that green and red will be treated as an unsigned 16-bit index into the current color palette. This color MAY<1> be dithered.
     * If this value is 0x1, blue MUST be 0x00.
     * blue (1 byte): An unsigned integer that specifies the intensity of the blue color channel. A value of 0x00 has the minimum blue intensity. A value of 0xFF has the maximum blue intensity.
     * green (1 byte): An unsigned integer that specifies the intensity of the green color channel. A value of 0x00 has the minimum green intensity. A value of 0xFF has the maximum green intensity.
     * red (1 byte): An unsigned integer that specifies the intensity of the red color channel. A value of 0x00 has the minimum red intensity. A value of 0xFF has the maximum red intensity.
     *
     *
     * ...
     *
     *
     * MSDN-GetSysColor
     * Value						Meaning
     * COLOR_3DDKSHADOW	21		Dark shadow for three-dimensional display elements.
     * COLOR_3DFACE		15		Face color for three-dimensional display elements and for dialog box backgrounds.
     * COLOR_3DHIGHLIGHT	20		Highlight color for three-dimensional display elements (for edges facing the light source.)
     * COLOR_3DHILIGHT		20		Highlight color for three-dimensional display elements (for edges facing the light source.)
     * COLOR_3DLIGHT		22		Light color for three-dimensional display elements (for edges facing the light source.)
     * COLOR_3DSHADOW		16		Shadow color for three-dimensional display elements (for edges facing away from the light source).
     * COLOR_ACTIVEBORDER	10		Active window border.
     * COLOR_ACTIVECAPTION	2		Active window title bar.	Specifies the left side color in the color gradient of an active window's title bar if the gradient effect is enabled.
     * COLOR_APPWORKSPACE	12		Background color of multiple document interface (MDI) applications.
     * COLOR_BACKGROUND	1		Desktop.
     * COLOR_BTNFACE		15		Face color for three-dimensional display elements and for dialog box backgrounds.
     * COLOR_BTNHIGHLIGHT	20		Highlight color for three-dimensional display elements (for edges facing the light source.)
     * COLOR_BTNHILIGHT	20		Highlight color for three-dimensional display elements (for edges facing the light source.)
     * COLOR_BTNSHADOW		16		Shadow color for three-dimensional display elements (for edges facing away from the light source).
     * COLOR_BTNTEXT		18		Text on push buttons.
     * COLOR_CAPTIONTEXT	9		Text in caption, size box, and scroll bar arrow box.
     * COLOR_DESKTOP		1		Desktop.
     * COLOR_GRADIENTACTIVECAPTION	27	Right side color in the color gradient of an active window's title bar.
     * COLOR_ACTIVECAPTION specifies the left side color. Use SPI_GETGRADIENTCAPTIONS with the SystemParametersInfo function to determine whether the gradient effect is enabled.
     * COLOR_GRADIENTINACTIVECAPTION	28	Right side color in the color gradient of an inactive window's title bar. COLOR_INACTIVECAPTION specifies the left side color.
     * COLOR_GRAYTEXT		17		Grayed (disabled) text. This color is set to 0 if the current display driver does not support a solid gray color.
     * COLOR_HIGHLIGHT		13		Item(s) selected in a control.
     * COLOR_HIGHLIGHTTEXT	14		Text of item(s) selected in a control.
     * COLOR_HOTLIGHT		26		Color for a hyperlink or hot-tracked item.
     * COLOR_INACTIVEBORDER11		Inactive window border.
     * COLOR_INACTIVECAPTION3		Inactive window caption.
     * Specifies the left side color in the color gradient of an inactive window's title bar if the gradient effect is enabled.
     * COLOR_INACTIVECAPTIONTEXT19	Color of text in an inactive caption.
     * COLOR_INFOBK		24		Background color for tooltip controls.
     * COLOR_INFOTEXT		23		Text color for tooltip controls.
     * COLOR_MENU			4		Menu background.
     * COLOR_MENUHILIGHT	29		The color used to highlight menu items when the menu appears as a flat menu (see SystemParametersInfo). The highlighted menu item is outlined with COLOR_HIGHLIGHT.
     * Windows 2000:  This value is not supported.
     * COLOR_MENUBAR		30		The background color for the menu bar when menus appear as flat menus (see SystemParametersInfo). However, COLOR_MENU continues to specify the background color of the menu popup.
     * Windows 2000:  This value is not supported.
     * COLOR_MENUTEXT		7		Text in menus.
     * COLOR_SCROLLBAR		0		Scroll bar gray area.
     * COLOR_WINDOW		5		Window background.
     * COLOR_WINDOWFRAME	6		Window frame.
     * COLOR_WINDOWTEXT	8		Text in windows.
     */
    private fun setFillColor(clrStructure: Int): java.awt.Color? {
        val b = ByteTools.longToByteArray(clrStructure.toLong())
        val bPaletteIndex: Boolean
        val bSchemeIndex: Boolean
        val bSysIndex: Boolean
        var fillclr: Short

        bPaletteIndex = b[4] and 0x1 == 0x1    // 	specifies whether the current palette will be used to determine the color
        bSchemeIndex = b[4] and 0x8 == 0x8    //  specifies whether the current application defined color scheme will be used to determine the color.
        bSysIndex = b[4] and 0x10 == 0x10    //  specifies whether the system color scheme will be used to determine the color.

        if (bPaletteIndex) {    // // GREEN and RED are treated as an unsigned 16-bit index into the current color palette. This color MAY be dithered. BLUE MUST be 0x00.
        }
        if (bSchemeIndex) {        //  RED is an index into the current scheme color table. GREEN and BLUE MUST be 0x00.
            fillclr = b[7].toShort()            // what does 80 mean??????
            if (fillclr > FormatHandle.COLORTABLE.size)
                fillclr = FormatHandle.interpretSpecialColorIndex(fillclr.toInt())
            fillColor = FormatHandle.COLORTABLE[fillclr]
            return fillColor
        }
        if (bSysIndex) {    // GREEN and RED will be treated as an unsigned 16-bit index into the system color table. Values less than 0x00F0 map directly to system colors.
            fillclr = ByteTools.readShort(b[6].toInt(), b[7].toInt())
            if (fillclr.toInt() == 0x00F0 //		Use the fill color of the shape.

                    || fillclr.toInt() == 0x00F1        //If the shape contains a line, use the line color of the shape. Otherwise, use the fill color.

                    || fillclr.toInt() == 0x00F2        //Use the line color of the shape.

                    || fillclr.toInt() == 0x00F3        //Use the shadow color of the shape.

                    || fillclr.toInt() == 0x00F4        //Use the current, or last-used, color.

                    || fillclr.toInt() == 0x00F5        //Use the fill background color of the shape.

                    || fillclr.toInt() == 0x00F6        //Use the line background color of the shape.

                    || fillclr.toInt() == 0x00F7)
            //If the shape contains a fill, use the fill color of the shape. Otherwise, use the line color.
                fillclr = FormatConstants.COLOR_WHITE.toShort()
            if (fillclr.toInt() == 0x40)
            // default fg color
                fillclr = FormatConstants.COLOR_WHITE.toShort()
            else if (fillclr.toInt() == 0x41)
            // default bg color
                fillclr = FormatConstants.COLOR_WHITE.toShort()
            else if (fillclr.toInt() == 0x4D) {        // default CHART fg color -- INDEX SPECIFIC!
                fillColor = null    // flag to map via series (bar) color defaults
                return fillColor
            } else if (fillclr.toInt() == 0x4E)
            // default CHART fg color
                fillclr = FormatConstants.COLOR_WHITE.toShort()
            else if (fillclr.toInt() == 0x4F)
            // chart neutral color == black
                fillclr = FormatConstants.COLOR_BLACK.toShort()

            if (fillclr < 0 || fillclr > FormatHandle.COLORTABLE.size)
                fillclr = FormatConstants.COLOR_WHITE.toShort()
            fillColor = FormatHandle.COLORTABLE[fillclr]
            return fillColor
        }

        // otherwise, r, g and blue are color values 0-255
        val bl = if (b[5] < 0) 255 + b[5] else b[5]
        val g = if (b[6] < 0) 255 + b[6] else b[6]
        val r = if (b[7] < 0) 255 + b[7] else b[7]
        fillColor = java.awt.Color(r, g, bl)
        return fillColor
    }

    companion object {
        /**
         * serialVersionUID
         */
        private val serialVersionUID = 465530579513265882L
        internal val LINEPROPTS_STYLE = 0
        internal val LINEPROPTS_WEIGHT = 1
        internal val LINEPROPTS_COLOR = 2
    }
}
