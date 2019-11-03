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

import io.starter.toolkit.ByteTools

import java.util.Locale


/**
 * **SETUP 0xA1h: Describes the settings in the Page Setup (Printing) dialog.**<br></br>
 *
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       iPaperSize	2       Paper size
 * 6       iScale	    2       Scaling factor
 * 8       iPageStart  2       Starting page num
 * 10      iFitWidth   2       Fit to width; num pages
 * 12      iFitHeight  2       Fit to height; num pages
 * 14      grbit       2       Option flags
 * 16      iRes  		2       Print Resolution
 * 18      iVRes       2       Vertical Print Resolution
 * 20      numHdr      8       Header margin (IEEE number)
 * 28      numFtr      8       Footer margin (IEEE number)
 * 36      iCopies     2       Number of copies
 *
 *
 * offset  Bits   MASK     name        	contents
 * ---
 * 0       0   	0x0001  fLeftToRight   Print over, then down
 * 1   	0x0002  fLandscape     =0 Landscape 1 Portrait
 * 2		0x0004	fNoPls			=1 data not received from printer. settings invalid --
 * =A bit that specifies whether the iPaperSize, iScale, iRes, iVRes,
 * iCopies, fNoOrient, and fPortrait data are undefined and ignored.
 * If the value is 1, they are undefined and ignored.
 * 3		0x0008	fNoColor		=1 black and white
 * 4		0x0010	fDraft			=1 draft quality
 * 5		0x0020	fNotes			=1 print cell notes
 * 6		0x0040	fNoOrient		=1 orientation not set
 * 7		0x0080	fUsePage		=1 use custom starting page num, not Auto
 *
</pre> *
 *
 *
 *
 *
 * Paper Size:
 * Value	Meaning
 * 1		US Letter 8 1/2 x 11 in
 * 2		US Letter Small 8 1/2 x 11 in
 * 3		US Tabloid 11 x 17 in
 * 4		US Ledger 17 x 11 in
 * 5		US Legal 8 1/2 x 14 in
 * 6		US Statement 5 1/2 x 8 1/2 in
 * 7		US Executive 7 1/4 x 10 1/2 in
 * 8		A3 297 x 420 mm
 * 9		A4 210 x 297 mm
 * 10		A4 Small 210 x 297 mm
 * 11		A5 148 x 210 mm
 * 12		B4 (JIS) 250 x 354
 * 13		B5 (JIS) 182 x 257 mm
 * 14		Folio 8 1/2 x 13 in
 * 15		Quarto 215 x 275 mm
 * 16		10 x 14 in
 * 17		11 x 17 in
 * 18		US Note 8 1/2 x 11 in
 * 19		US Envelope #9 3 7/8 x 8 7/8
 * 20		US Envelope #10 4 1/8 x 9 1/2
 * 21		US Envelope #11 4 1/2 x 10 3/8
 * 22		US Envelope #12 4 \276 x 11
 * 23		US Envelope #14 5 x 11 1/2
 * 24		C size sheet
 * 25		D size sheet
 * 26		E size sheet
 * 27		Envelope DL 110 x 220mm
 * 28		Envelope C5 162 x 229 mm
 * 29		Envelope C3 324 x 458 mm
 * 30		Envelope C4 229 x 324 mm
 * 31		Envelope C6 114 x 162 mm
 * 32		Envelope C65 114 x 229 mm
 * 33		Envelope B4 250 x 353 mm
 * 34		Envelope B5 176 x 250 mm
 * 35		Envelope B6 176 x 125 mm
 * 36		Envelope 110 x 230 mm
 * 37		US Envelope Monarch 3.875 x 7.5 in
 * 38		6 3/4 US Envelope 3 5/8 x 6 1/2 in
 * 39		US Std Fanfold 14 7/8 x 11 in
 * 40		German Std Fanfold 8 1/2 x 12 in
 * 41		German Legal Fanfold 8 1/2 x 13 in
 * 42		B4 (ISO) 250 x 353 mm
 * 43		Japanese Postcard 100 x 148 mm
 * 44		9 x 11 in
 * 45		10 x 11 in
 * 46		15 x 11 in
 * 47		Envelope Invite 220 x 220 mm
 * 48		RESERVED--DO NOT USE
 * 49		RESERVED--DO NOT USE
 * 50		US Letter Extra 9 \275 x 12 in
 * 51		US Legal Extra 9 \275 x 15 in
 * 52		US Tabloid Extra 11.69 x 18 in
 * 53		A4 Extra 9.27 x 12.69 in
 * 54		Letter Transverse 8 \275 x 11 in
 * 55		A4 Transverse 210 x 297 mm
 * 56		Letter Extra Transverse 9\275 x 12 in
 * 57		SuperA/SuperA/A4 227 x 356 mm
 * 58		SuperB/SuperB/A3 305 x 487 mm
 * 59		US Letter Plus 8.5 x 12.69 in
 * 60		A4 Plus 210 x 330 mm
 * 61		A5 Transverse 148 x 210 mm
 * 62		B5 (JIS) Transverse 182 x 257 mm
 * 63		A3 Extra 322 x 445 mm
 * 64		A5 Extra 174 x 235 mm
 * 65		B5 (ISO) Extra 201 x 276 mm
 * 66		A2 420 x 594 mm
 * 67		A3 Transverse 297 x 420 mm
 * 68		A3 Extra Transverse 322 x 445 mm
 * 69		Japanese Double Postcard 200 x 148 mm
 * 70		A6 105 x 148 mm
 * 71		Japanese Envelope Kaku #2
 * 72		Japanese Envelope Kaku #3
 * 73		Japanese Envelope Chou #3
 * 74		Japanese Envelope Chou #4
 * 75		Letter Rotated 11 x 8 1/2 11 in
 * 76		A3 Rotated 420 x 297 mm
 * 77		A4 Rotated 297 x 210 mm
 * 78		A5 Rotated 210 x 148 mm
 * 79		B4 (JIS) Rotated 364 x 257 mm
 * 80		B5 (JIS) Rotated 257 x 182 mm
 * 81		Japanese Postcard Rotated 148 x 100 mm
 * 82		Double Japanese Postcard Rotated 148 x 200 mm
 * 83		A6 Rotated 148 x 105 mm
 * 84		Japanese Envelope Kaku #2 Rotated
 * 85		Japanese Envelope Kaku #3 Rotated
 * 86		Japanese Envelope Chou #3 Rotated
 * 87		Japanese Envelope Chou #4 Rotated
 * 88		B6 (JIS) 128 x 182 mm
 * 89		B6 (JIS) Rotated 182 x 128 mm
 * 90		12 x 11 in
 * 91		Japanese Envelope You #4
 * 92		Japanese Envelope You #4 Rotated
 * 93		PRC 16K 146 x 215 mm
 * 94		PRC 32K 97 x 151 mm
 * 95		PRC 32K(Big) 97 x 151 mm
 * 96		PRC Envelope #1 102 x 165 mm
 * 97		PRC Envelope #2 102 x 176 mm
 * 98		PRC Envelope #3 125 x 176 mm
 * 99		PRC Envelope #4 110 x 208 mm
 * 100		PRC Envelope #5 110 x 220 mm
 * 101		PRC Envelope #6 120 x 230 mm
 * 102		PRC Envelope #7 160 x 230 mm
 * 103		PRC Envelope #8 120 x 309 mm
 * 104		PRC Envelope #9 229 x 324 mm
 * 105		PRC Envelope #10 324 x 458 mm
 * 106		PRC 16K Rotated
 * 107		PRC 32K Rotated
 * 108		PRC 32K(Big) Rotated
 * 109		PRC Envelope #1 Rotated 165 x 102 mm
 * 110		PRC Envelope #2 Rotated 176 x 102 mm
 * 111		PRC Envelope #3 Rotated 176 x 125 mm
 * 112		PRC Envelope #4 Rotated 208 x 110 mm
 * 113		PRC Envelope #5 Rotated 220 x 110 mm
 * 114		PRC Envelope #6 Rotated 230 x 120 mm
 * 115		PRC Envelope #7 Rotated 230 x 160 mm
 * 116		PRC Envelope #8 Rotated 309 x 120 mm
 * 117		PRC Envelope #9 Rotated 324 x 229 mm
 * 118		PRC Envelope #10 Rotated 458 x 324 mm
 *
 * @see WorkBook
 *
 * @see XLSRecord
 */

/**
 *
 *
 */
class Setup : io.starter.formats.XLS.XLSRecord() {


    private var paperSize: Short = -1 // Paper size
    private var scale: Short = -1 // Scaling factor
    private var pageStart: Short = -1 // Starting page num
    private var fitWidth: Short = -1 // Fit to width; num pages
    private var fitHeight: Short = -1 // Fit to height; num pages
    private var grbit: Short = -1 // Option flags
    private var resolution: Short = -1 // Print Resolution
    private var verticalResolution: Short = -1 // Vertical Print Resolution
    private var headerMargin = -1.0 // Header margin (IEEE number)
    private var footerMargin = -1.0 // Footer margin (IEEE number)
    private var copies: Short = -1 // Number of copies

    /**
     *
     * @return Returns the draft.
     */
    /**
     *
     * @param draft The draft to set.
     */
    var draft: Boolean
        get() = grbit and BITMASK_DRAFT == BITMASK_DRAFT
        set(draft) {
            if (draft)
                grbit = (grbit or BITMASK_DRAFT).toShort()
            else {
                val grbittemp = (grbit xor BITMASK_DRAFT).toShort()
                grbit = (grbittemp and grbit).toShort()
            }
            this.setGrbit()
        }

    /**
     *
     * @return Returns the landscape.
     */
    /**
     *
     * @param landscape The landscape to set.
     */
    // This flag is actually whether it's in *portrait* mode
    // This flag is actually whether it's in *portrait* mode
    var landscape: Boolean
        get() {
            val gx = (grbit and BITMASK_LANDSCAPE).toShort()
            return gx.toInt() != BITMASK_LANDSCAPE
        }
        set(landscape) {
            if (!landscape)
                grbit = (grbit or BITMASK_LANDSCAPE).toShort()
            else {
                grbit = (grbit and BITMASK_LANDSCAPE.inv()).toShort()
            }
            this.setGrbit()
            setInitialized(true)
        }

    /**
     *
     * @return Returns the leftToRight.
     */
    /**
     *
     * @param leftToRight The leftToRight to set.
     */
    var leftToRight: Boolean
        get() = grbit and BITMASK_LEFTTORIGHT == BITMASK_LEFTTORIGHT
        set(leftToRight) {
            if (leftToRight)
                grbit = (grbit or BITMASK_LEFTTORIGHT).toShort()
            else
                grbit = (grbit xor BITMASK_LEFTTORIGHT).toShort()
            this.setGrbit()
        }

    /**
     *
     * @return Returns the noColor.
     */
    /**
     *
     * @param noColor The noColor to set.
     */
    var noColor: Boolean
        get() = grbit and BITMASK_NOCOLOR == BITMASK_NOCOLOR
        set(noColor) {
            if (noColor)
                grbit = (grbit or BITMASK_NOCOLOR).toShort()
            else
                grbit = (grbit xor BITMASK_NOCOLOR).toShort()
            this.setGrbit()
        }

    /**
     *
     * @return Returns the noOrient.
     */
    /**
     *
     * @param noOrient The noOrient to set.
     */
    var noOrient: Boolean
        get() = grbit and BITMASK_NOORIENT == BITMASK_NOORIENT
        set(noOrient) {
            if (noOrient)
                grbit = (grbit or BITMASK_NOORIENT).toShort()
            else
                grbit = (grbit and BITMASK_NOORIENT.inv()).toShort()
            this.setGrbit()
        }

    /**
     *
     * @return Returns the noPrintData.
     */
    val noPrintData: Boolean
        get() = grbit and BITMASK_NOPRINTDATA == BITMASK_NOPRINTDATA

    /**
     *
     * @return Returns whether to print Notes.
     */
    /**
     *
     * @param whether to print Notes
     */
    var printNotes: Boolean
        get() = grbit and BITMASK_PRINTNOTES == BITMASK_PRINTNOTES
        set(printNotes) {
            if (printNotes)
                grbit = (grbit or BITMASK_PRINTNOTES).toShort()
            else
                grbit = (grbit xor BITMASK_PRINTNOTES).toShort()
            this.setGrbit()
        }

    /**
     *
     * @return Returns whether to use a custom start page
     */
    /**
     *
     * @param usePage The usePage to set.
     */
    var usePage: Boolean
        get() = grbit and BITMASK_USEPAGE == BITMASK_USEPAGE
        set(usePage) {
            if (usePage)
                grbit = (grbit or BITMASK_USEPAGE).toShort()
            else
                grbit = (grbit xor BITMASK_USEPAGE).toShort()
            this.setGrbit()
        }


    override fun setSheet(sheet: Sheet?) {
        super.setSheet(sheet)
        (sheet as Boundsheet).addPrintRec(this)
    }

    /**
     *
     * @param isInitialized Set false if settings have changed
     */
    fun setInitialized(isInitialized: Boolean) {
        /**
         * meaning: Paper size, scaling factor, paper orientation (portrait/landscape),
         * print resolution and number of copies are NOT initialized
         */
        if (!isInitialized)
        // if isn't initialized
            grbit = (grbit or BITMASK_NOPRINTDATA).toShort()
        else if (grbit and BITMASK_NOPRINTDATA == BITMASK_NOPRINTDATA)
            grbit = (grbit xor BITMASK_NOPRINTDATA).toShort()
        this.setGrbit()
    }


    override fun init() {
        super.init()
        var pos = 0
        paperSize = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()) //Paper size
        scale = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()) //Scaling factor
        pageStart = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()) //Starting page num
        fitWidth = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()) //Fit to width; num pages
        fitHeight = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()) //Fit to height; num pages
        grbit = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()) //Option flags
        resolution = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt())    //Print Resolution
        verticalResolution = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()) //Vertical Print Resolution

        // IEEE fp numbers
        val bx1 = byteArrayOf(this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++))
        headerMargin = ByteTools.eightBytetoLEDouble(bx1) //Header margin (IEEE number)
        val bx2 = byteArrayOf(this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++), this.getByteAt(pos++))
        footerMargin = ByteTools.eightBytetoLEDouble(bx2) //Footer margin (IEEE number)
        copies = ByteTools.readShort(this.getByteAt(pos++).toInt(), this.getByteAt(pos++).toInt()) //Number of copies

        if (grbit and 0x4 == 0x4) {//	 Paper size, scaling factor, paper orientation (portrait/landscape), print resolution and number of copies are not initialised
            scale = 100
            grbit = grbit or 0x2    // set orientation to portrait
            // defaults for resolution?
            // default for paper size
            if ("AR AS BR BS BM BO CA CL CO CR CU DM DO EC GD GP GU GT GY HN HT JM KN KY LC MX NI PA PE PR PY SR SV TC US UM UY VC VE VI ".indexOf(Locale.getDefault().country + " ") == -1)
            //    		if ("AR AS BR BS BM BO CA CL CO CR CU DM DO EC GD GP GU GT GY HN HT JM KN KY LC MX NI PA PE PR PY SR SV TC US UM UY VC VE VI ".indexOf(Locale.GERMANY.getCountry() + " ") == -1)
                paperSize = 9    // A4
            else
                paperSize = 1    // Letter
        }
    }

    /**
     *
     * @return Returns the copies.
     */
    fun getCopies(): Short {
        return copies
    }

    /**
     *
     * @param copies The copies to set.
     */
    fun setCopies(c: Short) {
        copies = c
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(copies)
        System.arraycopy(b, 0, data!!, 32, 2)
        this.setData(data)
        setInitialized(true)
    }

    /**
     *
     * @return Returns the fitHeight.
     */
    fun getFitHeight(): Short {
        return fitHeight
    }

    /**
     *
     * @param fitHeight The fitHeight to set.
     */
    fun setFitHeight(f: Short) {
        fitHeight = f
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(fitHeight)
        System.arraycopy(b, 0, data!!, 8, 2)
        this.setData(data)
    }

    /**
     *
     * @return Returns the fitWidth.
     */
    fun getFitWidth(): Short {
        return fitWidth
    }

    /**
     *
     * @param fitWidth The fitWidth to set.
     */
    fun setFitWidth(f: Short) {
        fitWidth = f
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(fitWidth)
        System.arraycopy(b, 0, data!!, 6, 2)
        this.setData(data)

    }

    /**
     *
     * @return Returns the footerMargin.
     */
    fun getFooterMargin(): Double {
        return footerMargin
    }

    /**
     *
     * @param footerMargin The footerMargin to set.
     */
    fun setFooterMargin(f: Double) {
        footerMargin = f
        val data = this.getData()
        val b = ByteTools.doubleToLEByteArray(footerMargin)
        System.arraycopy(b, 0, data!!, 24, 8)
        this.setData(data)
    }

    /**
     *
     * @return Returns the headerMargin.
     */
    fun getHeaderMargin(): Double {
        return headerMargin
    }

    /**
     *
     * @param headerMargin The headerMargin to set.
     */
    fun setHeaderMargin(h: Double) {
        headerMargin = h
        val data = this.getData()
        val b = ByteTools.doubleToLEByteArray(headerMargin)
        System.arraycopy(b, 0, data!!, 16, 8)
        this.setData(data)
    }

    /**
     *
     * @return Returns the pageStart.
     */
    fun getPageStart(): Short {
        return pageStart
    }

    /**
     *
     * @param pageStart The pageStart to set.
     */
    fun setPageStart(p: Short) {
        pageStart = p
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(pageStart)
        System.arraycopy(b, 0, data!!, 4, 2)
        this.setData(data)
    }

    /**
     *
     * @return Returns the paperSize.
     */
    fun getPaperSize(): Short {
        return paperSize
    }

    /**
     *
     * @param paperSize The paperSize to set.
     */
    fun setPaperSize(p: Short) {
        paperSize = p
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(paperSize)
        System.arraycopy(b, 0, data!!, 0, 2)
        this.setData(data)
        setInitialized(true)
    }

    /**
     *
     * @return Returns the resolution.
     */
    fun getResolution(): Short {
        return resolution
    }

    /**
     *
     * @param resolution The resolution to set.
     */
    fun setResolution(r: Short) {
        resolution = r
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(resolution)
        System.arraycopy(b, 0, data!!, 12, 2)
        this.setData(data)
        setInitialized(true)
    }

    /**
     *
     * @return Returns the scale.
     */
    fun getScale(): Short {
        return scale
    }

    /**
     *
     * @param scale The scale to set.
     */
    fun setScale(s: Short) {
        scale = s
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(scale)
        System.arraycopy(b, 0, data!!, 2, 2)
        this.setData(data)
        setInitialized(true)
    }

    /**
     *
     * @return Returns the verticalResolution.
     */
    fun getVerticalResolution(): Short {
        return verticalResolution
    }

    /**
     *
     * @param verticalResolution The verticalResolution to set.
     */
    fun setVerticalResolution(v: Short) {
        verticalResolution = v
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(verticalResolution)
        System.arraycopy(b, 0, data!!, 14, 2)
        this.setData(data)
        setInitialized(true)
    }


    /*offset  Bits   MASK     name        	contents
	---
	0       0   	0x0001   fLeftToRight   Print over, then down
	       	1   	0x0002   fLandscape     =0 Landscape 1 Portrait
			2		0x0004	fNoPls			=1 data not received from printer. settings invalid
			3		0x0008	fNoColor		=1 black and white
			4		0x0010	fDraft			=1 draft quality
			5		0x0020	fNotes			=1 print cell notes
			6		0x0040	fNoOrient		=1 orientation not set
			7		0x0080	fUsePage		=1 use custom starting page num, not Auto
	*/

    /**
     * Apply all the borderLineStyle fields into the current border line styles flag
     *
     */
    fun setGrbit() {
        /* as grbit is already set, no need for below ... (besides it doesn't work)
    	grbit = 0x0; incorrect!
        short tempLR 		= (short)(BITMASK_LEFTTORIGHT 	<< 0x0);
        short tempFL 		= (short)(BITMASK_LANDSCAPE  	<< 0x1);
        short tempPD		= (short)(BITMASK_NOPRINTDATA 	<< 0x2);
        short tempNC 		= (short)(BITMASK_NOCOLOR 		<< 0x3);
        short tempDR 		= (short)(BITMASK_DRAFT 		<< 0x4);
        short tempPN 		= (short)(BITMASK_PRINTNOTES 	<< 0x5);
        short tempNO 		= (short)(BITMASK_NOORIENT 		<< 0x6);
        short tempUP 		= (short)(BITMASK_USEPAGE 		<< 0x7);
        grbit = (short)(grbit | tempLR);
        grbit = (short)(grbit | tempFL);
        grbit = (short)(grbit | tempPD);
        grbit = (short)(grbit | tempNC);
        grbit = (short)(grbit | tempDR);
        grbit = (short)(grbit | tempPN);
        grbit = (short)(grbit | tempNO);
        grbit = (short)(grbit | tempUP);
        */
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(grbit)
        System.arraycopy(b, 0, data!!, 10, 2)
        this.setData(data)
    }


    fun setGrbitOLD() {
        val data = this.getData()
        val b = ByteTools.shortToLEBytes(grbit)
        System.arraycopy(b, 0, data!!, 10, 2)
        this.setData(data)
    }

    companion object {

        private val serialVersionUID = 1835707090231884340L

        internal val BITMASK_LEFTTORIGHT = 0x0001
        internal val BITMASK_LANDSCAPE = 0x0002
        internal val BITMASK_NOPRINTDATA = 0x0004
        internal val BITMASK_NOCOLOR = 0x0008
        internal val BITMASK_DRAFT = 0x0010
        internal val BITMASK_PRINTNOTES = 0x0020
        internal val BITMASK_NOORIENT = 0x0040
        internal val BITMASK_USEPAGE = 0x0080
    }
}