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

import io.starter.formats.XLS.*

/**
 * The PrinterSettingsHandle gives you control over the printer settings for a Sheet such as whether to print in landscape or portrait mode.
 * <br></br><br></br>
 * The PrinterSettingsHandle provides fine-grained control over printing settings
 * in Excel.
 * <br></br><br></br>
 * NOTE: you can only view the effects of these methods in an open Excel file when
 * you use the "Print Setup" command.
 * <br></br><br></br>
 * OpenXLS does not currently support directly sending
 * spreadsheet data to a printer.
 *
 *
 * <br></br><br></br>**
 * Example Usage:**
 * <pre>
 * ...
 * PrinterSettingsHandle printersetup = sheet.getPrinterSettings();
 * // Paper Size
 * printersetup.setPaperSize(PrinterSettingsHandle.PAPER_SIZE_LEDGER_17x11);
 * // Scaling
 * printersetup.setScale(125);
 * //	resolution
 * printersetup.setResolution(300);
 * ...
</pre> *
 */
/* order of records in Page Settings Block:
 * HORIZONTALPAGEBREAKS
   VERTICALPAGEBREAKS
   HEADER
   FOOTER
   HCENTER
   VCENTER
   LEFTMARGIN
   RIGHTMARGIN
   TOPMARGIN
   BOTTOMMARGIN
   PLS
   SETUP
   BITMAP
 *
 */
class PrinterSettingsHandle
/**
 * default constructor
 */
(internal var sheet: Boundsheet) : Handle {

    private var printerSettings: Setup? = null
    private var hCenter: HCenter? = null
    private var vCenter: VCenter? = null
    private var leftMargin: LeftMargin? = null
    private var rightMargin: RightMargin? = null
    private var topMargin: TopMargin? = null
    private var bottomMargin: BottomMargin? = null
    private var grid: PrintGrid? = null
    private var headers: PrintRowCol? = null
    private var wsBool: WsBool? = null

    // the following are unimplemented printer setting recs:
    // HORIZONTALPAGEBREAKS;
    // VERTICALPAGEBREAKS;


    /**
     * get the number of copies to print
     *
     * @return the number of copies
     */
    val copies: Short
        get() = printerSettings!!.copies


    /**
     * get the footer margin size in inches
     *
     * @return the footer margin in inches
     */
    /**
     * sets the footer margin in inches
     *
     * @param footer margin
     */
    var footerMargin: Double
        get() = printerSettings!!.footerMargin
        set(f) {
            printerSettings!!.footerMargin = f
        }

    /**
     * get the header margin size in inches
     *
     * @return the header margin in inches
     */
    /**
     * sets the Header margin in inches
     *
     * @param header margin
     */
    var headerMargin: Double
        get() = printerSettings!!.headerMargin
        set(h) {
            printerSettings!!.headerMargin = h
        }

    /**
     * get the landscape orientation
     *
     * @return whether the print orientation is set to landscape
     */
    /**
     * set the print orientation to landscape or portrait
     *
     * @param landscape
     */
    // use the orientation setting
    var landscape: Boolean
        get() = printerSettings!!.landscape
        set(b) {
            printerSettings!!.noOrient = false
            printerSettings!!.landscape = b
        }

    /**
     * get the left-to-right print orientation
     *
     * @return whether the print orientation is set to left-to-right
     */
    /**
     * set the print orientation to left-to-right printing
     *
     * @param leftToRight
     */
    var leftToRight: Boolean
        get() = printerSettings!!.leftToRight
        set(b) {
            printerSettings!!.leftToRight = b
        }

    /**
     * get whether printing is in black and white
     *
     * @return black and white
     */
    /**
     * sets the output to black and white
     *
     * @param noColor
     */
    var noColor: Boolean
        get() = printerSettings!!.noColor
        set(b) {
            printerSettings!!.noColor = b
        }

    /**
     * get whether to ignore orientation
     *
     * @return ignore orientation
     */
    val noOrient: Boolean
        get() = printerSettings!!.noOrient

    /**
     * get whether printer data is missing
     *
     * @return whether the printer data is missing
     */
    val noPrintData: Boolean
        get() = printerSettings!!.noPrintData

    /**
     * get the page to start printing from
     *
     * @return the page to start printing from
     */
    /**
     * set the default page to start printing from
     *
     * @param p
     */
    var pageStart: Short
        get() = printerSettings!!.pageStart
        set(p) {
            printerSettings!!.pageStart = p
        }

    /**
     * Returns the paper size setting for the printer setup based on the
     * following table:
     *
     * @return paper size
     */
    val paperSize: Short
        get() = printerSettings!!.paperSize

    /**
     * @return whether to print Notes.
     */
    /**
     * sets the whether to print cell notes
     *
     * @param printNotes whether to print Notes.
     */
    var printNotes: Boolean
        get() = printerSettings!!.printNotes
        set(b) {
            printerSettings!!.printNotes = b
        }


    /**
     * get the print resolution
     *
     * @return the printer resolution in DPI
     */
    val resolution: Short
        get() = printerSettings!!.resolution

    /**
     * get the scale of the printer output in whole percentages
     *
     *
     * ie: 25 = 25%
     *
     * @return the scale of printer output
     */
    val scale: Short
        get() = printerSettings!!.scale

    /**
     * use custom start page for auto numbering
     *
     * @return Returns whether to use a custom start page
     */
    /**
     * @param usePage whether to use custom Page to start printing from.
     */
    var usePage: Boolean
        get() = printerSettings!!.usePage
        set(usePage) {
            printerSettings!!.usePage = usePage
        }

    /**
     * get the vertical print resolution
     *
     * @return the vertical printer resolution in DPI
     */
    /**
     * @param verticalResolution The vertical Resolution in DPI
     */
    var verticalResolution: Short
        get() = printerSettings!!.verticalResolution
        set(verticalResolution) {
            printerSettings!!.verticalResolution = verticalResolution
        }

    /**
     * Whether the sheet should be centered horizontally.
     */
    /**
     * Sets whether the page should be centered horizontally.
     */
    var isHCenter: Boolean
        get() = hCenter!!.isHCenter
        set(center) {
            hCenter!!.isHCenter = center
        }

    /**
     * Whether the sheet should be centered vertically.
     */
    /**
     * Sets whether the sheet should be centered vertically.
     */
    var isVCenter: Boolean
        get() = vCenter!!.isVCenter
        set(center) {
            vCenter!!.isVCenter = center
        }

    /**
     * Whether the grid lines will be printed.
     */
    val isPrintGridLines: Boolean
        get() = grid!!.isPrintGrid

    /**
     * Whether the row and column headers will be printed.
     */
    /**
     * Sets whether to print the row and column headers.
     */
    var isPrintRowColHeaders: Boolean
        get() = headers!!.isPrintHeaders
        set(print) {
            headers!!.isPrintHeaders = print
        }

    /**
     * Gets whether the sheet will be printed fit to some number of pages.
     */
    /**
     * Sets whether the sheet will be printed fit to some number of pages.
     */
    var isFitToPage: Boolean
        get() = wsBool!!.isFitToPage
        set(value) {
            wsBool!!.isFitToPage = value
        }

    /**
     * get the draft quality setting
     *
     * @return draft quality
     */
    /**
     * set for draft quality output
     *
     * @param whether to use draft quality
     */
    var draft: Boolean
        get() = printerSettings!!.draft
        set(b) {
            printerSettings!!.draft = b
        }

    /**
     * get the number of pages to fit the printout to height
     *
     * @return fit to height
     */
    val fitHeight: Short
        get() = printerSettings!!.fitHeight

    /**
     * get the number of pages to fit the printout to width
     *
     * @return fit to width
     */
    val fitWidth: Short
        get() = printerSettings!!.fitWidth

    /**
     * Gets the range specifying the titles printed on each page.
     */
    /**
     * Sets the range specifying the titles printed on each page.
     * The reference for the row(s) to repeat e.g. $1:$1 for row 1
     * For Columns, type the reference to the column or columns that
     * you want to set as a title e.g. $A:$B for columns A and B
     */
    // note:  MUST be in $ROW:$ROW or $COL:$COL format, for both
    // can be $R:$R, $C:$C for both
    //do before setNewScope as it blows out itab
    // This shouldn't be possible.
    // pre-process range to ensure in proper format, ensure all absolute ($) refs +
    // handle wholerow-wholecol refs + complex ranges
    // TODO: Do what?? remove??
    // concatenate terms into one ptgmemfunc-style expression
    // varies by column
    // varies by row
    var titles: String?
        get() {
            val range = sheet.getName("Built-in: PRINT_TITLES") ?: return null
            return range.expressionString
        }
        set(range) {
            var range = range
            var name = sheet.getName("Built-in: PRINT_TITLES")
            if (name == null)
                try {
                    name = Name(sheet.workBook, "Print_Titles")
                    name.setBuiltIn(0x07.toByte())
                    name.setNewScope(sheet.sheetNum + 1)
                } catch (e: WorkSheetNotFoundException) {
                    throw Error("sheet not found re-scoping name")
                }

            if (range == null) return
            val ranges = range.split(",".toRegex()).dropLastWhile({ it.isEmpty() }).toTypedArray()
            range = ""
            for (i in ranges.indices) {
                if (i > 0)
                    range += ","
                var r = ""
                val rc = ExcelTools.getRangeCoords(ranges[i])
                if (rc[0] == rc[2])
                    r = "$" + ExcelTools.getAlphaVal(rc[1]) + ":$" + ExcelTools.getAlphaVal(rc[3])
                if (rc[1] == rc[3]) {
                    r = "$" + rc[0] + ":$" + rc[2]
                }
                range += sheet.sheetName + "!" + r
            }
            name.location = range
        }

    init {

        val iter = sheet.printRecs.iterator()
        while (iter.hasNext()) {
            val record = iter.next() as BiffRec

            if (record is Setup)
                printerSettings = record
            else if (record is HCenter)
                hCenter = record
            else if (record is VCenter)
                vCenter = record
            else if (record is LeftMargin)
                leftMargin = record        // missing in default set of records
            else if (record is RightMargin)
                rightMargin = record        // missing in default set of records
            else if (record is TopMargin)
                topMargin = record            // missing in default set of records
            else if (record is BottomMargin)
                bottomMargin = record    // missing in default set of records
            else if (record is PrintGrid)
                grid = record
            else if (record is PrintRowCol)
                headers = record
            else if (record is WsBool)
                wsBool = record
        }
        // Actually, the below comment is incorrect: do NOT do this unconditionally
        //printerSettings.setNoPrintData(false); // there IS printer setup data
    }

    /**
     * Gets the left print margin.
     */
    fun getLeftMargin(): Double {
        if (leftMargin == null) {
            leftMargin = LeftMargin()
            this.sheet.addMarginRecord(leftMargin!!)
        }
        return leftMargin!!.margin
    }

    /**
     * Gets the right print margin.
     */
    fun getRightMargin(): Double {
        if (rightMargin == null) {
            rightMargin = RightMargin()
            this.sheet.addMarginRecord(rightMargin!!)
        }
        return rightMargin!!.margin
    }

    /**
     * Gets the top print margin.
     */
    fun getTopMargin(): Double {
        if (topMargin == null) {
            topMargin = TopMargin()
            this.sheet.addMarginRecord(topMargin!!)
        }
        return topMargin!!.margin
    }

    /**
     * Gets the bottom print margin.
     */
    fun getBottomMargin(): Double {
        if (bottomMargin == null) {
            bottomMargin = BottomMargin()
            this.sheet.addMarginRecord(bottomMargin!!)
        }
        return bottomMargin!!.margin
    }

    /**
     * Set the output to print onto this number of pages high
     *
     *
     * ie: setFitHeight(10) will stretch the print out to fit 10
     * pages high
     *
     * @param number of pages to fit to height
     */
    fun setFitHeight(numpages: Int) {
        printerSettings!!.fitHeight = numpages.toShort()
    }

    /**
     * Set the output to print onto this number of pages wide
     *
     *
     * ie: setFitWidth(10) will stretch the print out to fit 10
     * pages wide
     *
     * @param number of pages to fit to width
     */
    fun setFitWidth(numpages: Int) {
        printerSettings!!.fitWidth = numpages.toShort()
    }

    /**
     * Sets the sheet's left print margin.
     */
    fun setLeftMargin(value: Double) {
        if (leftMargin == null) {
            leftMargin = LeftMargin()
            this.sheet.addMarginRecord(leftMargin!!)
        }
        leftMargin!!.margin = value
    }

    /**
     * Sets the sheet's right print margin.
     */
    fun setRightMargin(value: Double) {
        if (rightMargin == null) {
            rightMargin = RightMargin()
            this.sheet.addMarginRecord(rightMargin!!)
        }
        rightMargin!!.margin = value
    }

    /**
     * Sets the sheet's top print margin.
     */
    fun setTopMargin(value: Double) {
        if (topMargin == null) {
            topMargin = TopMargin()
            this.sheet.addMarginRecord(topMargin!!)
        }
        topMargin!!.margin = value
    }

    /**
     * Sets the sheet's bottom print margin.
     */
    fun setBottomMargin(value: Double) {
        if (bottomMargin == null) {
            bottomMargin = BottomMargin()
            this.sheet.addMarginRecord(bottomMargin!!)
        }
        bottomMargin!!.margin = value
    }

    /**
     * Set the output printer resolution
     *
     * @param resolution The resolution to set in DPI.
     */
    fun setResolution(r: Int) {
        printerSettings!!.resolution = r.toShort()
    }

    /**
     * scale the printer output in whole percentages
     *
     *
     * ie: 25 = 25%
     *
     * @param scale The scale to set.
     */
    fun setScale(scale: Int) {
        printerSettings!!.scale = scale.toShort()
    }

    /**
     * @param copies The number of copies to print
     */
    fun setCopies(copies: Int) {
        printerSettings!!.copies = copies.toShort()
    }

    /**
     * sets the paper size based on the paper size table
     *
     * @param the paper size index
     */
    fun setPaperSize(p: Int) {
        printerSettings!!.paperSize = p.toShort()
    }

    /**
     * Sets whether to print the grid lines.
     */
    fun setPrintGrid(print: Boolean) {
        grid!!.isPrintGrid = print
    }

    companion object {


        //	 paper size ints
        val PAPER_SIZE_UNDEFINED = 0
        val PAPER_SIZE_LETTER_8_5x11 = 1 // Letter 81/2" x 11"
        val PAPER_SIZE_LETTER_SMALL = 2 // Letter small 81/2" x 11"
        val PAPER_SIZE_TABLOID_11x17 = 3 // Tabloid 11" x 17"
        val PAPER_SIZE_LEDGER_17x11 = 4 // Ledger 17" x 11"
        val PAPER_SIZE_LEGAL_8_5x14 = 5 // Legal 81/2" x 14"
        val PAPER_SIZE_STATEMENT_5_5x8_5 = 6 // Statement 51/2" x 81/2"
        val PAPER_SIZE_LETTER_EXTRA_9_5Ax12 = 50 // Letter Extra 91/2" x 12
        val PAPER_SIZE_LEGAL_EXTRA_9_5Ax15 = 51 // Legal Extra 91/2" x 15"
        val PAPER_SIZE_TABLOID_EXTRA_1111_16Ax18 = 52 // Tabloid Extra 1111/16" x 18"
        val PAPER_SIZE_A4_EXTRA_235MM_X_322MM = 53 // A4 Extra 235mm x 322mm
        val PAPER_SIZE_LETTER_TRANSVERSE_8_5Ax11 = 54 // Letter Transverse 81/2" x 11"
        val PAPER_SIZE_EXECUTIVE_7_QUARTER_X_10_5 = 7 // Executive 71/4" x 101/2"
        val PAPER_SIZE_TRANSVERSE_210MM_X_297MM = 55 // A4 Transverse 210mm x 297mm
        val PAPER_SIZE_A3_297MM_X_420MM = 8 // 8 A3 297mm x 420mm
        val PAPER_SIZE_LETTER_EXTRA_TRANSV_9_5_X_12 = 56 // 56 Letter Extra Transv. 91/2" x 12"
        val PAPER_SIZE_A4_210MM_X_297MM = 9 // A4 210mm x 297mm
        val PAPER_SIZE_SUPER_A_A4_227MM_X_356MM = 57 // Super A/A4 227mm x 356mm
        val PAPER_SIZE_A4_SMALL_210MM_X_297MM = 10 // A4 small 210mm x 297mm
        val PAPER_SIZE_SUPER_B_A3_305MM_X_487MM = 58 // Super B/A3 305mm x 487mm
        val PAPER_SIZE_A5_148MM_X_210MM = 11 // A5 148mm x 210mm
        val PAPER_SIZE_LETTER_PLUS = 59 // Letter Plus
        val PAPER_SIZE_2_X_1211_16 = 81 // 2" x 1211/16"
        val PAPER_SIZE_B4_JIS_257MM_X_364MM = 12 // B4 (JIS) 257mm x 364mm
        val PAPER_SIZE_A4_PLUS_210MM_X_330MM = 60 // A4 Plus 210mm x 330mm
        val PAPER_SIZE_B5_JIS_182MM_X_257MM = 13 // B5 (JIS) 182mm x 257mm
        val PAPER_SIZE_A5_TRANSVERSE_148MM_X_210MM = 61 // A5 Transverse 148mm x 210mm
        val PAPER_SIZE_FOLIO_8_5_X_13 = 14 // Folio 81/2" x 13"
        val PAPER_SIZE_B5_JIS_TRANSVERSE_182MM_X_257MM = 62 // B5 (JIS) Transverse 182mm x 257mm
        val PAPER_SIZE_QUATRO_215MM_X_275MM = 15 // Quarto 215mm x 275mm
        val PAPER_SIZE_A3_EXTRA_322MM_X_445MM = 63 // A3 Extra 322mm x 445mm
        val PAPER_SIZE_10Ax14_10_X_14 = 16 // 10x14 10" x 14"
        val PAPER_SIZE_A5_EXTRA_174MM_X_235 = 64 // A5 Extra 174mm x 235mm
        val PAPER_SIZE_11Ax17_11_X_17 = 17 // 11x17 11" x 17"
        val PAPER_SIZE_B5_ISO_EXTRA_201MM_X_276MM = 65 // B5 (ISO) Extra 201mm x 276mm
        val PAPER_SIZE_NOTE_8_5_X_11 = 18 // Note 81/2" x 11"
        val PAPER_SIZE_A2_420MM_X_594MM = 66 // A2 420mm x 594mm
        val PAPER_SIZE_ENVELOPE_9_3_78_X_8_78 = 19 // Envelope #9 37/8" x 87/8"
        val PAPER_SIZE_A3_TRANSVERSE_297MM_X_420MM = 67 // A3 Transverse 297mm x 420mm
        val PAPER_SIZE_ENVELOPE_10_4_18_X_9_5 = 20 // Envelope #10 41/8" x 91/2"
        val PAPER_SIZE_EXTRA_TRANSVERSE_322MM_X_445MM = 68 // A3 Extra Transverse 322mm x 445mm
        val PAPER_SIZE_ENVELOPE_11_4_5_X_10_38 = 21 // Envelope #11 41/2" x 103/8"
        val PAPER_SIZE_DBL_JAP_POSTCARD_200MM_X_148MM = 69 // Dbl. Japanese Postcard 200mm x 148mm
        val PAPER_SIZE_ENVELOPE_12_4_34_X_11 = 22 // Envelope #12 43/4" x 11"
        val PAPER_SIZE_A6_105MM_X_148MM = 70 // A6 105mm x 148mm
        val PAPER_SIZE_ENVELOPE_14_5_X_11_5 = 23 // Envelope #14 5" x 111/2"
        val PAPER_SIZE_C_17_X_22_72 = 24 // C 17" x 22" 72
        val PAPER_SIZE_D_22_X_34_73 = 25 // D 22" x 34" 73
        val PAPER_SIZE_E_34_X_44_74 = 26 // E 34" x 44" 74
        val PAPER_SIZE_DL_ENVELOPE_110MM_X_110MM_X_220MM = 27 // Envelope DL 110mm x 220mm
        val PAPER_SIZE_LETTER_ROTATED_11_X_8_5 = 75 // Letter Rotated 11" x 81/2"
        val PAPER_SIZE_ENVELOPE_C5_162MM_X_229M = 28 // Envelope C5 162mm x 229mm
        val PAPER_SIZE_A3_ROTATED_420MM_X_297MM = 76 // A3 Rotated 420mm x 297mm
        val PAPER_SIZE_ENVELOPE_C3_324MM_X_458MM = 29 // Envelope C3 324mm x 458mm
        val PAPER_SIZE_A4_ROTATED_297MM_X_210MM = 77 // A4 Rotated 297mm x 210mm
        val PAPER_SIZE_ENVELOPE_C4_229MM_X_324MM = 30 // Envelope C4 229mm x 324mm
        val PAPER_SIZE_A5_ROTATED_210MM_X_148MM = 78 // A5 Rotated 210mm x 148mm
        val PAPER_SIZE_ENVELOPE_C6_115MM_X_162MM = 31 // Envelope C6 114mm x 162mm
        val PAPER_SIZE_ENVELOPE_C6_C5_114MM_X_229MM = 32 // Envelope C6/C5 114mm x 229mm
        val PAPER_SIZE_B4_ISO_250MM_X_353MM = 33 // B4 (ISO) 250mm x 353mm
        val PAPER_SIZE_B5_ISO_176MM_X_250MM = 34 // B5 (ISO) 176mm x 250mm
        val PAPER_SIZE_DBL_JAP_POSTCARD_ROT_148MM_X_200MM = 82 // Dbl. Jap. Postcard Rot. 148mm x 200mm
        val PAPER_SIZE_B6_ISO_125MM_X_176MM = 35 // B6 (ISO) 125mm x 176mm
        val PAPER_SIZE_ENVELOPE_ITALY_10MM_X_230MM = 36 // Envelope Italy 110mm x 230mm 84
        val PAPER_SIZE_ENVELOPE_MONARCH_3_7_8_X_7_5 = 37 // Envelope Monarch 37/8" x 71/2" 85
        val PAPER_SIZE_6_3_4_ENVELOPE_3_5_8_X_6_5 = 38 // 63/4 Envelope 35/8" x 61/2" 86
        val PAPER_SIZE_US_STANDARD_FANFOLD_147_8_X_11 = 39 // US Standard Fanfold 147/8" x 11" 87
        val PAPER_SIZE_GERMAN_STD_FANFOLD_8_5_X_12 = 40 // German Std. Fanfold 81/2" x 12"
        val PAPER_SIZE_GERMAN_LEGAL_FANFOLD_8_5_X_13 = 41 // German Legal Fanfold 81/2" x 13"
        // public static final int PAPER_SIZE_B4_ISO_250MM_X_353MM = 42 ; // B4 (ISO) 250mm x 353mm
        val PAPER_SIZE_JAP_POSTCARD_100M_X_148MM = 43 // Japanese Postcard 100mm x 148mm
        val PAPER_SIZE_9_X_11 = 44 // 9x11 9" x 11"
        val PAPER_SIZE_10_X_11 = 45 // 10x11 10" x 11"
        val PAPER_SIZE_15_X_11 = 46 // 15x11 15" x 11"
        val PAPER_SIZE_ENVELOPE_INVITE_220MM_X_220MM = 47 // Envelope Invite 220mm x 220mm
        val PAPER_SIZE_B4_JIS_ROTATED_364MM_X_257MM = 79 // B4 (JIS) Rotated 364mm x 257mm
        val PAPER_SIZE_B5_JIS_ROTATED_257MMX_X_182MM = 80 // B5 (JIS) Rotated 257mm x 182mm
        val PAPER_SIZE_JAP_POSTCARD_ROT_148MM_X_100MM = 81 // Japanese Postcard Rot. 148mm x 100mm
        val PAPER_SIZE_A6_ROTATED_148MM_X_105MM = 83 // A6 Rotated 148mm x 105mm
        val PAPER_SIZE_B6_JIS_128MM_X_182MM = 88 // B6 (JIS) 128mm x 182mm
        val PAPER_SIZE_B6_JIS_ROT_182MM_X_128MM = 89 // B6 (JIS) Rotated 182mm x 128mm
        val PAPER_SIZE_12_X_11 = 90 // 12x11 12" x 11"
    }

}