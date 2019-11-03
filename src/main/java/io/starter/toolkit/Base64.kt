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
package io.starter.toolkit

/**
 * Encodes and decodes to and from Base64 notation.
 *
 *
 *
 * Change Log:
 *
 *
 *  * v2.1 - Cleaned up javadoc comments and unused variables and methods. Added
 * some convenience methods for reading and writing to and from files.
 *  * v2.0.2 - Now specifies UTF-8 encoding in places where the code fails on systems
 * with other encodings (like EBCDIC).
 *  * v2.0.1 - Fixed an error when decoding a single byte, that is, when the
 * encoded data was a single byte.
 *  * v2.0 - I got rid of methods that used booleans to set options.
 * Now everything is more consolidated and cleaner. The code now detects
 * when data that's being decoded is gzip-compressed and will decompress it
 * automatically. Generally things are cleaner. You'll probably have to
 * change some method calls that you were making to support the new
 * options format (<tt>int</tt>s that you "OR" together).
 *  * v1.5.1 - Fixed bug when decompressing and decoding to a
 * byte[] using <tt>decode( String s, boolean gzipCompressed )</tt>.
 * Added the ability to "suspend" encoding in the Output Stream so
 * you can turn on and off the encoding if you need to embed base64
 * data in an otherwise "normal" stream (like an XML file).
 *  * v1.5 - Output stream pases on flush() command but doesn't do anything itself.
 * This helps when using GZIP streams.
 * Added the ability to GZip-compress objects before encoding them.
 *  * v1.4 - Added helper methods to read/write files.
 *  * v1.3.6 - Fixed OutputStream.flush() so that 'position' is reset.
 *  * v1.3.5 - Added flag to turn on and off line breaks. Fixed bug in input stream
 * where last buffer being read, if not completely full, was not returned.
 *  * v1.3.4 - Fixed when "improperly padded stream" error was thrown at the wrong time.
 *  * v1.3.3 - Fixed I/O streams which were totally messed up.
 *
 *
 *
 *
 * I am placing this code in the Public Domain. Do with it as you will.
 * This software comes with no guarantees or warranties but with
 * plenty of well-wishing instead!
 * Please visit [http://iharder.net/base64](http://iharder.net/base64)
 * periodically to check for updates or to contribute improvements.
 *
 *
 * @author Robert Harder
 * @author rob@iharder.net
 * @version 2.1
 */
object Base64 {

    /* ********  P U B L I C   F I E L D S  ******** */


    /**
     * No options specified. Value is zero.
     */
    val NO_OPTIONS = 0

    /**
     * Specify encoding.
     */
    val ENCODE = 1


    /**
     * Specify decoding.
     */
    val DECODE = 0


    /**
     * Specify that data should be gzip-compressed.
     */
    val GZIP = 2


    /**
     * Don't break lines when encoding (violates strict Base64 specification)
     */
    val DONT_BREAK_LINES = 8


    /* ********  P R I V A T E   F I E L D S  ******** */


    /**
     * Maximum line length (76) of Base64 output.
     */
    private val MAX_LINE_LENGTH = 76


    /**
     * The equals sign (=) as a byte.
     */
    private val EQUALS_SIGN = '='.toByte()


    /**
     * The new line character (\n) as a byte.
     */
    private val NEW_LINE = '\n'.toByte()


    /**
     * Preferred encoding.
     */
    private val PREFERRED_ENCODING = "UTF-8"


    /**
     * The 64 valid Base64 values.
     */
    private val ALPHABET: ByteArray
    private val _NATIVE_ALPHABET = byteArrayOf('A'.toByte(), 'B'.toByte(), 'C'.toByte(), 'D'.toByte(), 'E'.toByte(), 'F'.toByte(), 'G'.toByte(), 'H'.toByte(), 'I'.toByte(), 'J'.toByte(), 'K'.toByte(), 'L'.toByte(), 'M'.toByte(), 'N'.toByte(), 'O'.toByte(), 'P'.toByte(), 'Q'.toByte(), 'R'.toByte(), 'S'.toByte(), 'T'.toByte(), 'U'.toByte(), 'V'.toByte(), 'W'.toByte(), 'X'.toByte(), 'Y'.toByte(), 'Z'.toByte(), 'a'.toByte(), 'b'.toByte(), 'c'.toByte(), 'd'.toByte(), 'e'.toByte(), 'f'.toByte(), 'g'.toByte(), 'h'.toByte(), 'i'.toByte(), 'j'.toByte(), 'k'.toByte(), 'l'.toByte(), 'm'.toByte(), 'n'.toByte(), 'o'.toByte(), 'p'.toByte(), 'q'.toByte(), 'r'.toByte(), 's'.toByte(), 't'.toByte(), 'u'.toByte(), 'v'.toByte(), 'w'.toByte(), 'x'.toByte(), 'y'.toByte(), 'z'.toByte(), '0'.toByte(), '1'.toByte(), '2'.toByte(), '3'.toByte(), '4'.toByte(), '5'.toByte(), '6'.toByte(), '7'.toByte(), '8'.toByte(), '9'.toByte(), '+'.toByte(), '/'.toByte())/* May be something funny like EBCDIC */


    /**
     * Translates a Base64 value to either its 6-bit reconstruction value
     * or a negative number indicating some other meaning.
     */
    private val DECODABET = byteArrayOf(-9, -9, -9, -9, -9, -9, -9, -9, -9, // Decimal  0 -  8
            -5, -5, // Whitespace: Tab and Linefeed
            -9, -9, // Decimal 11 - 12
            -5, // Whitespace: Carriage Return
            -9, -9, -9, -9, -9, -9, -9, -9, -9, -9, -9, -9, -9, // Decimal 14 - 26
            -9, -9, -9, -9, -9, // Decimal 27 - 31
            -5, // Whitespace: Space
            -9, -9, -9, -9, -9, -9, -9, -9, -9, -9, // Decimal 33 - 42
            62, // Plus sign at decimal 43
            -9, -9, -9, // Decimal 44 - 46
            63, // Slash at decimal 47
            52, 53, 54, 55, 56, 57, 58, 59, 60, 61, // Numbers zero through nine
            -9, -9, -9, // Decimal 58 - 60
            -1, // Equals sign at decimal 61
            -9, -9, -9, // Decimal 62 - 64
            0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, // Letters 'A' through 'N'
            14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, // Letters 'O' through 'Z'
            -9, -9, -9, -9, -9, -9, // Decimal 91 - 96
            26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, // Letters 'a' through 'm'
            39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, // Letters 'n' through 'z'
            -9, -9, -9, -9                                 // Decimal 123 - 126
    )/*,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,     // Decimal 127 - 139
                    -9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,     // Decimal 140 - 152
                    -9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,     // Decimal 153 - 165
                    -9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,     // Decimal 166 - 178
                    -9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,     // Decimal 179 - 191
                    -9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,     // Decimal 192 - 204
                    -9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,     // Decimal 205 - 217
                    -9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,     // Decimal 218 - 230
                    -9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,     // Decimal 231 - 243
                    -9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9,-9         // Decimal 244 - 255 */

    // I think I end up not using the BAD_ENCODING indicator.
    //private final static byte BAD_ENCODING    = -9; // Indicates error in encoding
    private val WHITE_SPACE_ENC: Byte = -5 // Indicates white space in encoding
    private val EQUALS_SIGN_ENC: Byte = -1 // Indicates equals sign in encoding

    /** Determine which ALPHABET to use.  */
    init {
        var __bytes: ByteArray
        try {
            __bytes = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/".toByteArray(charset(PREFERRED_ENCODING))
        }   // end try
        catch (use: java.io.UnsupportedEncodingException) {
            __bytes = _NATIVE_ALPHABET // Fall back to native encoding
        }
        // end catch
        ALPHABET = __bytes
    }   // end static


    /* ********  E N C O D I N G   M E T H O D S  ******** */


    /**
     * Encodes up to the first three bytes of array <var>threeBytes</var>
     * and returns a four-byte array in Base64 notation.
     * The actual number of significant bytes in your array is
     * given by <var>numSigBytes</var>.
     * The array <var>threeBytes</var> needs only be as big as
     * <var>numSigBytes</var>.
     * Code can reuse a byte array by passing a four-byte array as <var>b4</var>.
     *
     * @param b4          A reusable byte array to reduce array instantiation
     * @param threeBytes  the array to convert
     * @param numSigBytes the number of significant bytes in your array
     * @return four byte array in Base64 notation.
     */
    private fun encode3to4(b4: ByteArray, threeBytes: ByteArray?, numSigBytes: Int): ByteArray {
        encode3to4(threeBytes, 0, numSigBytes, b4, 0)
        return b4
    }   // end encode3to4


    /**
     * Encodes up to three bytes of the array <var>source</var>
     * and writes the resulting four Base64 bytes to <var>destination</var>.
     * The source and destination arrays can be manipulated
     * anywhere along their length by specifying
     * <var>srcOffset</var> and <var>destOffset</var>.
     * This method does not check to make sure your arrays
     * are large enough to accomodate <var>srcOffset</var> + 3 for
     * the <var>source</var> array or <var>destOffset</var> + 4 for
     * the <var>destination</var> array.
     * The actual number of significant bytes in your array is
     * given by <var>numSigBytes</var>.
     *
     * @param source      the array to convert
     * @param srcOffset   the index where conversion begins
     * @param numSigBytes the number of significant bytes in your array
     * @param destination the array to hold the conversion
     * @param destOffset  the index where output will be put
     * @return the <var>destination</var> array
     */
    private fun encode3to4(
            source: ByteArray?, srcOffset: Int, numSigBytes: Int,
            destination: ByteArray, destOffset: Int): ByteArray {
        //           1         2         3
        // 01234567890123456789012345678901 Bit position
        // --------000000001111111122222222 Array position from threeBytes
        // --------|    ||    ||    ||    | Six bit groups to index ALPHABET
        //          >>18  >>12  >> 6  >> 0  Right shift necessary
        //                0x3f  0x3f  0x3f  Additional AND

        // Create buffer with zero-padding if there are only one or two
        // significant bytes passed in the array.
        // We have to shift left 24 in order to flush out the 1's that appear
        // when Java treats a value as negative that is cast from a byte to an int.
        val inBuff = ((if (numSigBytes > 0) (source!![srcOffset] shl 24).ushr(8) else 0)
                or (if (numSigBytes > 1) (source!![srcOffset + 1] shl 24).ushr(16) else 0)
                or if (numSigBytes > 2) (source!![srcOffset + 2] shl 24).ushr(24) else 0)

        when (numSigBytes) {
            3 -> {
                destination[destOffset] = ALPHABET[inBuff.ushr(18)]
                destination[destOffset + 1] = ALPHABET[inBuff.ushr(12) and 0x3f]
                destination[destOffset + 2] = ALPHABET[inBuff.ushr(6) and 0x3f]
                destination[destOffset + 3] = ALPHABET[inBuff and 0x3f]
                return destination
            }

            2 -> {
                destination[destOffset] = ALPHABET[inBuff.ushr(18)]
                destination[destOffset + 1] = ALPHABET[inBuff.ushr(12) and 0x3f]
                destination[destOffset + 2] = ALPHABET[inBuff.ushr(6) and 0x3f]
                destination[destOffset + 3] = EQUALS_SIGN
                return destination
            }

            1 -> {
                destination[destOffset] = ALPHABET[inBuff.ushr(18)]
                destination[destOffset + 1] = ALPHABET[inBuff.ushr(12) and 0x3f]
                destination[destOffset + 2] = EQUALS_SIGN
                destination[destOffset + 3] = EQUALS_SIGN
                return destination
            }

            else -> return destination
        }   // end switch
    }   // end encode3to4


    /**
     * Serializes an object and returns the Base64-encoded
     * version of that serialized object. If the object
     * cannot be serialized or there is another error,
     * the method will return <tt>null</tt>.
     *
     *
     * Valid options:<pre>
     * GZIP: gzip-compresses object before encoding it.
     * DONT_BREAK_LINES: don't break lines at 76 characters
     * *Note: Technically, this makes your encoding non-compliant.*
    </pre> *
     *
     *
     * Example: `encodeObject( myObj, Base64.GZIP )` or
     *
     *
     * Example: `encodeObject( myObj, Base64.GZIP | Base64.DONT_BREAK_LINES )`
     *
     * @param serializableObject The object to encode
     * @param options            Specified options
     * @return The Base64-encoded object
     * @see Base64.GZIP
     *
     * @see Base64.DONT_BREAK_LINES
     */
    @JvmOverloads
    fun encodeObject(serializableObject: java.io.Serializable, options: Int = NO_OPTIONS): String? {
        // Streams
        var baos: java.io.ByteArrayOutputStream? = null
        var b64os: java.io.OutputStream? = null
        var oos: java.io.ObjectOutputStream? = null
        var gzos: java.util.zip.GZIPOutputStream? = null

        // Isolate options
        val gzip = options and GZIP
        val dontBreakLines = options and DONT_BREAK_LINES

        try {
            // ObjectOutputStream -> (GZIP) -> Base64 -> ByteArrayOutputStream
            baos = java.io.ByteArrayOutputStream()
            b64os = Base64.OutputStream(baos, ENCODE or dontBreakLines)

            // GZip?
            if (gzip == GZIP) {
                gzos = java.util.zip.GZIPOutputStream(b64os)
                oos = java.io.ObjectOutputStream(gzos)
            }   // end if: gzip
            else
                oos = java.io.ObjectOutputStream(b64os)

            oos.writeObject(serializableObject)
        }   // end try
        catch (e: java.io.IOException) {
            e.printStackTrace()
            return null
        }   // end catch
        finally {
            try {
                oos!!.close()
            } catch (e: Exception) {
            }

            try {
                gzos!!.close()
            } catch (e: Exception) {
            }

            try {
                b64os!!.close()
            } catch (e: Exception) {
            }

            try {
                baos!!.close()
            } catch (e: Exception) {
            }

        }   // end finally

        // Return value according to relevant encoding.
        try {
            return String(baos!!.toByteArray(), PREFERRED_ENCODING)
        }   // end try
        catch (uue: java.io.UnsupportedEncodingException) {
            return String(baos!!.toByteArray())
        }
        // end catch

    }   // end encode


    /**
     * Encodes a byte array into Base64 notation.
     *
     *
     * Valid options:<pre>
     * GZIP: gzip-compresses object before encoding it.
     * DONT_BREAK_LINES: don't break lines at 76 characters
     * *Note: Technically, this makes your encoding non-compliant.*
    </pre> *
     *
     *
     * Example: `encodeBytes( myData, Base64.GZIP )` or
     *
     *
     * Example: `encodeBytes( myData, Base64.GZIP | Base64.DONT_BREAK_LINES )`
     *
     * @param source  The data to convert
     * @param options Specified options
     * @see Base64.GZIP
     *
     * @see Base64.DONT_BREAK_LINES
     */
    fun encodeBytes(source: ByteArray, options: Int): String? {
        return encodeBytes(source, 0, source.size, options)
    }   // end encodeBytes


    /**
     * Encodes a byte array into Base64 notation.
     *
     *
     * Valid options:<pre>
     * GZIP: gzip-compresses object before encoding it.
     * DONT_BREAK_LINES: don't break lines at 76 characters
     * *Note: Technically, this makes your encoding non-compliant.*
    </pre> *
     *
     *
     * Example: `encodeBytes( myData, Base64.GZIP )` or
     *
     *
     * Example: `encodeBytes( myData, Base64.GZIP | Base64.DONT_BREAK_LINES )`
     *
     * @param source  The data to convert
     * @param off     Offset in array where conversion should begin
     * @param len     Length of data to convert
     * @param options Specified options
     * @see Base64.GZIP
     *
     * @see Base64.DONT_BREAK_LINES
     */
    @JvmOverloads
    fun encodeBytes(source: ByteArray, off: Int = 0, len: Int = source.size, options: Int = NO_OPTIONS): String? {
        // Isolate options
        val dontBreakLines = options and DONT_BREAK_LINES
        val gzip = options and GZIP

        // Compress?
        if (gzip == GZIP) {
            var baos: java.io.ByteArrayOutputStream? = null
            var gzos: java.util.zip.GZIPOutputStream? = null
            var b64os: Base64.OutputStream? = null


            try {
                // GZip -> Base64 -> ByteArray
                baos = java.io.ByteArrayOutputStream()
                b64os = Base64.OutputStream(baos, ENCODE or dontBreakLines)
                gzos = java.util.zip.GZIPOutputStream(b64os)

                gzos.write(source, off, len)
                gzos.close()
            }   // end try
            catch (e: java.io.IOException) {
                e.printStackTrace()
                return null
            }   // end catch
            finally {
                try {
                    gzos!!.close()
                } catch (e: Exception) {
                }

                try {
                    b64os!!.close()
                } catch (e: Exception) {
                }

                try {
                    baos!!.close()
                } catch (e: Exception) {
                }

            }   // end finally

            // Return value according to relevant encoding.
            try {
                return String(baos!!.toByteArray(), PREFERRED_ENCODING)
            }   // end try
            catch (uue: java.io.UnsupportedEncodingException) {
                return String(baos!!.toByteArray())
            }
            // end catch
        }   // end if: compress
        else {
            // Convert option to boolean in way that code likes it.
            val breakLines = dontBreakLines == 0

            val len43 = len * 4 / 3
            val outBuff = ByteArray(len43                      // Main 4:3

                    + (if (len % 3 > 0) 4 else 0)      // Account for padding

                    + if (breakLines) len43 / MAX_LINE_LENGTH else 0) // New lines
            var d = 0
            var e = 0
            val len2 = len - 2
            var lineLength = 0
            while (d < len2) {
                encode3to4(source, d + off, 3, outBuff, e)

                lineLength += 4
                if (breakLines && lineLength == MAX_LINE_LENGTH) {
                    outBuff[e + 4] = NEW_LINE
                    e++
                    lineLength = 0
                }   // end if: end of line
                d += 3
                e += 4
            }   // en dfor: each piece of array

            if (d < len) {
                encode3to4(source, d + off, len - d, outBuff, e)
                e += 4
            }   // end if: some padding needed


            // Return value according to relevant encoding.
            try {
                return String(outBuff, 0, e, PREFERRED_ENCODING)
            }   // end try
            catch (uue: java.io.UnsupportedEncodingException) {
                return String(outBuff, 0, e)
            }
            // end catch

        }// Else, don't compress. Better not to use streams at all then.
        // end else: don't compress

    }   // end encodeBytes


    /* ********  D E C O D I N G   M E T H O D S  ******** */


    /**
     * Decodes four bytes from array <var>source</var>
     * and writes the resulting bytes (up to three of them)
     * to <var>destination</var>.
     * The source and destination arrays can be manipulated
     * anywhere along their length by specifying
     * <var>srcOffset</var> and <var>destOffset</var>.
     * This method does not check to make sure your arrays
     * are large enough to accomodate <var>srcOffset</var> + 4 for
     * the <var>source</var> array or <var>destOffset</var> + 3 for
     * the <var>destination</var> array.
     * This method returns the actual number of bytes that
     * were converted from the Base64 encoding.
     *
     * @param source      the array to convert
     * @param srcOffset   the index where conversion begins
     * @param destination the array to hold the conversion
     * @param destOffset  the index where output will be put
     * @return the number of decoded bytes converted
     */
    private fun decode4to3(source: ByteArray, srcOffset: Int, destination: ByteArray, destOffset: Int): Int {
        // Example: Dk==
        if (source[srcOffset + 2] == EQUALS_SIGN) {
            // Two ways to do the same thing. Don't know which way I like best.
            //int outBuff =   ( ( DECODABET[ source[ srcOffset    ] ] << 24 ) >>>  6 )
            //              | ( ( DECODABET[ source[ srcOffset + 1] ] << 24 ) >>> 12 );
            val outBuff = DECODABET[source[srcOffset]] and 0xFF shl 18 or (DECODABET[source[srcOffset + 1]] and 0xFF shl 12)

            destination[destOffset] = outBuff.ushr(16).toByte()
            return 1
        } else if (source[srcOffset + 3] == EQUALS_SIGN) {
            // Two ways to do the same thing. Don't know which way I like best.
            //int outBuff =   ( ( DECODABET[ source[ srcOffset     ] ] << 24 ) >>>  6 )
            //              | ( ( DECODABET[ source[ srcOffset + 1 ] ] << 24 ) >>> 12 )
            //              | ( ( DECODABET[ source[ srcOffset + 2 ] ] << 24 ) >>> 18 );
            val outBuff = (DECODABET[source[srcOffset]] and 0xFF shl 18
                    or (DECODABET[source[srcOffset + 1]] and 0xFF shl 12)
                    or (DECODABET[source[srcOffset + 2]] and 0xFF shl 6))

            destination[destOffset] = outBuff.ushr(16).toByte()
            destination[destOffset + 1] = outBuff.ushr(8).toByte()
            return 2
        } else {
            try {
                // Two ways to do the same thing. Don't know which way I like best.
                //int outBuff =   ( ( DECODABET[ source[ srcOffset     ] ] << 24 ) >>>  6 )
                //              | ( ( DECODABET[ source[ srcOffset + 1 ] ] << 24 ) >>> 12 )
                //              | ( ( DECODABET[ source[ srcOffset + 2 ] ] << 24 ) >>> 18 )
                //              | ( ( DECODABET[ source[ srcOffset + 3 ] ] << 24 ) >>> 24 );
                val outBuff = (DECODABET[source[srcOffset]] and 0xFF shl 18
                        or (DECODABET[source[srcOffset + 1]] and 0xFF shl 12)
                        or (DECODABET[source[srcOffset + 2]] and 0xFF shl 6)
                        or (DECODABET[source[srcOffset + 3]] and 0xFF))


                destination[destOffset] = (outBuff shr 16).toByte()
                destination[destOffset + 1] = (outBuff shr 8).toByte()
                destination[destOffset + 2] = outBuff.toByte()

                return 3
            } catch (e: Exception) {
                io.starter.toolkit.Logger.log("" + source[srcOffset] + ": " + DECODABET[source[srcOffset]])
                io.starter.toolkit.Logger.log("" + source[srcOffset + 1] + ": " + DECODABET[source[srcOffset + 1]])
                io.starter.toolkit.Logger.log("" + source[srcOffset + 2] + ": " + DECODABET[source[srcOffset + 2]])
                io.starter.toolkit.Logger.log("" + source[srcOffset + 3] + ": " + DECODABET[source[srcOffset + 3]])
                return -1
            }
            //e nd catch
        }// Example: DkLE
        // Example: DkL=
    }   // end decodeToBytes


    /**
     * Very low-level access to decoding ASCII characters in
     * the form of a byte array. Does not support automatically
     * gunzipping or any other "fancy" features.
     *
     * @param source The Base64 encoded data
     * @param off    The offset of where to begin decoding
     * @param len    The length of characters to decode
     * @return decoded data
     */
    fun decode(source: ByteArray, off: Int, len: Int): ByteArray? {
        val len34 = len * 3 / 4
        val outBuff = ByteArray(len34) // Upper limit on size of output
        var outBuffPosn = 0

        val b4 = ByteArray(4)
        var b4Posn = 0
        var i = 0
        var sbiCrop: Byte = 0
        var sbiDecode: Byte = 0
        i = off
        while (i < off + len) {
            sbiCrop = (source[i] and 0x7f).toByte() // Only the low seven bits
            sbiDecode = DECODABET[sbiCrop]

            if (sbiDecode >= WHITE_SPACE_ENC)
            // White space, Equals sign or better
            {
                if (sbiDecode >= EQUALS_SIGN_ENC) {
                    b4[b4Posn++] = sbiCrop
                    if (b4Posn > 3) {
                        outBuffPosn += decode4to3(b4, 0, outBuff, outBuffPosn)
                        b4Posn = 0

                        // If that was the equals sign, break out of 'for' loop
                        if (sbiCrop == EQUALS_SIGN)
                            break
                    }   // end if: quartet built

                }   // end if: equals sign or better

            }   // end if: white space, equals sign or better
            else {
                System.err.println("Bad Base64 input character at " + i + ": " + source[i] + "(decimal)")
                return null
            }   // end else:
            i++
        }   // each input character

        val out = ByteArray(outBuffPosn)
        System.arraycopy(outBuff, 0, out, 0, outBuffPosn)
        return out
    }   // end decode


    /**
     * Decodes data from Base64 notation, automatically
     * detecting gzip-compressed data and decompressing it.
     *
     * @param s the string to decode
     * @return the decoded data
     */
    fun decode(s: String): ByteArray? {
        var bytes: ByteArray?
        try {
            bytes = s.toByteArray(charset(PREFERRED_ENCODING))
        }   // end try
        catch (uee: java.io.UnsupportedEncodingException) {
            bytes = s.toByteArray()
        }
        // end catch
        //</change>

        // Decode
        bytes = decode(bytes, 0, bytes!!.size)


        // Check to see if it's gzip-compressed
        // GZIP Magic Two-Byte Number: 0x8b1f (35615)
        if (bytes != null && bytes.size >= 4) {

            val head = bytes[0].toInt() and 0xff or (bytes[1] shl 8 and 0xff00)
            if (java.util.zip.GZIPInputStream.GZIP_MAGIC == head) {
                var bais: java.io.ByteArrayInputStream? = null
                var gzis: java.util.zip.GZIPInputStream? = null
                var baos: java.io.ByteArrayOutputStream? = null
                val buffer = ByteArray(2048)
                var length = 0

                try {
                    baos = java.io.ByteArrayOutputStream()
                    bais = java.io.ByteArrayInputStream(bytes)
                    gzis = java.util.zip.GZIPInputStream(bais)

                    while ((length = gzis.read(buffer)) >= 0) {
                        baos.write(buffer, 0, length)
                    }   // end while: reading input

                    // No error? Get new bytes.
                    bytes = baos.toByteArray()

                }   // end try
                catch (e: java.io.IOException) {
                    // Just return originally-decoded bytes
                }   // end catch
                finally {
                    try {
                        baos!!.close()
                    } catch (e: Exception) {
                    }

                    try {
                        gzis!!.close()
                    } catch (e: Exception) {
                    }

                    try {
                        bais!!.close()
                    } catch (e: Exception) {
                    }

                }   // end finally

            }   // end if: gzipped
        }   // end if: bytes.length >= 2

        return bytes
    }   // end decode


    /**
     * Attempts to decode Base64 data and deserialize a Java
     * Object within. Returns <tt>null</tt> if there was an error.
     *
     * @param encodedObject The Base64 data to decode
     * @return The decoded and deserialized object
     */
    fun decodeToObject(encodedObject: String): Any? {
        // Decode and gunzip if necessary
        val objBytes = decode(encodedObject)

        var bais: java.io.ByteArrayInputStream? = null
        var ois: java.io.ObjectInputStream? = null
        var obj: Any? = null

        try {
            bais = java.io.ByteArrayInputStream(objBytes!!)
            ois = java.io.ObjectInputStream(bais)

            obj = ois.readObject()
        }   // end try
        catch (e: java.io.IOException) {
            e.printStackTrace()
            obj = null
        }   // end catch
        catch (e: java.lang.ClassNotFoundException) {
            e.printStackTrace()
            obj = null
        }   // end catch
        finally {
            try {
                bais!!.close()
            } catch (e: Exception) {
            }

            try {
                ois!!.close()
            } catch (e: Exception) {
            }

        }   // end finally

        return obj
    }   // end decodeObject


    /**
     * Convenience method for encoding data to a file.
     *
     * @param dataToEncode byte array of data to encode in base64 form
     * @param filename     Filename for saving encoded data
     * @return <tt>true</tt> if successful, <tt>false</tt> otherwise
     */
    fun encodeToFile(dataToEncode: ByteArray, filename: String): Boolean {
        var success = false
        var bos: Base64.OutputStream? = null
        try {
            bos = Base64.OutputStream(
                    java.io.FileOutputStream(filename), Base64.ENCODE)
            bos.write(dataToEncode)
            success = true
        }   // end try
        catch (e: java.io.IOException) {

            success = false
        }   // end catch: IOException
        finally {
            try {
                bos!!.close()
            } catch (e: Exception) {
            }

        }   // end finally

        return success
    }   // end encodeToFile


    /**
     * Convenience method for decoding data to a file.
     *
     * @param dataToDecode Base64-encoded data as a string
     * @param filename     Filename for saving decoded data
     * @return <tt>true</tt> if successful, <tt>false</tt> otherwise
     */
    fun decodeToFile(dataToDecode: String, filename: String): Boolean {
        var success = false
        var bos: Base64.OutputStream? = null
        try {
            bos = Base64.OutputStream(
                    java.io.FileOutputStream(filename), Base64.DECODE)
            bos.write(dataToDecode.toByteArray(charset(PREFERRED_ENCODING)))
            success = true
        }   // end try
        catch (e: java.io.IOException) {
            success = false
        }   // end catch: IOException
        finally {
            try {
                bos!!.close()
            } catch (e: Exception) {
            }

        }   // end finally

        return success
    }   // end decodeToFile


    /**
     * Convenience method for reading a base64-encoded
     * file and decoding it.
     *
     * @param filename Filename for reading encoded data
     * @return decoded byte array or null if unsuccessful
     */
    fun decodeFromFile(filename: String): ByteArray? {
        var decodedData: ByteArray? = null
        var bis: Base64.InputStream? = null
        try {
            // Set up some useful variables
            val file = java.io.File(filename)
            var buffer: ByteArray? = null
            var length = 0
            var numBytes = 0

            // Check for size of file
            if (file.length() > Integer.MAX_VALUE) {
                System.err.println("File is too big for this convenience method (" + file.length() + " bytes).")
                return null
            }   // end if: file too big for int index
            buffer = ByteArray(file.length().toInt())

            // Open a stream
            bis = Base64.InputStream(
                    java.io.BufferedInputStream(
                            java.io.FileInputStream(file)), Base64.DECODE)

            // Read until done
            while ((numBytes = bis.read(buffer, length, 4096)) >= 0)
                length += numBytes

            // Save in a variable to return
            decodedData = ByteArray(length)
            System.arraycopy(buffer, 0, decodedData, 0, length)

        }   // end try
        catch (e: java.io.IOException) {
            System.err.println("Error decoding from file $filename")
        }   // end catch: IOException
        finally {
            try {
                bis!!.close()
            } catch (e: Exception) {
            }

        }   // end finally

        return decodedData
    }   // end decodeFromFile


    /**
     * Convenience method for reading a binary file
     * and base64-encoding it.
     *
     * @param filename Filename for reading binary data
     * @return base64-encoded string or null if unsuccessful
     */
    fun encodeFromFile(filename: String): String? {
        var encodedData: String? = null
        var bis: Base64.InputStream? = null
        try {
            // Set up some useful variables
            val file = java.io.File(filename)
            val buffer = ByteArray((file.length() * 1.4).toInt())
            var length = 0
            var numBytes = 0

            // Open a stream
            bis = Base64.InputStream(
                    java.io.BufferedInputStream(
                            java.io.FileInputStream(file)), Base64.ENCODE)

            // Read until done
            while ((numBytes = bis.read(buffer, length, 4096)) >= 0)
                length += numBytes

            // Save in a variable to return
            encodedData = String(buffer, 0, length, Base64.PREFERRED_ENCODING)

        }   // end try
        catch (e: java.io.IOException) {
            System.err.println("Error encoding from file $filename")
        }   // end catch: IOException
        finally {
            try {
                bis!!.close()
            } catch (e: Exception) {
            }

        }   // end finally

        return encodedData
    }   // end encodeFromFile


    /* ********  I N N E R   C L A S S   I N P U T S T R E A M  ******** */


    /**
     * A [Base64.InputStream] will read data from another
     * <tt>java.io.InputStream</tt>, given in the constructor,
     * and encode/decode to/from Base64 notation on the fly.
     *
     * @see Base64
     */
    class InputStream
    /**
     * Constructs a [Base64.InputStream] in
     * either ENCODE or DECODE mode.
     *
     *
     * Valid options:<pre>
     * ENCODE or DECODE: Encode or Decode as data is read.
     * DONT_BREAK_LINES: don't break lines at 76 characters
     * (only meaningful when encoding)
     * *Note: Technically, this makes your encoding non-compliant.*
    </pre> *
     *
     *
     * Example: `new Base64.InputStream( in, Base64.DECODE )`
     *
     * @param in      the <tt>java.io.InputStream</tt> from which to read data.
     * @param options Specified options
     * @see Base64.ENCODE
     *
     * @see Base64.DECODE
     *
     * @see Base64.DONT_BREAK_LINES
     */
    @JvmOverloads constructor(`in`: java.io.InputStream, options: Int = DECODE) : java.io.FilterInputStream(`in`) {
        private val encode: Boolean         // Encoding or decoding
        private var position: Int = 0       // Current position in the buffer
        private val buffer: ByteArray         // Small buffer holding converted data
        private val bufferLength: Int   // Length of buffer (3 or 4)
        private var numSigBytes: Int = 0    // Number of meaningful bytes in the buffer
        private var lineLength: Int = 0
        private val breakLines: Boolean     // Break lines at less than 80 characters


        init {
            this.breakLines = options and DONT_BREAK_LINES != DONT_BREAK_LINES
            this.encode = options and ENCODE == ENCODE
            this.bufferLength = if (encode) 4 else 3
            this.buffer = ByteArray(bufferLength)
            this.position = -1
            this.lineLength = 0
        }   // end constructor

        /**
         * Reads enough of the input stream to convert
         * to/from Base64 and returns the next byte.
         *
         * @return next byte
         */
        @Throws(java.io.IOException::class)
        override fun read(): Int {
            // Do we need to get data?
            if (position < 0) {
                if (encode) {
                    val b3 = ByteArray(3)
                    var numBinaryBytes = 0
                    for (i in 0..2) {
                        try {
                            val b = `in`.read()

                            // If end of stream, b is -1.
                            if (b >= 0) {
                                b3[i] = b.toByte()
                                numBinaryBytes++
                            }   // end if: not end of stream

                        }   // end try: read
                        catch (e: java.io.IOException) {
                            // Only a problem if we got no data at all.
                            if (i == 0)
                                throw e

                        }
                        // end catch
                    }   // end for: each needed input byte

                    if (numBinaryBytes > 0) {
                        encode3to4(b3, 0, numBinaryBytes, buffer, 0)
                        position = 0
                        numSigBytes = 4
                    }   // end if: got data
                    else {
                        return -1
                    }   // end else
                }   // end if: encoding
                else {
                    val b4 = ByteArray(4)
                    var i = 0
                    i = 0
                    while (i < 4) {
                        // Read four "meaningful" bytes:
                        var b = 0
                        do {
                            b = `in`.read()
                        } while (b >= 0 && DECODABET[b and 0x7f] <= WHITE_SPACE_ENC)

                        if (b < 0)
                            break // Reads a -1 if end of stream

                        b4[i] = b.toByte()
                        i++
                    }   // end for: each needed input byte

                    if (i == 4) {
                        numSigBytes = decode4to3(b4, 0, buffer, 0)
                        position = 0
                    }   // end if: got four characters
                    else return if (i == 0) {
                        -1
                    }   // end else if: also padded correctly
                    else {
                        // Must have broken out from above.
                        throw java.io.IOException("Improperly padded Base64 input.")
                    }   // end

                }// Else decoding
                // end else: decode
            }   // end else: get data

            // Got data?
            if (position >= 0) {
                // End of relevant data?
                if (/*!encode &&*/ position >= numSigBytes)
                    return -1

                if (encode && breakLines && lineLength >= MAX_LINE_LENGTH) {
                    lineLength = 0
                    return '\n'.toInt()
                }   // end if
                else {
                    lineLength++   // This isn't important when decoding
                    // but throwing an extra "if" seems
                    // just as wasteful.

                    val b = buffer[position++].toInt()

                    if (position >= bufferLength)
                        position = -1

                    return b and 0xFF // This is how you "cast" a byte that's
                    // intended to be unsigned.
                }   // end else
            }   // end if: position >= 0
            else {
                // When JDK1.4 is more accepted, use an assertion here.
                throw java.io.IOException("Error in Base64 code reading stream.")
            }// Else error
            // end else
        }   // end read


        /**
         * Calls [.read] repeatedly until the end of stream
         * is reached or <var>len</var> bytes are read.
         * Returns number of bytes read into array or -1 if
         * end of stream is encountered.
         *
         * @param dest array to hold values
         * @param off  offset for array
         * @param len  max number of bytes to read into array
         * @return bytes read into array or -1 if end of stream is encountered.
         */
        @Throws(java.io.IOException::class)
        override fun read(dest: ByteArray, off: Int, len: Int): Int {
            var i: Int
            var b: Int
            i = 0
            while (i < len) {
                b = read()

                //if( b < 0 && i == 0 )
                //    return -1;

                if (b >= 0)
                    dest[off + i] = b.toByte()
                else return if (i == 0)
                    -1
                else
                    break // Out of 'for' loop
                i++
            }   // end for: each byte read
            return i
        }   // end read

    }
    /**
     * Constructs a [Base64.InputStream] in DECODE mode.
     *
     * @param in the <tt>java.io.InputStream</tt> from which to read data.
     */// end constructor
    // end inner class InputStream


    /* ********  I N N E R   C L A S S   O U T P U T S T R E A M  ******** */


    /**
     * A [Base64.OutputStream] will write data to another
     * <tt>java.io.OutputStream</tt>, given in the constructor,
     * and encode/decode to/from Base64 notation on the fly.
     *
     * @see Base64
     */
    class OutputStream
    /**
     * Constructs a [Base64.OutputStream] in
     * either ENCODE or DECODE mode.
     *
     *
     * Valid options:<pre>
     * ENCODE or DECODE: Encode or Decode as data is read.
     * DONT_BREAK_LINES: don't break lines at 76 characters
     * (only meaningful when encoding)
     * *Note: Technically, this makes your encoding non-compliant.*
    </pre> *
     *
     *
     * Example: `new Base64.OutputStream( out, Base64.ENCODE )`
     *
     * @param out     the <tt>java.io.OutputStream</tt> to which data will be written.
     * @param options Specified options.
     * @see Base64.ENCODE
     *
     * @see Base64.DECODE
     *
     * @see Base64.DONT_BREAK_LINES
     */
    @JvmOverloads constructor(out: java.io.OutputStream, options: Int = ENCODE) : java.io.FilterOutputStream(out) {
        private val encode: Boolean
        private var position: Int = 0
        private var buffer: ByteArray? = null
        private val bufferLength: Int
        private var lineLength: Int = 0
        private val breakLines: Boolean
        private val b4: ByteArray // Scratch used in a few places
        private var suspendEncoding: Boolean = false


        init {
            this.breakLines = options and DONT_BREAK_LINES != DONT_BREAK_LINES
            this.encode = options and ENCODE == ENCODE
            this.bufferLength = if (encode) 3 else 4
            this.buffer = ByteArray(bufferLength)
            this.position = 0
            this.lineLength = 0
            this.suspendEncoding = false
            this.b4 = ByteArray(4)
        }   // end constructor


        /**
         * Writes the byte to the output stream after
         * converting to/from Base64 notation.
         * When encoding, bytes are buffered three
         * at a time before the output stream actually
         * gets a write() call.
         * When decoding, bytes are buffered four
         * at a time.
         *
         * @param theByte the byte to write
         */
        @Throws(java.io.IOException::class)
        override fun write(theByte: Int) {
            // Encoding suspended?
            if (suspendEncoding) {
                super.out.write(theByte)
                return
            }   // end if: supsended

            // Encode?
            if (encode) {
                buffer[position++] = theByte.toByte()
                if (position >= bufferLength)
                // Enough to encode.
                {
                    out.write(encode3to4(b4, buffer, bufferLength))

                    lineLength += 4
                    if (breakLines && lineLength >= MAX_LINE_LENGTH) {
                        out.write(NEW_LINE.toInt())
                        lineLength = 0
                    }   // end if: end of line

                    position = 0
                }   // end if: enough to output
            }   // end if: encoding
            else {
                // Meaningful Base64 character?
                if (DECODABET[theByte and 0x7f] > WHITE_SPACE_ENC) {
                    buffer[position++] = theByte.toByte()
                    if (position >= bufferLength)
                    // Enough to output.
                    {
                        val len = Base64.decode4to3(buffer!!, 0, b4, 0)
                        out.write(b4, 0, len)
                        //out.write( Base64.decode4to3( buffer ) );
                        position = 0
                    }   // end if: enough to output
                }   // end if: meaningful base64 character
                else if (DECODABET[theByte and 0x7f] != WHITE_SPACE_ENC) {
                    throw java.io.IOException("Invalid character in Base64 data.")
                }   // end else: not white space either
            }// Else, Decoding
            // end else: decoding
        }   // end write


        /**
         * Calls [.write] repeatedly until <var>len</var>
         * bytes are written.
         *
         * @param theBytes array from which to read bytes
         * @param off      offset for array
         * @param len      max number of bytes to read into array
         */
        @Throws(java.io.IOException::class)
        override fun write(theBytes: ByteArray, off: Int, len: Int) {
            // Encoding suspended?
            if (suspendEncoding) {
                super.out.write(theBytes, off, len)
                return
            }   // end if: supsended

            for (i in 0 until len) {
                write(theBytes[off + i].toInt())
            }   // end for: each byte written

        }   // end write


        /**
         * Method added by PHIL. [Thanks, PHIL. -Rob]
         * This pads the buffer without closing the stream.
         */
        @Throws(java.io.IOException::class)
        fun flushBase64() {
            if (position > 0) {
                if (encode) {
                    out.write(encode3to4(b4, buffer, position))
                    position = 0
                }   // end if: encoding
                else {
                    throw java.io.IOException("Base64 input not properly padded.")
                }   // end else: decoding
            }   // end if: buffer partially full

        }   // end flush


        /**
         * Flushes and closes (I think, in the superclass) the stream.
         */
        @Throws(java.io.IOException::class)
        override fun close() {
            // 1. Ensure that pending characters are written
            flushBase64()

            // 2. Actually close the stream
            // Base class both flushes and closes.
            super.close()

            buffer = null
            out = null
        }   // end close


        /**
         * Suspends encoding of the stream.
         * May be helpful if you need to embed a piece of
         * base640-encoded data in a stream.
         */
        @Throws(java.io.IOException::class)
        fun suspendEncoding() {
            flushBase64()
            this.suspendEncoding = true
        }   // end suspendEncoding


        /**
         * Resumes encoding of the stream.
         * May be helpful if you need to embed a piece of
         * base640-encoded data in a stream.
         */
        fun resumeEncoding() {
            this.suspendEncoding = false
        }   // end resumeEncoding


    }
    /**
     * Constructs a [Base64.OutputStream] in ENCODE mode.
     *
     * @param out the <tt>java.io.OutputStream</tt> to which data will be written.
     */// end constructor
    // end inner class OutputStream


}
/**
 * Defeats instantiation.
 */
/**
 * Serializes an object and returns the Base64-encoded
 * version of that serialized object. If the object
 * cannot be serialized or there is another error,
 * the method will return <tt>null</tt>.
 * The object is not GZip-compressed before being encoded.
 *
 * @param serializableObject The object to encode
 * @return The Base64-encoded object
 */// end encodeObject
/**
 * Encodes a byte array into Base64 notation.
 * Does not GZip-compress data.
 *
 * @param source The data to convert
 */// end encodeBytes
/**
 * Encodes a byte array into Base64 notation.
 * Does not GZip-compress data.
 *
 * @param source The data to convert
 * @param off    Offset in array where conversion should begin
 * @param len    Length of data to convert
 */// end encodeBytes
// end class Base64
