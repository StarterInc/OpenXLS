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
package io.starter.OpenXLS

import io.starter.formats.LEO.LEOFile
import io.starter.formats.XLS.OOXMLAdapter
import io.starter.toolkit.Logger
import io.starter.toolkit.TempFileManager

import java.io.*
import java.net.URL
import java.net.URLConnection
import java.nio.charset.StandardCharsets
import java.util.Collections
import java.util.HashMap

/**
 * Functionality common to all document types.
 */
abstract class DocumentHandle : Document, Handle, Closeable {

    /**
     * The level of debugging output requested by the user.
     * Higher values should produce more output.
     */
    // TODO: should the debug level be static?
    protected var DEBUGLEVEL = 0

    /**
     * The user-visible display name or title of this document.
     */
    protected var name: String? = null

    /**
     * The file associated with this document.
     * This will generally be the file the document was parsed from, if any.
     */
    /**
     * Gets the file associated with this document.
     * For documents read in from a file, this defaults to that file. If no
     * file is associated with this document, for example if the document was
     * parsed from a stream, this may return `null`.
     */
    /**
     * Sets the file associated with this document.
     */
    var file: File? = null

    /**
     * Store for workbook properties.
     */
    private var props: MutableMap<String, Any> = HashMap()

    /**
     * Handling for a streaming worksheet based workbook
     */
    private var streamingSheets = false

    /**
     * Retrieves a Map containing the workbook properties store.
     * This is not an Excel-compatible feature.
     *
     * @return an immutable Map containing the current workbook properties
     */
    /**
     * Replaces the workbook properties with the values in a given Map.
     * This is not an Excel-compatible feature.
     *
     * @param properties the values that will replace the existing properties
     */
    var properties: Map<String, Any>
        get() = Collections.unmodifiableMap(props)
        set(properties) {
            props = HashMap()
            props.putAll(properties)
        }

    /**
     * Gets the file name associated with this document.
     * For documents read in from a file, this defaults to that file. If no
     * file is associated with this document, for example if the document was
     * parsed from a stream, this may return `null`.
     *
     */
    /**
     * Sets the file name associated with this document.
     *
     */
    var fileName: String
        @Deprecated("Use {@link #getFile()} instead.")
        get() = if (file != null) file!!.path else "New Document.doc"
        @Deprecated("Use {@link #setFile(File)} instead.")
        set(name) {
            file = File(name).absoluteFile
        }

    /**
     * Gets the constant representing this document's native format.
     */
    abstract val format: Int

    /**
     * Gets the file name extension for this document's native format.
     */
    abstract val fileExtension: String

    /**
     * default constructor
     */
    constructor() : super() {

    }

    /**
     * Apr 5, 2011
     *
     * @param urlx
     */
    constructor(urlx: InputStream) : super() {
        // TODO Auto-generated constructor stub
        Logger.logErr("DocumentHandle InputStream Constructor Not Implemented")
    }

    constructor(input: File) {
        // TODO Auto-generated constructor stub
    }

    /**
     * Retrieves a property in the workbook property store.
     * This is not an Excel-compatible feature.
     *
     * @param name the name of the property to retrieve
     * @return the value of the requested property or null if it doesn't exist
     */
    override fun getProperty(name: String): Any {
        return props[name]
    }

    /**
     * Sets the value of a property in the workbook property store.
     * This is not an Excel-compatible feature.
     *
     * @param name  the name of the property which should be updated
     * @param value the value to which the property should be set
     */
    override fun addProperty(name: String, `val`: Any) {
        props[name] = `val`
    }

    /**
     * Sets the user-visible descriptive name or title of this document.
     * Some formats will persist this setting in the document itself.
     */
    override fun setName(nm: String) {
        name = nm
    }

    /**
     * Handling for streaming sheets.  Currently this is in development and unsupported
     *
     * @param streamSheets
     */
    fun setStreamingSheets(streamSheets: Boolean) {
        this.streamingSheets = streamSheets
    }

    /**
     * Sets the debugging output level.
     * Higher values will produce more output. Output at higher values will
     * generally only be of use to OpenXLS developers. Increased output incurs
     * a performance penalty, so it is recommended this be left at zero unless
     * you are reporting a bug.
     */
    override fun setDebugLevel(level: Int) {
        DEBUGLEVEL = level
    }

    fun getDebugLevel(): Int {
        return DEBUGLEVEL
    }

    /**
     * Gets the user-visible descriptive name or title of this document.
     */
    override fun getName(): String {
        return if (name != null)
            name
        else
            "Untitled Document"
    }

    /**
     * Resets the document state to what it was when it was loaded.
     *
     * @throws UnsupportedOperationException if there is not sufficient data
     * available to perform the reversion
     */
    abstract override fun reset()

    /**
     * Writes the document to the given stream in the requested format.
     *
     * @param dest   the stream to which the document should be written
     * @param format the constant representing the desired output format
     * @throws IllegalArgumentException if the given type code is invalid
     * @throws IOException              if an error occurs while writing to the stream
     */
    @Throws(IOException::class)
    abstract override fun write(dest: OutputStream, format: Int)

    /**
     * Writes the document to the given stream in its native format.
     *
     * @param dest the stream to which the document should be written
     * @throws IOException if an error occurs while writing to the stream
     */
    @Throws(IOException::class)
    override fun write(dest: OutputStream) {
        this.write(dest, FORMAT_NATIVE)
    }

    /**
     * Writes the document to the given file in the requested format.
     *
     * @param file   the path to which the document should be written
     * @param format the constant representing the desired output format
     * @throws IllegalArgumentException if the given type code is invalid
     * @throws IOException              if an error occurs while writing to the file
     */
    @Throws(IOException::class)
    override fun write(file: File, format: Int) {
        if (format > WorkBookHandle.FORMAT_XLS && this.file != null)
            OOXMLAdapter.refreshPassThroughFiles(this as WorkBookHandle)

        if (file.exists())
            file.delete() // try this
        val stream = BufferedOutputStream(
                FileOutputStream(file))
        this.write(stream, format)
        this.file = file // necesary for OOXML re-write ...
        stream.flush()
        stream.close()
    }

    /**
     * Writes the document to the given file in its native format.
     *
     * @param file the path to which the document should be written
     * @throws IOException if an error occurs while writing to the stream
     */
    @Throws(IOException::class)
    override fun write(file: File) {
        this.write(file, FORMAT_NATIVE)
    }

    /**
     * Returns a string representation of the object.
     * This is currently equivalent to [.getName].
     */
    override fun toString(): String {
        return getName()
    }

    companion object {

        var workingdir = System.getProperty("user.dir") + "/tmp"

        /**
         * Format constant for the most appropriate format for this document.
         * If the document was read in from a file, this is usually the format that
         * was read in.
         */
        val FORMAT_NATIVE = 0

        /**
         * Gets the OpenXLS version number.
         */
        val version: String
            get() = GetInfo.getVersion()

        /**
         * Looks for magic numbers in the given input data and attempts to parse
         * it with an appropriate `DocumentHandle` subclass. Detection
         * is performed on a best-effort basis and is not guaranteed to be accurate.
         *
         * @throws IOException       if an error occurs while reading from the stream
         * @throws WorkBookException if parsing fails
         */
        @Throws(IOException::class)
        fun getInstance(input: InputStream): DocumentHandle {
            val bufferedStream = BufferedInputStream(input)
            // read in that start of the file for checking magic numbers
            val headerBytes: ByteArray
            val count: Int
            // make sure the file is long enough to get magic numbers
            bufferedStream.mark(1028)
            headerBytes = ByteArray(512)
            count = bufferedStream.read(headerBytes)
            bufferedStream.reset()

            // if it starts with the LEO magic number check the header
            if (LEOFile.checkIsLEO(headerBytes, count)) {
                val leo = LEOFile(bufferedStream)

                return if (leo.hasWorkBook())
                    WorkBookHandle(leo)
                else
                    throw WorkBookException(
                            "input is LEO but no supported format detected", -1)
            }

            val headerString: String
            headerString = String(headerBytes, 0, count, StandardCharsets.UTF_8)

            // if it's a ZIP archive, try parsing as OOXML
            if (headerString.startsWith("PK")) {
                return WorkBookHandle(bufferedStream)
            }

            if (headerString.indexOf(",") > -1 && headerString.indexOf(",") > -1) {
                // init a blank workbook
                val book = WorkBookHandle()

                // map CSV into workbook
                try {
                    val sheet = book.getWorkSheet(0)
                    sheet.readCSV(BufferedReader(
                            InputStreamReader(bufferedStream)))
                    return book
                } catch (e: Exception) {
                    throw WorkBookException(
                            "Error encountered importing CSV: $e",
                            WorkBookException.ILLEGAL_INIT_ERROR)
                }

            } else {
                throw WorkBookException("unknown file format", -1)
            }
        }

        /**
         * Downloads the resource at the given URL to a temporary file.
         *
         * @param u the URL representing the resource to be downloaded
         * @return the path to a temporary file containing the downloaded resource
         * or `null` if an error occurred
         */
        @Deprecated("The download should be handled outside OpenXLS.\n" +
                "      There is no specific replacement for this method.")
        protected fun getFileFromURL(u: URL): File? {
            try {
                val fx = TempFileManager.createTempFile("upload-" + System.currentTimeMillis(), ".tmp")

                val uc = u.openConnection()
                val contentType = uc.contentType
                val contentLength = uc.contentLength
                if (contentType.startsWith("text/") || contentLength == -1) {
                    throw IOException("This is not a binary file.")
                }
                val raw = uc.getInputStream()
                val `in` = BufferedInputStream(raw)
                val data = ByteArray(contentLength)
                var bytesRead = 0
                var offset = 0
                while (offset < contentLength) {
                    bytesRead = `in`.read(data, offset, data.size - offset)
                    if (bytesRead == -1)
                        break
                    offset += bytesRead
                }
                `in`.close()

                if (offset != contentLength) {
                    throw IOException("Only read " + offset
                            + " bytes; Expected " + contentLength + " bytes")
                }

                // String filename =
                // u.getFile().substring(filename.lastIndexOf('/') + 1);
                val out = FileOutputStream(fx)
                out.write(data)
                out.flush()
                out.close()
                return fx
            } catch (e: Exception) {
                Logger.logErr("Could not load WorkBook from URL: $e")
                return null
            }

        }
    }
}