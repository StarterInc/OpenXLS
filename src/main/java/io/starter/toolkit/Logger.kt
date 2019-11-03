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
/**
 * Logger.java
 */
package io.starter.toolkit

import java.io.*
import java.nio.charset.Charset
import java.text.SimpleDateFormat
import java.util.Date

/** System-wide Logging facility
 *
 * <br></br>
 * Logger can be used to output standardized messages to System.out and System.err, as well as
 * to pluggable Logger implementations.
 * <br></br><br></br>
 * To install a custom Logger implementation, instantiate a class that implements Logger, then
 * set the system property: "io.starter.toolkit.logger"
 * <br></br><br></br>
 * For example:
 * <pre>
 * CustomLog mylogr = new CustomLog();
 * Properties props = System.getProperties();
 * props.put("io.starter.toolkit.logger", mylogr );
</pre> *
 * <br></br>
 * The default Logger settings can be controlled using System properties.
 * <pre>
 * props.put("io.starter.toolkit.logger.dateformat", "MMM yyyy mm:ss" );
 * props.put("io.starter.toolkit.logger.dateformat", "none" );
 *
</pre> *
 *
 *
 */
class Logger : PrintStream {

    @Deprecated("Just use <code>this</code>. ")
    protected var ous: PrintStream = this
    private val lineBuffer = StringBuffer()

    constructor(target: Logger) : this() {
        this.init(target)
    }

    @Throws(UnsupportedEncodingException::class)
    constructor(target: Logger, charset: String) : this(charset) {
        this.init(target)
    }

    @JvmOverloads
    constructor(target: OutputStream, autoFlush: Boolean = false) : this(OutputStreamWriter(target), autoFlush) {
    }

    @Throws(UnsupportedEncodingException::class)
    constructor(target: OutputStream, charset: String, autoFlush: Boolean) : this(OutputStreamWriter(target, charset), charset, autoFlush) {
    }

    @JvmOverloads
    constructor(target: Writer, autoFlush: Boolean = false) : this() {
        this.init(target, autoFlush)
    }

    @Throws(UnsupportedEncodingException::class)
    constructor(target: Writer, charset: String, autoFlush: Boolean) : this(charset) {
        this.init(target, autoFlush)
    }

    private constructor() : super(IndirectOutputStream(), true) {
        (out as IndirectOutputStream).sink = WriterOutputStream(this,
                Charset.defaultCharset())
    }

    @Throws(UnsupportedEncodingException::class)
    private constructor(charset: String) : super(IndirectOutputStream(), true, charset) {
        (out as IndirectOutputStream).sink = WriterOutputStream(this, charset)
    }

    private fun init(target: Logger) {
        targetLogger = target
        targetWriter = null
        autoFlush = false // has no meaning for Logger target
    }

    private fun init(target: Writer, autoFlush: Boolean) {
        targetLogger = null
        targetWriter = BufferedWriter(target)
        // autoFlush = autoFlush;
    }

    /** Installs this logger as the default logger and replaces the standard
     * output and error streams.
     */
    fun install() {
        logger = this
        System.setOut(this)
        System.setErr(this)
    }

    fun log(message: String, ex: Exception, trace: Boolean) {
        if (null != targetLogger)
            targetLogger!!.log(message, ex, trace)
        else
            log(formatThrowable(message, ex, trace))
    }

    fun log(message: String, ex: Exception) {
        if (null != targetLogger)
            targetLogger!!.log(message, ex)
        else
            log(formatThrowable(message, ex, false))
    }

    /* ---------- PrintStream methods ---------- */

    fun logLine() {
        synchronized(lineBuffer) {
            // if the line buffer ends with a newline, strip it
            val length = lineBuffer.length
            if (length >= endl.length && endl == lineBuffer
                            .substring(length - endl.length, length))
                lineBuffer.setLength(length - endl.length)

            // log and reset the line buffer but don't log empty lines
            if (lineBuffer.length > 0) {
                log(lineBuffer.toString())
                lineBuffer.setLength(0)
            }
        }
    }

    override fun append(value: Char): Logger {
        lineBuffer.append(value)
        return this
    }

    override fun append(value: CharSequence?): Logger {
        lineBuffer.append(value)
        return this
    }

    override fun append(value: CharSequence?, start: Int, end: Int): Logger {
        lineBuffer.append(value, start, end)
        return this
    }

    override fun print(b: Boolean) {
        lineBuffer.append(b)
    }

    override fun print(c: Char) {
        lineBuffer.append(c)
    }

    override fun print(i: Int) {
        lineBuffer.append(i)
    }

    override fun print(l: Long) {
        lineBuffer.append(l)
    }

    override fun print(f: Float) {
        lineBuffer.append(f)
    }

    override fun print(d: Double) {
        lineBuffer.append(d)
    }

    override fun print(s: CharArray) {
        lineBuffer.append(s)
    }

    override fun print(s: String?) {
        synchronized(lineBuffer) {
            lineBuffer.append(s)
            if (s!!.endsWith(endl))
                this.println()
        }
    }

    override fun print(obj: Any?) {
        lineBuffer.append(obj)
    }

    override fun println(x: Boolean) {
        lineBuffer.append(x)
        this.println()
    }

    override fun println(x: Char) {
        lineBuffer.append(x)
        this.println()
    }

    override fun println(x: Int) {
        lineBuffer.append(x)
        this.println()
    }

    override fun println(x: Long) {
        lineBuffer.append(x)
        this.println()
    }

    override fun println(x: Float) {
        lineBuffer.append(x)
        this.println()
    }

    override fun println(x: Double) {
        lineBuffer.append(x)
        this.println()
    }

    override fun println(x: CharArray) {
        lineBuffer.append(x)
        this.println()
    }

    override fun println(x: String) {
        lineBuffer.append(x)
        this.println()
    }

    override fun println(x: Any?) {
        lineBuffer.append(x)
        this.println()
    }

    override fun println() {
        synchronized(lineBuffer) {
            // flush the input stream into the line buffer
            super.flush()

            // log the current line
            logLine()
        }
    }

    companion object {

        /** Copy of `line.separator` system property to save lookups.  */
        private val endl = System
                .getProperty("line.separator")

        private var targetLogger: Logger? = null
        private var targetWriter: BufferedWriter? = null
        private var autoFlush: Boolean = false

        /* ---------- Logger methods ---------- */

        fun log(message: String) {
            if (null != targetLogger)
                Logger.log(message)
            else
                synchronized(targetWriter) {
                    try {
                        targetWriter!!.write(logDate)
                        targetWriter!!.write(" ")
                        targetWriter!!.write(message)
                        targetWriter!!.newLine()
                        if (autoFlush)
                            targetWriter!!.flush()
                    } catch (ex: IOException) {
                        // we're the logger, so we can't exactly log about it
                        // the interface doesn't support exceptions so just drop it
                    }

                }
        }

        /*
     * ---------- static convenience methods for logging
     * ----------
     */

        val INFO_STRING = ""
        val WARN_STRING = "WARNING: "
        val ERROR_STRING = "ERROR: "

        /** Gets the current system logger.
         */
        /** Replaces the system logger.  */
        var logger: Logger?
            get() {
                var logger: Logger?

                try {
                    logger = System.getProperties()["io.starter.toolkit.logger"] as Logger
                } catch (ex: Exception) {
                    logger = null
                }

                if (null == logger) {
                    if (System.err is Logger)
                        logger = System.err as Logger
                    else {
                        logger = Logger(System.err, true)
                    }
                    logger = logger
                }

                return logger
            }
            set(logger) {
                System.getProperties()["io.starter.toolkit.logger"] = logger
            }

        fun formatThrowable(message: String, ex: Throwable, trace: Boolean): String {
            val writer = StringWriter()
            writer.write(message)

            if (trace) {
                writer.write(endl)
                writer.write(endl)

                val printer = PrintWriter(writer)
                ex.printStackTrace(printer)
                printer.flush()
            } else {
                writer.write(": ")
                writer.write(ex.toString())
            }

            return writer.toString()
        }

        /** Logs a fatal error message to the system logger.
         */
        fun logErr(message: String, ex: Exception) {
            logger!!.log(ERROR_STRING + message, ex)
        }

        /** Logs a fatal error message to the system logger.
         */
        fun logErr(message: String, ex: Throwable) {
            logger
            Logger.log(formatThrowable(ERROR_STRING + message, ex, false))
        }

        /** Logs a fatal error message to the system logger.
         */
        fun logErr(message: String) {
            logger
            Logger.log(ERROR_STRING + message)
        }

        /** Logs a fatal error message to the system logger.
         */
        fun logErr(message: String, ex: Exception, trace: Boolean) {
            logger!!.log(ERROR_STRING + message, ex, trace)
        }

        /** Logs the string conversion of an object to the system logger.
         */
        fun log(`object`: Any) {
            logInfo(`object`.toString())
        }

        /** Logs a non-fatal warning to the system logger.
         */
        fun logWarn(message: String) {
            logger
            Logger.log(WARN_STRING + message)
        }

        /** Logs the string conversion of an exception to the system logger as a
         * fatal error message.
         */
        fun logErr(ex: Exception) {
            logErr(ex.toString())
        }

        /** Logs an informational message to the system logger.
         */
        fun logInfo(message: String) {
            logger
            Logger.log(INFO_STRING + message)
        }

        /** Attempts to replace the standard output stream with a
         * `Logger` instance that writes to the named file. If the
         * operation fails a message will be logged to the system logger and the
         * method will return without throwing an exception.
         */
        fun setOut(filename: String) {
            try {
                val logfile = java.io.File(filename)
                val sysout = FileOutputStream(logfile)
                System.setOut(Logger(sysout))
            } catch (e: Exception) {
                Logger.logErr("Setting System Output Stream in Logger failed: ", e)
            }

        }

        /** Attempts to replace the standard error stream with a
         * `Logger` instance that writes to the named file. If the
         * operation fails a message will be logged to the system logger and the
         * method will return without throwing an exception.
         */
        fun setErr(filename: String) {
            try {
                val logfile = java.io.File(filename)
                val sysout = FileOutputStream(logfile)
                System.setErr(Logger(sysout))
            } catch (e: Exception) {
                Logger.logErr("Setting System Error Stream in Logger failed: ", e)
            }

        }

        /** The default time stamp format for [.getLogDate].  */
        val DATE_FORMAT = "yyyy-MM-dd HH:mm:ss:SSSS"

        private val dateFormat = SimpleDateFormat(
                DATE_FORMAT)
        private var dateSpec = DATE_FORMAT

        /** Returns the current time in a configurable format.
         * If the system property `io.starter.toolkit.logger.dateformat`
         * exists and is a valid date format pattern it will be used. Otherwise the
         * [default format pattern][.DATE_FORMAT] will be used.
         */
        val logDate: String
            get() {
                val spec = System
                        .getProperty("io.starter.toolkit.logger.dateformat")
                if (null != spec) {
                    if ("none".equals(spec, ignoreCase = true))
                        return ""

                    if (dateSpec != spec) {
                        try {
                            dateFormat.applyPattern(spec)
                        } catch (e: IllegalArgumentException) {
                            dateFormat.applyPattern(DATE_FORMAT)
                        }

                        dateSpec = spec
                    }
                }

                return dateFormat.format(Date())
            }
    }

}
