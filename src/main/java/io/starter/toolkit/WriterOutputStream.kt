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

import java.io.IOException
import java.io.OutputStream
import java.io.Writer
import java.nio.ByteBuffer
import java.nio.CharBuffer
import java.nio.charset.Charset
import java.nio.charset.CharsetDecoder
import java.nio.charset.CoderResult
import java.nio.charset.CodingErrorAction

/**
 * A `WriterOutputStream` is a bridge from byte streams to character
 * streams: bytes written to it are decoded into characters using a specified
 * [charset][Charset]. The charset that it uses may be specified by
 * name or given explicitly, or the system default charset may be used. The
 * decoded characters are written to a provided [Appendable], which will
 * usually be a [Writer].
 *
 *
 * The input is buffered so that writes don't need to be aligned to
 * character boundaries. Because conversion is only performed when the input
 * buffer is full (or when the stream is flushed) output also behaves as if it
 * were buffered. It is therefore usually unnecessary to externally buffer the
 * input or output.
 *
 *
 * In some charsets some or all characters are represented by multi-byte
 * sequences. If a byte sequence is encountered that is not valid in the input
 * charset or that cannot be mapped to a valid Unicode character it will be
 * replaced in the output with the value `"\uFFFD"`. If more control
 * over the decoding process is required use [CharsetDecoder].
 *
 * @see InputStreamReader
 *
 * @see OutputStreamWriter
 */
class WriterOutputStream
/**
 * Creates a `WriterOutputStream` with the given charset.
 *
 * @param target  the sink for the decoded characters
 * @param charset the character set with which to interpret the input bytes
 */
@JvmOverloads constructor(private val target: Appendable, charset: Charset = Charset.defaultCharset()) : OutputStream() {
    private var decoder: CharsetDecoder? = null
    private val bytesPerChar: Float

    private var inputBuffer: ByteBuffer? = null
    private var outputBuffer: CharBuffer? = null

    /**
     * Creates a `WriterOutputStream` with the named charset.
     *
     * @param target  the sink for the decoded characters
     * @param charset the character set with which to interpret the input bytes
     */
    constructor(target: Appendable, charset: String) : this(target, Charset.forName(charset)) {}

    init {

        bytesPerChar = charset.newEncoder().maxBytesPerChar()
        decoder = charset.newDecoder()
        decoder!!.onMalformedInput(CodingErrorAction.REPLACE)
        decoder!!.onUnmappableCharacter(CodingErrorAction.REPLACE)

        inputBuffer = ByteBuffer.allocate(
                Math.ceil((BUFFER_SIZE * bytesPerChar).toDouble()).toInt())
        outputBuffer = CharBuffer.allocate(BUFFER_SIZE)
    }

    @Synchronized
    @Throws(IOException::class)
    override fun write(buffer: ByteArray, offset: Int, length: Int) {
        var offset = offset
        var length = length
        if (null == decoder)
            throw IOException("this stream has been closed")

        // if the input buffer is too full decode it first
        if (inputBuffer!!.remaining() < bytesPerChar)
            this.decodeInputBuffer()

        // Append the input to the buffer if it'll fit. If not and there are
        // bytes left in the buffer fill it anyway so we don't lose them.
        if (length <= inputBuffer!!.remaining() || inputBuffer!!.position() > 0) {
            var fill = Math.min(inputBuffer!!.remaining(), length)
            inputBuffer!!.put(buffer, offset, fill)

            // if we've buffered the entire input, we're done
            if (fill == length) return

            // otherwise, decode the input buffer
            inputBuffer!!.flip()
            this.decode(inputBuffer)

            fill -= inputBuffer!!.remaining()
            offset += fill
            length -= fill
            inputBuffer!!.clear()
        }

        // if the remaining input won't fit in the buffer decode it directly
        if (length > inputBuffer!!.remaining()) {
            val tempBuffer = ByteBuffer.wrap(buffer, offset, length)
            this.decode(tempBuffer)

            // if any bytes are left over, put them in the input buffer
            if (tempBuffer.hasRemaining())
                inputBuffer!!.put(tempBuffer)
        } else
            inputBuffer!!.put(buffer, offset, length)// otherwise, just append it to the buffer
    }

    @Synchronized
    @Throws(IOException::class)
    override fun write(b: Int) {
        if (null == decoder)
            throw IOException("this stream has been closed")

        // if the buffer is full, decode it first
        if (!inputBuffer!!.hasRemaining())
            this.decodeInputBuffer()

        // append the input to the buffer
        inputBuffer!!.put(b.toByte())
    }

    /**
     * Flushes the input buffer through the decoder.
     * If the input buffer ends with an incomplete character it will remain in
     * the buffer. If the underlying character sink is a [Writer] its
     * `flush` method will be called after all buffered input has
     * been flushed.
     */
    @Throws(IOException::class)
    override fun flush() {
        synchronized(this) {
            if (null == decoder)
                throw IOException("this stream has been closed")

            this.decodeInputBuffer()
        }

        // flush the underlying Writer, if any
        if (target is Writer)
            target.flush()
    }

    @Throws(IOException::class)
    private fun decodeInputBuffer() {
        inputBuffer!!.flip()
        this.decode(inputBuffer)
        inputBuffer!!.compact()
    }

    @Throws(IOException::class)
    private fun decode(bytes: ByteBuffer) {
        var result: CoderResult

        do {
            outputBuffer!!.clear()
            result = decoder!!.decode(bytes, outputBuffer, false)

            outputBuffer!!.flip()
            target.append(outputBuffer)
        } while (result === CoderResult.OVERFLOW)
    }

    /**
     * Closes the stream, flushing it first.
     * If any partial characters remain in the input buffer the replacement
     * value will be output in their place. If the underlying character sink is
     * a [Writer] its `close` method will be called after all
     * buffered input has been flushed. Once the stream has been closed further
     * calls to `write` or `flush` will cause an
     * `IOException` to be thrown. Closing a previously closed
     * stream has no effect.
     */
    @Throws(IOException::class)
    override fun close() {
        synchronized(this) {
            var result: CoderResult
            if (null == decoder) return

            // flush the input buffer
            inputBuffer!!.flip()
            do {
                outputBuffer!!.clear()
                result = decoder!!.decode(inputBuffer, outputBuffer, true)

                outputBuffer!!.flip()
                target.append(outputBuffer)
            } while (result === CoderResult.OVERFLOW)

            // flush the decoder
            do {
                outputBuffer!!.clear()
                result = decoder!!.flush(outputBuffer)

                outputBuffer!!.flip()
                target.append(outputBuffer)
            } while (result === CoderResult.OVERFLOW)

            // release the buffers and decoder
            inputBuffer = null
            outputBuffer = null
            decoder = null
        }

        // close the underlying Writer, if any
        if (target is Writer)
            target.close()
    }

    companion object {
        private val BUFFER_SIZE = 8192
    }
}
/**
 * Creates a `WriterOutputStream` with the default charset.
 *
 * @param target the sink for the decoded characters
 */
