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
package io.starter.formats.LEO


/**
 * LEO File BIGBLOCK Information Record.
 *
 *
 * These blocks of data contain information related to the
 * LEO file format data blocks.
 *
 *
 * depending on the size and complexity of the file,
 * these records may all be contained in the 'header'
 * block of the LEO Stream File.
 *
 *
 * In files over that size, one or more BIGBLOCK records
 * are inserted directly in the midst of substream records.
 */
class BIGBLOCK : BlockImpl() {


    /**
     * returns the int representing the block type
     */
    override val blockType: Int
        get() = Block.BIG

    companion object {

        /**
         * serialVersionUID
         */
        private val serialVersionUID = 839197117033095054L
        val SIZE = 512
    }


}