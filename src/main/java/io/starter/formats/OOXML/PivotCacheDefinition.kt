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
package io.starter.formats.OOXML

import io.starter.OpenXLS.WorkBookHandle
import io.starter.toolkit.Logger
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserFactory

import java.io.InputStream

class PivotCacheDefinition(ref: String, sheet: String,
                           /**
                            * return the pivot cache id
                            */
                           val iCache: Int) : OOXMLElement {
    /**
     * returns the data source reference
     *
     * @return
     */
    val ref: String? = null
    /**
     * return the sheet the data reference is on
     *
     * @return
     */
    val sheet: String? = null

    override// TODO: Finish
    val ooxml: String?
        get() = null

    override fun cloneElement(): OOXMLElement? {
        return null
    }

    init {
        this.ref = ref
        this.sheet = sheet
    }

    companion object {

        private val serialVersionUID = -5070227633357072878L

        fun parseOOXML(bk: WorkBookHandle, cacheid: String?, ii: InputStream): PivotCacheDefinition? {
            var ref: String? = null
            var sheet: String? = null
            var icache = 1
            try {
                val factory = XmlPullParserFactory.newInstance()
                factory.isNamespaceAware = true
                val xpp = factory.newPullParser()
                xpp.setInput(ii, null) // using XML 1.0 specification
                var eventType = xpp.eventType

                while (eventType != XmlPullParser.END_DOCUMENT) {
                    if (eventType == XmlPullParser.START_TAG) {
                        val tnm = xpp.name
                        if (tnm == "pivotCacheDefinition") { // get attributes
                            //  r:id="rId1" refreshedBy="Kaia" refreshedDate="41038.467970833335" createdVersion="1" refreshedVersion="3" recordCount="4" upgradeOnRefresh="1">
                        } else if (tnm == "cacheSource") {
                            for (z in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(z)
                                val v = xpp.getAttributeValue(z)
                                if (nm == "type")
                                    if (v != "worksheet") {
                                        // consolidation, external, scenario --
                                        Logger.logWarn("PivotCacheDefinition: Data Souce $v Not Supported")
                                        return null
                                    }
                            }
                        } else if (tnm == "worksheetSource") {
                            // ref, sheet, id (sheet rid), name (range)
                            for (z in 0 until xpp.attributeCount) {
                                val nm = xpp.getAttributeName(z)
                                val v = xpp.getAttributeValue(z)
                                if (nm == "ref") {
                                    ref = v
                                } else if (nm == "sheet") {
                                    sheet = v
                                } else if (nm == "name") {
                                    ref = v
                                } else if (nm == "id") {

                                }
                            }
                        } else if (tnm == "cacheFields") {
                        }
                    } else if (eventType == XmlPullParser.END_TAG) { // go to end of file
                    }
                    eventType = xpp.next()
                }
            } catch (e: Exception) {
                Logger.logErr("PivotCacheDefinition.parseOOXML: $e")
            }

            if (cacheid != null) {
                // KSC: TESTING!!!			icache= Integer.valueOf((String)cacheid)+1;
                icache = Integer.valueOf(1)!!
            }
            icache = bk.workBook!!.addPivotStream(ref, sheet, icache)
            return PivotCacheDefinition(ref, sheet, icache)
        }
    }


}
