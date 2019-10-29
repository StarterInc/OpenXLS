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
/**
 * FormatCache.java
 */
package io.starter.OpenXLS;

import io.starter.toolkit.Logger;

import java.util.Iterator;
import java.util.Map;


/** Handles the caching of the Formats

 This class is no longer in use nor needed
 *
 *@deprecated
 */
// is this a valid class anymore?
public class FormatCache {

    Map<?, ?> mpx = new java.util.HashMap();

    /** Consolidate all identical formats to avoid too many formats errors
     *
     * @deprecated
     */
    public void pack() {
        Iterator<?> itx = mpx.keySet().iterator();
        while (itx.hasNext()) {
            Object oby = mpx.get(itx.next());
            FormatHandle thisfmt = (FormatHandle) oby;
            thisfmt = this.get(thisfmt);
        }
    }

    /**
     * @deprecated
     * @return
     */
    public FormatHandle get(FormatHandle fmx) {
        String fmt = fmx.toString();
        if (!mpx.containsKey(fmt)) {
            Logger.logErr("missing in cache: FH " + fmt);
        }
        FormatHandle ret = (FormatHandle) mpx.get(fmt);
        return ret;
    }

    /**
     * @deprecated
     * @return
     */
    public Object get(String f) {
        Object myo = mpx.get(f);
        return myo;
    }

    /**
     * @deprecated
     * @return
     */
    public int getInt(String f) {
        int findex = -1;
        if (mpx.containsKey(f)) findex = ((Integer) mpx.get(f)).intValue();
        return findex;
    }

}
