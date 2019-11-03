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

/* Based on Bare Bones Browser Launcher by Dem Pilafian,
 * which is in the public domain. You may find it at:
 * http://www.centerkey.com/java/browser/
 */

/**
 * Launches the user's default browser to display a web page.
 *
 * @author Dem Pilafian
 * @author Sam Hanes
 */
object BrowserLauncher {
    /**
     * List of potential browsers on systems without a default mechanism.
     */
    val browsers = arrayOf("google-chrome", "firefox", "opera", "konqueror", "epiphany", "seamonkey", "galeon", "kazehakase", "mozilla")

    /**
     * The browser that was last successfully run.
     */
    private var browser: String? = null

    /**
     * Opens the specified web page in the user's default browser.
     *
     * @param url the URL of the page to be opened
     * @throws Exception if an error occurred attempting to launch the browser.
     * If the browser is successfully started but later fails for
     * some reason, no exception will be thrown.
     */
    @Throws(Exception::class)
    fun open(url: String) {
        // Attempt to use the Desktop class from JDK 1.6+ (even if on 1.5)
        // This uses reflection to mimic the call:
        // java.awt.Desktop.getDesktop().browse( java.net.URI.create(url) );
        try {
            val desktop = Class.forName("java.awt.Desktop")
            desktop.getDeclaredMethod(
                    "browse", *arrayOf(java.net.URI::class.java))
                    .invoke(
                            desktop.getDeclaredMethod(
                                    "getDesktop", *null as Array<Class<*>>?)
                                    .invoke(null, *null as Array<Any>?),
                            java.net.URI.create(url))

            // If that didn't throw an exception, we're done
            return
        } catch (e: ClassNotFoundException) {
            // Intentionally empty, falls back to platform-dependent code
        } catch (e: NoSuchMethodException) {
            // Intentionally empty, falls back to platform-dependent code
        } catch (e: Exception) {
            throw Exception("failed to launch browser", e)
        }

        val osName = System.getProperty("os.name")
        try {
            // If this is OS X, use the FileManager class
            if (osName.startsWith("Mac OS"))
                Class.forName("com.apple.eio.FileManager")
                        .getDeclaredMethod("openURL", *arrayOf(String::class.java))
                        .invoke(null, url)
            else if (osName.startsWith("Windows"))
                Runtime.getRuntime().exec(
                        "rundll32 url.dll,FileProtocolHandler $url")
            else {
                // If we haven't found a browser yet, try some possible ones
                if (browser == null) {
                    for (idx in browsers.indices) {
                        if (Runtime.getRuntime().exec(
                                        arrayOf("which", browsers[idx]))
                                        .waitFor() == 0) {
                            browser = browsers[idx]
                        }
                    }

                    // If we couldn't find one, throw an exception
                    if (browser == null)
                        throw Exception("no browser found")
                }

                // Call the browser with the URL
                Runtime.getRuntime().exec(arrayOf<String>(browser, url))
            }// Otherwise, assume this is a POSIX-like system and
            // start trying possible browser commands
            // If this is Windows, call the FileProtocolHandler via rundll
        } catch (e: Exception) {
            throw Exception("failed to launch browser", e)
        }

    }
}
