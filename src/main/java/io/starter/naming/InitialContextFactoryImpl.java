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
package io.starter.naming;

import java.util.*;
import javax.naming.*;
import javax.naming.spi.InitialContextFactory;
/**
 *  Read the details at: http://java.sun.com/j2se/1.3/docs/guide/jndi/spec/spi/jndispiTOC.fm.html
 *
 * 
 */
public class InitialContextFactoryImpl implements InitialContextFactory {
	
	
	public Context getInitialContext(Hashtable env)         
		throws NamingException{
			InitialContextImpl contimple = new InitialContextImpl();
			
			return contimple;
		}
}
