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
package docs.samples.TemplateModifications;
import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import io.starter.OpenXLS.*;
import io.starter.formats.XLS.*;
import io.starter.toolkit.Logger;



/** TestXLS Demonstrates the basic functionality of of OpenXLS
 
 	We read in the Template file and add new cells and modify cells involved in a Chart and formula calculations,
 	then output the results.
 
 
 
 */
public class TestXLS{

    public static void main(String[] args){
        testit t = new testit();
		String s = "Test Successful.";
		
		t.test(s);
    }
}
  
   /** Test the various functions of the OpenXLS classes.
  
   */
class testit{
	public static final String wd = System.getProperty("user.dir")+"/docs/samples/TemplateModifications/";

        public void test(String args){
        try{
            
            // choose your input and output files.
            String finpath = wd + "template.xls";
      	    String foutpath = wd + "xlstestout.xls";
      	    
            //  Read in the workbook
            WorkBookHandle book = new WorkBookHandle(finpath);
     	    io.starter.OpenXLS.util.Logger.log("OpenXLS version " + WorkBookHandle.getVersion() + ".  WorkBook successfully loaded.");
     	    
     	    // Get a handle to the worksheet you want to work with
      	    WorkSheetHandle sheet1 =  book.getWorkSheet("Sheet1");
      	        
      	    /*  The following block shows how to get a handle to a cell,
      	        then get its value, or set a different value for an output
      	        file 
      	    */
      	    CellHandle cell = null;
           cell = sheet1.getCell("C35");
      	    
      	    /*  get the value of this cell.  If you know the datatype
      	        you can use the first method, otherwise cast it to a string
      	        the next two lines essentially do the same thing, but the getStringVal
      	        is more type-safe
      	    */
      	    float fl = cell.getFloatVal();
      	    String s = cell.getStringVal();
      	    io.starter.OpenXLS.util.Logger.log("The Cell Value of Sheet1:C35 is: " + s);
      	    
      	    // Set the value of this cell to something different
      	    cell.setVal(813.12);

			sheet1.add("Created with OpenXLS Version " +  WorkBookHandle.getVersion(), "B3");

      	    /*  Shorthand method for the above, it is not neccessary to 
      	        actually create the cellhandle 
      	    */
			CellHandle d35 = sheet1.getCell("D35");
      	    io.starter.OpenXLS.util.Logger.log("Retrieved Cell: " + d35);

      	    // set the Cell's new value
      	    d35.setVal(7234.94);
      	    
      	    // set new values without getting a CellHandle
      	    sheet1.setVal("E35", 7255.12);
      	    sheet1.setVal("F35", 9342.22);
			
            // set worksheet and table titles
            sheet1.add("This file was modified by the TestXLS program.","B16");

            sheet1.setSheetName("Modified Sheet");
            
            // Set names for Chart
            sheet1.setVal("B35", "Jennifer Jones");
            sheet1.setVal("B36", "Margueritte Jahnnsen");
            sheet1.setVal("B37", "Vishnu Prakash");
            sheet1.setVal("B38", "Carl West");

            // Set number vals for Chart
            String addr = "C";
            for(int r=35;r<=39;r++) {
                addr = "C" + r;
                sheet1.setVal(addr,15.53* r);
            }
            
            for(int r=35;r<=39;r++) {
                addr = "D" + r;
                sheet1.setVal(addr,10.11* r);
            }
            
            for(int r=35;r<=39;r++) {
                addr = "E" + r;
                sheet1.setVal(addr,12.13* r);
            }
            
            for(int r=35;r<=39;r++) {
                addr = "F" + r;
                sheet1.setVal(addr,17.21* r);
            }
            
            
            // Add new Cells individually to the WorkSheet
            sheet1.add(Integer.valueOf(987), "C13");            
            sheet1.add(new Double(23423478.234d),"C14");
            sheet1.add(new Float(23423478.234f),"C15");
			
            // Add new Cells en-masse
            for(int i = 1;i<50;i++){
                addr = "I"+String.valueOf(i);
                sheet1.add("Newly Added Cell " + i, addr);
				addr = "J"+String.valueOf(i);
				sheet1.add(Integer.valueOf(i), addr);
            }

      	    // stream the workbook to a byte array
            try{
          	    java.io.File f = new java.io.File(foutpath);
                FileOutputStream fos = new FileOutputStream(f);
                BufferedOutputStream bbout = new BufferedOutputStream(fos);
                book.write(bbout);
                bbout.flush();
    		    fos.close();
          	} catch (java.io.IOException e){Logger.logInfo("IOException in Tester.  "+e);}  
		   
		    io.starter.OpenXLS.util.Logger.log("Writing: " + foutpath);
		   FileInputStream fis= new FileInputStream(foutpath);
		    book = new WorkBookHandle(fis);
      	    io.starter.OpenXLS.util.Logger.log(book.getName() + " re-read sucessfully.");
      	} catch (java.io.IOException e){io.starter.OpenXLS.util.Logger.log("IOException in Tester.  "+e);
      	} catch (CellNotFoundException e){io.starter.OpenXLS.util.Logger.log("Cell Not Found in Template.  " + e);      	
      	}catch(WorkSheetNotFoundException e){io.starter.OpenXLS.util.Logger.log(e);}
      	catch (Exception e){io.starter.OpenXLS.util.Logger.log(e);}
    }

}