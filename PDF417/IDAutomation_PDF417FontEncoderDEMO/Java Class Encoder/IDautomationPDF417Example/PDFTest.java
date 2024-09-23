//*****************************************************************
// 
//  Simple PDF417 Encoder Example 1.0
//  
//  Copyright, IDAutomation.com, Inc. 2001. All rights reserved.
//  
//  http://www.IDAutomation.com/
//  
//  NOTICE:
//  You may incorporate our Source Code in your application
//  only if you own a valid PDF417 Font License
//  from IDAutomation.com, Inc. and the copyright notices 
//  are not removed from the source code.
//  
//*****************************************************************

import java.io.*;
import IDautomationPDFE.*;

class PDFTest
{
    public static void main ( String [] args )
    {
	//Here is the data we will encode
        String dataToEncode = "This is a test of the IDAutomation.com PDF417 Java Encoder.";

	//NOTE: "PDF417Encoder" is the class of the encoder
	PDF417Encoder pdfe=new PDF417Encoder();

      	System.out.println("\n"+ "This is an example of the PDF417 Java Font Encoder."+"\n");

	// This is an example of setting the number of columns (width)
	pdfe.PDFColumns = 7;

	// This is an example of formatting data to the font...
	System.out.println( pdfe.fontEncode(dataToEncode) );

     	System.out.println("\n"+ "The above code will create a PDF417 symbol when");
     	System.out.println("displayed or printed with the PDF417 font."+"\n"+"\n");
    }

}
