/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.hslf.model;

import junit.framework.TestCase;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.hslf.usermodel.RichTextRun;
import org.apache.poi.hslf.HSLFSlideShow;

import java.awt.*;
import java.awt.Rectangle;
import java.io.ByteArrayOutputStream;
import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

/**
 * Test setting text properties of newly added TextBoxes
 *
 * @author Yegor Kozlov
 */
public class TestSetBoldItalic extends TestCase {
    /**
     * Verify that we can add TextBox shapes to a slide
     * and set some of the style attributes
     */
    public void testTextBoxWrite() throws Exception {
        //String dirname = System.getProperty("HSLF.testdata.path");
        //String filename = dirname + "/with_textbox.ppt";
        //new SlideShow(new HSLFSlideShow(filename));

        SlideShow ppt = new SlideShow();
        Slide sl = ppt.createSlide();
        RichTextRun rt;

        String val = "Hello, World!";

        // Create a new textbox, and give it lots of properties
        TextBox txtbox = new TextBox();
        txtbox.setText(val);
        txtbox.setFontSize(42);
        txtbox.setBold(true);
        txtbox.setItalic(true);
        txtbox.setUnderline(false);
        sl.addShape(txtbox);

        // Check it before save
        rt = txtbox.getRichTextRuns()[0];
        assertEquals(val, rt.getText());
        assertEquals(42, rt.getFontSize());
        assertTrue(rt.isBold());
        assertTrue(rt.isItalic());

        // Serialize and read again
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);
        out.close();

        ppt = new SlideShow(new HSLFSlideShow(new ByteArrayInputStream(out.toByteArray())));
        sl = ppt.getSlides()[0];

        txtbox = (TextBox)sl.getShapes()[0];
        rt = txtbox.getRichTextRuns()[0];

        // Check after save
        assertEquals(val, rt.getText());
        assertEquals(42, rt.getFontSize());
        assertTrue(rt.isBold());
        assertTrue(rt.isItalic());
        assertFalse(rt.isUnderlined());
    }

}
