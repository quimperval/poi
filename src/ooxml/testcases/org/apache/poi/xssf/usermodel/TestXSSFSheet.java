/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.xssf.usermodel;

import org.apache.poi.ss.usermodel.BaseTestSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.XSSFITestDataProvider;
import org.apache.poi.xssf.XSSFTestDataSamples;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.helpers.ColumnHelper;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComments;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPane;


public class TestXSSFSheet extends BaseTestSheet {

    @Override
    protected XSSFITestDataProvider getTestDataProvider() {
        return XSSFITestDataProvider.getInstance();
    }

    //TODO column styles are not yet supported by XSSF
    public void testDefaultColumnStyle() {
        //super.testDefaultColumnStyle();
    }

    public void testTestGetSetMargin() {
        baseTestGetSetMargin(new double[]{0.7, 0.7, 0.75, 0.75, 0.3, 0.3});
    }

    public void testExistingHeaderFooter() {
        XSSFWorkbook workbook = XSSFTestDataSamples.openSampleWorkbook("45540_classic_Header.xlsx");
        XSSFOddHeader hdr;
        XSSFOddFooter ftr;

        // Sheet 1 has a header with center and right text
        XSSFSheet s1 = workbook.getSheetAt(0);
        assertNotNull(s1.getHeader());
        assertNotNull(s1.getFooter());
        hdr = (XSSFOddHeader) s1.getHeader();
        ftr = (XSSFOddFooter) s1.getFooter();

        assertEquals("&Ctestdoc&Rtest phrase", hdr.getText());
        assertEquals(null, ftr.getText());

        assertEquals("", hdr.getLeft());
        assertEquals("testdoc", hdr.getCenter());
        assertEquals("test phrase", hdr.getRight());

        assertEquals("", ftr.getLeft());
        assertEquals("", ftr.getCenter());
        assertEquals("", ftr.getRight());

        // Sheet 2 has a footer, but it's empty
        XSSFSheet s2 = workbook.getSheetAt(1);
        assertNotNull(s2.getHeader());
        assertNotNull(s2.getFooter());
        hdr = (XSSFOddHeader) s2.getHeader();
        ftr = (XSSFOddFooter) s2.getFooter();

        assertEquals(null, hdr.getText());
        assertEquals("&L&F", ftr.getText());

        assertEquals("", hdr.getLeft());
        assertEquals("", hdr.getCenter());
        assertEquals("", hdr.getRight());

        assertEquals("&F", ftr.getLeft());
        assertEquals("", ftr.getCenter());
        assertEquals("", ftr.getRight());

        // Save and reload
        XSSFWorkbook wb = XSSFTestDataSamples.writeOutAndReadBack(workbook);

        hdr = (XSSFOddHeader) wb.getSheetAt(0).getHeader();
        ftr = (XSSFOddFooter) wb.getSheetAt(0).getFooter();

        assertEquals("", hdr.getLeft());
        assertEquals("testdoc", hdr.getCenter());
        assertEquals("test phrase", hdr.getRight());

        assertEquals("", ftr.getLeft());
        assertEquals("", ftr.getCenter());
        assertEquals("", ftr.getRight());
    }

    public void testGetAllHeadersFooters() {
        XSSFWorkbook workbook = getTestDataProvider().createWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet 1");
        assertNotNull(sheet.getOddFooter());
        assertNotNull(sheet.getEvenFooter());
        assertNotNull(sheet.getFirstFooter());
        assertNotNull(sheet.getOddHeader());
        assertNotNull(sheet.getEvenHeader());
        assertNotNull(sheet.getFirstHeader());

        assertEquals("", sheet.getOddFooter().getLeft());
        sheet.getOddFooter().setLeft("odd footer left");
        assertEquals("odd footer left", sheet.getOddFooter().getLeft());

        assertEquals("", sheet.getEvenFooter().getLeft());
        sheet.getEvenFooter().setLeft("even footer left");
        assertEquals("even footer left", sheet.getEvenFooter().getLeft());

        assertEquals("", sheet.getFirstFooter().getLeft());
        sheet.getFirstFooter().setLeft("first footer left");
        assertEquals("first footer left", sheet.getFirstFooter().getLeft());

        assertEquals("", sheet.getOddHeader().getLeft());
        sheet.getOddHeader().setLeft("odd header left");
        assertEquals("odd header left", sheet.getOddHeader().getLeft());

        assertEquals("", sheet.getOddHeader().getRight());
        sheet.getOddHeader().setRight("odd header right");
        assertEquals("odd header right", sheet.getOddHeader().getRight());

        assertEquals("", sheet.getOddHeader().getCenter());
        sheet.getOddHeader().setCenter("odd header center");
        assertEquals("odd header center", sheet.getOddHeader().getCenter());

        // Defaults are odd
        assertEquals("odd footer left", sheet.getFooter().getLeft());
        assertEquals("odd header center", sheet.getHeader().getCenter());
    }

    public void testAutoSizeColumn() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet 1");
        sheet.createRow(0).createCell(13).setCellValue("test");

        sheet.autoSizeColumn(13);

        ColumnHelper columnHelper = sheet.getColumnHelper();
        CTCol col = columnHelper.getColumn(13, false);
        assertTrue(col.getBestFit());
    }


    public void testGetCellComment() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        XSSFComment comment = sheet.createComment();
        comment.setAuthor("test C10 author");
        sheet.setCellComment("C10", comment);

        assertNotNull(sheet.getCellComment(9, 2));
        assertEquals("test C10 author", sheet.getCellComment(9, 2).getAuthor());
    }

    public void testSetCellComment() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        XSSFComment comment = sheet.createComment();

        Cell cell = sheet.createRow(0).createCell((short) 0);
        CommentsTable comments = sheet.getCommentsTable();
        CTComments ctComments = comments.getCTComments();

        sheet.setCellComment("A1", comment);
        assertEquals("A1", ctComments.getCommentList().getCommentArray(0).getRef());
        comment.setAuthor("test A1 author");
        assertEquals("test A1 author", comments.getAuthor((int) ctComments.getCommentList().getCommentArray(0).getAuthorId()));
    }

    public void testGetActiveCell() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        sheet.setActiveCell("R5");

        assertEquals("R5", sheet.getActiveCell());

    }

    public void testCreateFreezePane() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();

        sheet.createFreezePane(2, 4);
        assertEquals(2.0, ctWorksheet.getSheetViews().getSheetViewArray(0).getPane().getXSplit(), 0.0);
        assertEquals(STPane.BOTTOM_RIGHT, ctWorksheet.getSheetViews().getSheetViewArray(0).getPane().getActivePane());
        sheet.createFreezePane(3, 6, 10, 10);
        assertEquals(3.0, ctWorksheet.getSheetViews().getSheetViewArray(0).getPane().getXSplit(), 0.0);
        //	assertEquals(10, sheet.getTopRow());
        //	assertEquals(10, sheet.getLeftCol());
        sheet.createSplitPane(4, 8, 12, 12, 1);
        assertEquals(8.0, ctWorksheet.getSheetViews().getSheetViewArray(0).getPane().getYSplit(), 0.0);
        assertEquals(STPane.BOTTOM_RIGHT, ctWorksheet.getSheetViews().getSheetViewArray(0).getPane().getActivePane());
    }

    public void testNewMergedRegionAt() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        CellRangeAddress region = CellRangeAddress.valueOf("B2:D4");
        sheet.addMergedRegion(region);
        assertEquals("B2:D4", sheet.getMergedRegion(0).formatAsString());
        assertEquals(1, sheet.getNumMergedRegions());
    }

    public void testRemoveMergedRegion_lowlevel() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CellRangeAddress region_1 = CellRangeAddress.valueOf("A1:B2");
        CellRangeAddress region_2 = CellRangeAddress.valueOf("C3:D4");
        CellRangeAddress region_3 = CellRangeAddress.valueOf("E5:F6");
        sheet.addMergedRegion(region_1);
        sheet.addMergedRegion(region_2);
        sheet.addMergedRegion(region_3);
        assertEquals("C3:D4", ctWorksheet.getMergeCells().getMergeCellArray(1).getRef());
        assertEquals(3, sheet.getNumMergedRegions());
        sheet.removeMergedRegion(1);
        assertEquals("E5:F6", ctWorksheet.getMergeCells().getMergeCellArray(1).getRef());
        assertEquals(2, sheet.getNumMergedRegions());
        sheet.removeMergedRegion(1);
        sheet.removeMergedRegion(0);
        assertEquals(0, sheet.getNumMergedRegions());
    }

    public void testSetDefaultColumnStyle() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        StylesTable stylesTable = workbook.getStylesSource();
        XSSFFont font = new XSSFFont();
        font.setFontName("Cambria");
        stylesTable.putFont(font);
        CTXf cellStyleXf = CTXf.Factory.newInstance();
        cellStyleXf.setFontId(1);
        cellStyleXf.setFillId(0);
        cellStyleXf.setBorderId(0);
        cellStyleXf.setNumFmtId(0);
        stylesTable.putCellStyleXf(cellStyleXf);
        CTXf cellXf = CTXf.Factory.newInstance();
        cellXf.setXfId(1);
        stylesTable.putCellXf(cellXf);
        XSSFCellStyle cellStyle = new XSSFCellStyle(1, 1, stylesTable);
        assertEquals(1, cellStyle.getFontIndex());

        sheet.setDefaultColumnStyle((short) 3, cellStyle);
        assertEquals(1, ctWorksheet.getColsArray(0).getColArray(0).getStyle());
    }


    public void testGroupUngroupColumn() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        //one level
        sheet.groupColumn((short) 2, (short) 7);
        sheet.groupColumn((short) 10, (short) 11);
        CTCols cols = sheet.getCTWorksheet().getColsArray(0);
        assertEquals(2, cols.sizeOfColArray());
        CTCol[] colArray = cols.getColArray();
        assertNotNull(colArray);
        assertEquals(2 + 1, colArray[0].getMin()); // 1 based
        assertEquals(7 + 1, colArray[0].getMax()); // 1 based
        assertEquals(1, colArray[0].getOutlineLevel());

        //two level
        sheet.groupColumn((short) 1, (short) 2);
        cols = sheet.getCTWorksheet().getColsArray(0);
        assertEquals(4, cols.sizeOfColArray());
        colArray = cols.getColArray();
        assertEquals(2, colArray[1].getOutlineLevel());

        //three level
        sheet.groupColumn((short) 6, (short) 8);
        sheet.groupColumn((short) 2, (short) 3);
        cols = sheet.getCTWorksheet().getColsArray(0);
        assertEquals(7, cols.sizeOfColArray());
        colArray = cols.getColArray();
        assertEquals(3, colArray[1].getOutlineLevel());
        assertEquals(3, sheet.getCTWorksheet().getSheetFormatPr().getOutlineLevelCol());

        sheet.ungroupColumn((short) 8, (short) 10);
        colArray = cols.getColArray();
        //assertEquals(3, colArray[1].getOutlineLevel());

        sheet.ungroupColumn((short) 4, (short) 6);
        sheet.ungroupColumn((short) 2, (short) 2);
        colArray = cols.getColArray();
        assertEquals(4, colArray.length);
        assertEquals(2, sheet.getCTWorksheet().getSheetFormatPr().getOutlineLevelCol());
    }


    public void testGroupUngroupRow() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        //one level
        sheet.groupRow(9, 10);
        assertEquals(2, sheet.getPhysicalNumberOfRows());
        CTRow ctrow = sheet.getRow(9).getCTRow();

        assertNotNull(ctrow);
        assertEquals(10, ctrow.getR());
        assertEquals(1, ctrow.getOutlineLevel());
        assertEquals(1, sheet.getCTWorksheet().getSheetFormatPr().getOutlineLevelRow());

        //two level
        sheet.groupRow(10, 13);
        assertEquals(5, sheet.getPhysicalNumberOfRows());
        ctrow = sheet.getRow(10).getCTRow();
        assertNotNull(ctrow);
        assertEquals(11, ctrow.getR());
        assertEquals(2, ctrow.getOutlineLevel());
        assertEquals(2, sheet.getCTWorksheet().getSheetFormatPr().getOutlineLevelRow());


        sheet.ungroupRow(8, 10);
        assertEquals(4, sheet.getPhysicalNumberOfRows());
        assertEquals(1, sheet.getCTWorksheet().getSheetFormatPr().getOutlineLevelRow());

        sheet.ungroupRow(10, 10);
        assertEquals(3, sheet.getPhysicalNumberOfRows());

        assertEquals(1, sheet.getCTWorksheet().getSheetFormatPr().getOutlineLevelRow());
    }

    public void testSetZoom() {
        XSSFWorkbook workBook = new XSSFWorkbook();
        XSSFSheet sheet1 = workBook.createSheet("new sheet");
        sheet1.setZoom(3, 4);   // 75 percent magnification
        long zoom = sheet1.getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale();
        assertEquals(zoom, 75);

        sheet1.setZoom(200);
        zoom = sheet1.getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale();
        assertEquals(zoom, 200);

        try {
            sheet1.setZoom(500);
            fail("Expecting exception");
        } catch (IllegalArgumentException e) {
            assertEquals("Valid scale values range from 10 to 400", e.getMessage());
        }
    }

    /**
     * TODO - while this is internally consistent, I'm not
     *  completely clear in all cases what it's supposed to
     *  be doing... Someone who understands the goals a little
     *  better should really review this!
     */
    public void testSetColumnGroupCollapsed(){
    	Workbook wb = new XSSFWorkbook();
    	XSSFSheet sheet1 =(XSSFSheet) wb.createSheet();

    	CTCols cols=sheet1.getCTWorksheet().getColsArray(0);
    	assertEquals(0,cols.sizeOfColArray());

    	sheet1.groupColumn( (short)4, (short)7 );
    	sheet1.groupColumn( (short)9, (short)12 );

    	assertEquals(2,cols.sizeOfColArray());

    	assertEquals(false,cols.getColArray(0).isSetHidden());
    	assertEquals(true, cols.getColArray(0).isSetCollapsed());
    	assertEquals(5, cols.getColArray(0).getMin()); // 1 based
    	assertEquals(8, cols.getColArray(0).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(1).isSetHidden());
    	assertEquals(true, cols.getColArray(1).isSetCollapsed());
    	assertEquals(10, cols.getColArray(1).getMin()); // 1 based
    	assertEquals(13, cols.getColArray(1).getMax()); // 1 based

    	sheet1.groupColumn( (short)10, (short)11 );
    	assertEquals(4,cols.sizeOfColArray());

    	assertEquals(false,cols.getColArray(0).isSetHidden());
    	assertEquals(true, cols.getColArray(0).isSetCollapsed());
    	assertEquals(5, cols.getColArray(0).getMin()); // 1 based
    	assertEquals(8, cols.getColArray(0).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(1).isSetHidden());
    	assertEquals(true, cols.getColArray(1).isSetCollapsed());
    	assertEquals(10, cols.getColArray(1).getMin()); // 1 based
    	assertEquals(10, cols.getColArray(1).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(2).isSetHidden());
    	assertEquals(true, cols.getColArray(2).isSetCollapsed());
    	assertEquals(11, cols.getColArray(2).getMin()); // 1 based
    	assertEquals(12, cols.getColArray(2).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(3).isSetHidden());
    	assertEquals(true, cols.getColArray(3).isSetCollapsed());
    	assertEquals(13, cols.getColArray(3).getMin()); // 1 based
    	assertEquals(13, cols.getColArray(3).getMax()); // 1 based

    	// collapse columns - 1
    	sheet1.setColumnGroupCollapsed( (short)5, true );
    	assertEquals(5,cols.sizeOfColArray());

    	assertEquals(true, cols.getColArray(0).isSetHidden());
    	assertEquals(true, cols.getColArray(0).isSetCollapsed());
    	assertEquals(5, cols.getColArray(0).getMin()); // 1 based
    	assertEquals(8, cols.getColArray(0).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(1).isSetHidden());
    	assertEquals(true, cols.getColArray(1).isSetCollapsed());
    	assertEquals(9, cols.getColArray(1).getMin()); // 1 based
    	assertEquals(9, cols.getColArray(1).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(2).isSetHidden());
    	assertEquals(true, cols.getColArray(2).isSetCollapsed());
    	assertEquals(10, cols.getColArray(2).getMin()); // 1 based
    	assertEquals(10, cols.getColArray(2).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(3).isSetHidden());
    	assertEquals(true, cols.getColArray(3).isSetCollapsed());
    	assertEquals(11, cols.getColArray(3).getMin()); // 1 based
    	assertEquals(12, cols.getColArray(3).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(4).isSetHidden());
    	assertEquals(true, cols.getColArray(4).isSetCollapsed());
    	assertEquals(13, cols.getColArray(4).getMin()); // 1 based
    	assertEquals(13, cols.getColArray(4).getMax()); // 1 based


    	// expand columns - 1
    	sheet1.setColumnGroupCollapsed( (short)5, false );

    	assertEquals(false,cols.getColArray(0).isSetHidden());
    	assertEquals(true, cols.getColArray(0).isSetCollapsed());
    	assertEquals(5, cols.getColArray(0).getMin()); // 1 based
    	assertEquals(8, cols.getColArray(0).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(1).isSetHidden());
    	assertEquals(false,cols.getColArray(1).isSetCollapsed());
    	assertEquals(9, cols.getColArray(1).getMin()); // 1 based
    	assertEquals(9, cols.getColArray(1).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(2).isSetHidden());
    	assertEquals(true, cols.getColArray(2).isSetCollapsed());
    	assertEquals(10, cols.getColArray(2).getMin()); // 1 based
    	assertEquals(10, cols.getColArray(2).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(3).isSetHidden());
    	assertEquals(true, cols.getColArray(3).isSetCollapsed());
    	assertEquals(11, cols.getColArray(3).getMin()); // 1 based
    	assertEquals(12, cols.getColArray(3).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(4).isSetHidden());
    	assertEquals(true, cols.getColArray(4).isSetCollapsed());
    	assertEquals(13, cols.getColArray(4).getMin()); // 1 based
    	assertEquals(13, cols.getColArray(4).getMax()); // 1 based


    	//collapse - 2
    	sheet1.setColumnGroupCollapsed( (short)9, true );
    	assertEquals(6,cols.sizeOfColArray());
    	assertEquals(false,cols.getColArray(0).isSetHidden());
    	assertEquals(true, cols.getColArray(0).isSetCollapsed());
    	assertEquals(5, cols.getColArray(0).getMin()); // 1 based
    	assertEquals(8, cols.getColArray(0).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(1).isSetHidden());
    	assertEquals(false,cols.getColArray(1).isSetCollapsed());
    	assertEquals(9, cols.getColArray(1).getMin()); // 1 based
    	assertEquals(9, cols.getColArray(1).getMax()); // 1 based
    	assertEquals(true, cols.getColArray(2).isSetHidden());
    	assertEquals(true, cols.getColArray(2).isSetCollapsed());
    	assertEquals(10, cols.getColArray(2).getMin()); // 1 based
    	assertEquals(10, cols.getColArray(2).getMax()); // 1 based
    	assertEquals(true, cols.getColArray(3).isSetHidden());
    	assertEquals(true, cols.getColArray(3).isSetCollapsed());
    	assertEquals(11, cols.getColArray(3).getMin()); // 1 based
    	assertEquals(12, cols.getColArray(3).getMax()); // 1 based
    	assertEquals(true, cols.getColArray(4).isSetHidden());
    	assertEquals(true, cols.getColArray(4).isSetCollapsed());
    	assertEquals(13, cols.getColArray(4).getMin()); // 1 based
    	assertEquals(13, cols.getColArray(4).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(5).isSetHidden());
    	assertEquals(true, cols.getColArray(5).isSetCollapsed());
    	assertEquals(14, cols.getColArray(5).getMin()); // 1 based
    	assertEquals(14, cols.getColArray(5).getMax()); // 1 based


    	//expand - 2
    	sheet1.setColumnGroupCollapsed( (short)9, false );
    	assertEquals(6,cols.sizeOfColArray());
    	assertEquals(14,cols.getColArray(5).getMin());

    	//outline level 2: the line under ==> collapsed==True
    	assertEquals(2,cols.getColArray(3).getOutlineLevel());
    	assertEquals(true,cols.getColArray(4).isSetCollapsed());

    	assertEquals(false,cols.getColArray(0).isSetHidden());
    	assertEquals(true, cols.getColArray(0).isSetCollapsed());
    	assertEquals(5, cols.getColArray(0).getMin()); // 1 based
    	assertEquals(8, cols.getColArray(0).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(1).isSetHidden());
    	assertEquals(false,cols.getColArray(1).isSetCollapsed());
    	assertEquals(9, cols.getColArray(1).getMin()); // 1 based
    	assertEquals(9, cols.getColArray(1).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(2).isSetHidden());
    	assertEquals(true, cols.getColArray(2).isSetCollapsed());
    	assertEquals(10, cols.getColArray(2).getMin()); // 1 based
    	assertEquals(10, cols.getColArray(2).getMax()); // 1 based
    	assertEquals(true, cols.getColArray(3).isSetHidden());
    	assertEquals(true, cols.getColArray(3).isSetCollapsed());
    	assertEquals(11, cols.getColArray(3).getMin()); // 1 based
    	assertEquals(12, cols.getColArray(3).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(4).isSetHidden());
    	assertEquals(true, cols.getColArray(4).isSetCollapsed());
    	assertEquals(13, cols.getColArray(4).getMin()); // 1 based
    	assertEquals(13, cols.getColArray(4).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(5).isSetHidden());
    	assertEquals(false,cols.getColArray(5).isSetCollapsed());
    	assertEquals(14, cols.getColArray(5).getMin()); // 1 based
    	assertEquals(14, cols.getColArray(5).getMax()); // 1 based

    	//DOCUMENTARE MEGLIO IL DISCORSO DEL LIVELLO
    	//collapse - 3
    	sheet1.setColumnGroupCollapsed( (short)10, true );
    	assertEquals(6,cols.sizeOfColArray());
    	assertEquals(false,cols.getColArray(0).isSetHidden());
    	assertEquals(true, cols.getColArray(0).isSetCollapsed());
    	assertEquals(5, cols.getColArray(0).getMin()); // 1 based
    	assertEquals(8, cols.getColArray(0).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(1).isSetHidden());
    	assertEquals(false,cols.getColArray(1).isSetCollapsed());
    	assertEquals(9, cols.getColArray(1).getMin()); // 1 based
    	assertEquals(9, cols.getColArray(1).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(2).isSetHidden());
    	assertEquals(true, cols.getColArray(2).isSetCollapsed());
    	assertEquals(10, cols.getColArray(2).getMin()); // 1 based
    	assertEquals(10, cols.getColArray(2).getMax()); // 1 based
    	assertEquals(true, cols.getColArray(3).isSetHidden());
    	assertEquals(true, cols.getColArray(3).isSetCollapsed());
    	assertEquals(11, cols.getColArray(3).getMin()); // 1 based
    	assertEquals(12, cols.getColArray(3).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(4).isSetHidden());
    	assertEquals(true, cols.getColArray(4).isSetCollapsed());
    	assertEquals(13, cols.getColArray(4).getMin()); // 1 based
    	assertEquals(13, cols.getColArray(4).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(5).isSetHidden());
    	assertEquals(false,cols.getColArray(5).isSetCollapsed());
    	assertEquals(14, cols.getColArray(5).getMin()); // 1 based
    	assertEquals(14, cols.getColArray(5).getMax()); // 1 based


    	//expand - 3
    	sheet1.setColumnGroupCollapsed( (short)10, false );
    	assertEquals(6,cols.sizeOfColArray());
    	assertEquals(false,cols.getColArray(0).getHidden());
    	assertEquals(false,cols.getColArray(5).getHidden());
    	assertEquals(false,cols.getColArray(4).isSetCollapsed());

//  	write out and give back
    	// Save and re-load
    	wb = XSSFTestDataSamples.writeOutAndReadBack(wb);
    	sheet1 = (XSSFSheet)wb.getSheetAt(0);
    	assertEquals(6,cols.sizeOfColArray());

    	assertEquals(false,cols.getColArray(0).isSetHidden());
    	assertEquals(true, cols.getColArray(0).isSetCollapsed());
    	assertEquals(5, cols.getColArray(0).getMin()); // 1 based
    	assertEquals(8, cols.getColArray(0).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(1).isSetHidden());
    	assertEquals(false,cols.getColArray(1).isSetCollapsed());
    	assertEquals(9, cols.getColArray(1).getMin()); // 1 based
    	assertEquals(9, cols.getColArray(1).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(2).isSetHidden());
    	assertEquals(true, cols.getColArray(2).isSetCollapsed());
    	assertEquals(10, cols.getColArray(2).getMin()); // 1 based
    	assertEquals(10, cols.getColArray(2).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(3).isSetHidden());
    	assertEquals(true, cols.getColArray(3).isSetCollapsed());
    	assertEquals(11, cols.getColArray(3).getMin()); // 1 based
    	assertEquals(12, cols.getColArray(3).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(4).isSetHidden());
    	assertEquals(false,cols.getColArray(4).isSetCollapsed());
    	assertEquals(13, cols.getColArray(4).getMin()); // 1 based
    	assertEquals(13, cols.getColArray(4).getMax()); // 1 based
    	assertEquals(false,cols.getColArray(5).isSetHidden());
    	assertEquals(false,cols.getColArray(5).isSetCollapsed());
    	assertEquals(14, cols.getColArray(5).getMin()); // 1 based
    	assertEquals(14, cols.getColArray(5).getMax()); // 1 based
    }

    /**
     * TODO - while this is internally consistent, I'm not
     *  completely clear in all cases what it's supposed to
     *  be doing... Someone who understands the goals a little
     *  better should really review this!
     */
    public void testSetRowGroupCollapsed(){
    	Workbook wb = new XSSFWorkbook();
    	XSSFSheet sheet1 = (XSSFSheet)wb.createSheet();

    	sheet1.groupRow( 5, 14 );
    	sheet1.groupRow( 7, 14 );
    	sheet1.groupRow( 16, 19 );

    	assertEquals(14,sheet1.getPhysicalNumberOfRows());
    	assertEquals(false,sheet1.getRow(6).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(6).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(7).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(7).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(9).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(9).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(14).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(14).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(16).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(16).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(18).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(18).getCTRow().isSetHidden());

    	//collapsed
    	sheet1.setRowGroupCollapsed( 7, true );

    	assertEquals(false,sheet1.getRow(6).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(6).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(7).getCTRow().isSetCollapsed());
    	assertEquals(true, sheet1.getRow(7).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(9).getCTRow().isSetCollapsed());
    	assertEquals(true, sheet1.getRow(9).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(14).getCTRow().isSetCollapsed());
    	assertEquals(true, sheet1.getRow(14).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(16).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(16).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(18).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(18).getCTRow().isSetHidden());

    	//expanded
    	sheet1.setRowGroupCollapsed( 7, false );

    	assertEquals(false,sheet1.getRow(6).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(6).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(7).getCTRow().isSetCollapsed());
    	assertEquals(true, sheet1.getRow(7).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(9).getCTRow().isSetCollapsed());
    	assertEquals(true, sheet1.getRow(9).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(14).getCTRow().isSetCollapsed());
    	assertEquals(true, sheet1.getRow(14).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(16).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(16).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(18).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(18).getCTRow().isSetHidden());


    	// Save and re-load
    	wb = XSSFTestDataSamples.writeOutAndReadBack(wb);
    	sheet1 = (XSSFSheet)wb.getSheetAt(0);

    	assertEquals(false,sheet1.getRow(6).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(6).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(7).getCTRow().isSetCollapsed());
    	assertEquals(true, sheet1.getRow(7).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(9).getCTRow().isSetCollapsed());
    	assertEquals(true, sheet1.getRow(9).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(14).getCTRow().isSetCollapsed());
    	assertEquals(true, sheet1.getRow(14).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(16).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(16).getCTRow().isSetHidden());
    	assertEquals(false,sheet1.getRow(18).getCTRow().isSetCollapsed());
    	assertEquals(false,sheet1.getRow(18).getCTRow().isSetHidden());
    }

    /**
     * Get / Set column width and check the actual values of the underlying XML beans
     */
    public void testColumnWidth_lowlevel() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet 1");
        sheet.setColumnWidth(1, 22 * 256);
        assertEquals(22 * 256, sheet.getColumnWidth(1));

        // Now check the low level stuff, and check that's all
        //  been set correctly
        XSSFSheet xs = sheet;
        CTWorksheet cts = xs.getCTWorksheet();

        CTCols[] cols_s = cts.getColsArray();
        assertEquals(1, cols_s.length);
        CTCols cols = cols_s[0];
        assertEquals(1, cols.sizeOfColArray());
        CTCol col = cols.getColArray(0);

        // XML is 1 based, POI is 0 based
        assertEquals(2, col.getMin());
        assertEquals(2, col.getMax());
        assertEquals(22.0, col.getWidth(), 0.0);
        assertTrue(col.getCustomWidth());

        // Now set another
        sheet.setColumnWidth(3, 33 * 256);

        cols_s = cts.getColsArray();
        assertEquals(1, cols_s.length);
        cols = cols_s[0];
        assertEquals(2, cols.sizeOfColArray());

        col = cols.getColArray(0);
        assertEquals(2, col.getMin()); // POI 1
        assertEquals(2, col.getMax());
        assertEquals(22.0, col.getWidth(), 0.0);
        assertTrue(col.getCustomWidth());

        col = cols.getColArray(1);
        assertEquals(4, col.getMin()); // POI 3
        assertEquals(4, col.getMax());
        assertEquals(33.0, col.getWidth(), 0.0);
        assertTrue(col.getCustomWidth());
    }
}
