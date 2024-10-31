package org.example;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelSplitter {

    private final String fileName;
    private final int maxRows;


    /**
     * Constructor
     * @param fileNamePath
     * @param maxRows
     */
    public ExcelSplitter(String fileNamePath, int maxRows){
        this.fileName = fileNamePath;
        this.maxRows = maxRows;
        try {
            OPCPackage pkg = OPCPackage.open(fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet.getPhysicalNumberOfRows() > 1000){
                List<SXSSFWorkbook> wbs = splitWorkbook(workbook);
                writeWorkBooks(wbs);
            }
            pkg.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public List<SXSSFWorkbook> splitWorkbook (XSSFWorkbook workbook) {
        List<SXSSFWorkbook> workbooks = new ArrayList<SXSSFWorkbook>();
        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFSheet sh = wb.createSheet();

        SXSSFRow newrow;
        SXSSFCell newcell;

        int rowCount = 0;
        int colCount = 0;

        XSSFSheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            newrow = sh.createRow(rowCount++);

            if (rowCount == maxRows) {
                workbooks.add(wb);
                wb = new SXSSFWorkbook();
                sh = wb.createSheet();
                rowCount = 0;

            }
            for (Cell cell : row) {
                newcell = newrow.createCell(colCount++);
                newcell = setValue(newcell, cell);

                CellStyle newCellStyle = wb.createCellStyle();
                newCellStyle.cloneStyleFrom(cell.getCellStyle());
                newcell.setCellStyle(newCellStyle);
            }
            colCount = 0;
        }
        if (wb.getSheetAt(0).getPhysicalNumberOfRows() > 0) {
            workbooks.add(wb);
        }
        return workbooks;
    }

    public SXSSFCell setValue (SXSSFCell newCell, Cell cell){
        switch (cell.getCellType()){
            case NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)){
                    newCell.setCellValue(cell.getDateCellValue());
                } else {
                    newCell.setCellValue(cell.getNumericCellValue());
                }
                break;
            case STRING:
                newCell.setCellValue(cell.getStringCellValue());
                break;
            case FORMULA:
                newCell.setCellValue(cell.getCellFormula());
                break;
            case BOOLEAN:
                newCell.setCellValue(cell.getBooleanCellValue());
                break;
            default:
                System.out.println("Could not determine cell type");
                break;
        }
        return newCell;
    }

    private void writeWorkBooks (List<SXSSFWorkbook> wbs){
        FileOutputStream out;
        try {
            for (int i = 0; i < wbs.size(); i++) {
                String newFileName = extractFileName(fileName);
                out = new FileOutputStream(new File(newFileName + "_" + i + ".xlsx"));
                wbs.get(i).write(out);
                out.close();
            }
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private String extractFileName(String filepath){
        if (filepath != null && filepath.contains("/")){
            return filepath.substring(filepath.lastIndexOf('/')+1);
        } else if (filepath != null && filepath.contains("//")) {
            return filepath.substring(filepath.lastIndexOf("//")+1);
        }
        return filepath;
    }








}
