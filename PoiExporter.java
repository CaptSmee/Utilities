import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import javax.faces.context.ExternalContext;
import javax.faces.context.FacesContext;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class contains static methods for exporting data using apache poi
 * @author DCConway
 */
public class PoiExporter {
    
    private PoiExporter(){}
    
    /**
     * This method creates an excel spreadsheet in .xlsx format and streams the 
     * resulting spreadsheet back to the client. The method uses reflection to 
     * match class member field names to the string values in the fieldsUsed parameter.
     * The column headers are created based on the String values in the headers
     * parameter. 
     * The data parameter is a collection of Class<?> whose type matches the "theClass"
     * parameter which is also a Class<?>. Class member fields used in creating the 
     * spreadsheet must have a value at the time of class construction or the value 
     * will not show up in the spreadsheet.
     * @param filename - Name of the spreadsheet to be created.
     * @param headers - List of Column headers for the spreadsheet.
     * @param data - Collection of type <?>.
     * @param theClass - Class<?>
     * @param fieldsUsed - List of Class<?> member fields to be used.
     * @param context - FacesContext.getCurrentInstance.
     * @throws NoSuchMethodException
     * @throws ClassNotFoundException
     * @throws IllegalAccessException
     * @throws IllegalArgumentException
     * @throws InvocationTargetException
     * @throws FileNotFoundException
     * @throws IOException 
     */
    public static void exportSpecificDataToExcel(String filename,ArrayList<String> headers, ArrayList<?> data,Class<?> theClass,ArrayList<String> fieldsUsed,FacesContext context) 
        throws NoSuchMethodException, ClassNotFoundException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, FileNotFoundException, IOException{
        
        // create workbook
        Workbook wb = new XSSFWorkbook();
        // creation helper
        CreationHelper helper = wb.getCreationHelper();
        CellStyle wrapStyle = wb.createCellStyle();
        wrapStyle.setWrapText(true);
        // create sheet
        Sheet sheet1 = wb.createSheet("Export");
        // create header row and cells
        Row headerRow = sheet1.createRow(0);
        for(int i = 0; i < headers.size(); i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }
        int largeCell = -1;
        // fields
        Field[] fields = theClass.getDeclaredFields();
        // loop over the data list
        for(int i=0; i < data.size(); i++){
            Row row = sheet1.createRow(i+1);
            row.setRowStyle(wrapStyle);
            for(int j = 0; j < fieldsUsed.size(); j++){
                for(Field f : fields){
                    if(f.getName().equals(fieldsUsed.get(j))){
                        f.setAccessible(true);
                        Cell cell = row.createCell(j);
                        if(f.getType().getName().equals("java.util.Date")){
                            CellStyle cellStyle = wb.createCellStyle();
                            cellStyle.setDataFormat(helper.createDataFormat().getFormat("m/dd/yyyy"));
                            cell.setCellValue((Date)f.get(data.get(i)));
                            cell.setCellStyle(cellStyle);
                        }else{
                            if(f.get(data.get(i)) != null && ((String)f.get(data.get(i))).length() > 100){
                                largeCell = j;
                                sheet1.setColumnWidth(j, 18000);
                                cell.setCellValue((String)f.get(data.get(i)));
                                cell.setCellStyle(wrapStyle);
                            }else{
                                cell.setCellValue((String)f.get(data.get(i)));
                            }
                        }
                    }
                }
            }

        }
        // set the columns to auto size
        Sheet sheet = wb.getSheetAt(0);
        for(int i = 0; i < fieldsUsed.size(); i++){
            if(i == largeCell){
                continue;
            }
            sheet.autoSizeColumn(i);
        }
        // write to stream
        ExternalContext ec = FacesContext.getCurrentInstance().getExternalContext();
        HttpServletResponse response = (HttpServletResponse) ec.getResponse();
        OutputStream out = response.getOutputStream();
        response.reset();
        response.setHeader("Content-Type", "application/excel");
        response.setHeader("Content-Disposition", "inline; filename=\"" + filename + "\"");
        
        wb.write(out);
        wb.close();
        FacesContext.getCurrentInstance().responseComplete();
            
    }
    
    /**
     * This method creates an excel spreadsheet in .xlsx format and streams the 
     * resulting spreadsheet back to the client. 
     * The keys in the HasMap must be named the same as the headers to make
     * sure that the correct data goes into the column cell.
     * @param filename - name of the excel file to be created.
     * @param headers - List<String> of column header names.
     * @param data - ArrayList<HashMap<String,String>> of the data to be inserted into the columns
     * @param context - FacesContext.getCurrentInstance.
     */
    public static void exportUsingList(String filename,ArrayList<String> headers, ArrayList<HashMap<String,String>> data,FacesContext context) throws IOException{
        
        // create workbook
        Workbook wb = new XSSFWorkbook();
        CellStyle wrapStyle = wb.createCellStyle();
        wrapStyle.setWrapText(true);
        // create sheet
        Sheet sheet1 = wb.createSheet("Export");
        // create header row and cells
        Row headerRow = sheet1.createRow(0);
        for(int i = 0; i < headers.size(); i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }
        
        // for each hashmap in the list
        int largeCell = -1;
        for(int i = 0; i < data.size(); i++){
            // create a row
            Row row = sheet1.createRow(i+1);
            row.setRowStyle(wrapStyle);
            // loop on the headers list, pull the value by the key and create a cell
            for(int j = 0; j < headers.size(); j++){
                Cell cell = row.createCell(j);
                HashMap<String,String> map = data.get(i);
                if(map.get(headers.get(j)) != null && map.get(headers.get(j)).length() > 100){
                    largeCell = j;
                    sheet1.setColumnWidth(j, 18000);
                    cell.setCellValue(map.get(headers.get(j)));
                    cell.setCellStyle(wrapStyle);
                }else{
                    cell.setCellValue(map.get(headers.get(j)));
                }
            }
        }
        
        Sheet sheet = wb.getSheetAt(0);
        for(int i = 0; i < data.size(); i++){
            if(i == largeCell){
                continue;
            }
            sheet.autoSizeColumn(i);
        }
        
        // write to stream
        ExternalContext ec = FacesContext.getCurrentInstance().getExternalContext();
        HttpServletResponse response = (HttpServletResponse) ec.getResponse();
        OutputStream out = response.getOutputStream();
        response.setHeader("Content-Type", "application/excel");
        response.setHeader("Content-Disposition", "inline; filename=\"" + filename + "\"");

        wb.write(out);
        wb.close();
        FacesContext.getCurrentInstance().responseComplete();
    }
    
    /**
     * Method HAS NOT BEEN TESTED. USE AT YOUR OWN RISK.
     * This method creates an excel spreadsheet in .xlsx format and streams the 
     * resulting spreadsheet back to the client. The method uses reflection to 
     * pull all of the classes members fields, attempts to determine their type
     * and cast them to appropriate type that can be inserted as a cell value
     * The column headers are created based on the String values in the headers
     * parameter. 
     * The data parameter is a collection of Class<?> whose type matches the "theClass"
     * parameter which is also a Class<?>. 
     * Class member fields used in creating the 
     * spreadsheet must have a value at the time of class construction or the value 
     * will not show up in the spreadsheet.
     * @param filename
     * @param headers
     * @param data
     * @param theClass
     * @param context
     * @throws NoSuchMethodException
     * @throws ClassNotFoundException
     * @throws IllegalAccessException
     * @throws IllegalArgumentException
     * @throws InvocationTargetException
     * @throws FileNotFoundException
     * @throws IOException 
     */
    public static void exportToExcel(String filename,ArrayList<String> headers, ArrayList<?> data,Class<?> theClass,FacesContext context) 
        throws NoSuchMethodException, ClassNotFoundException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, FileNotFoundException, IOException{
        
        // create workbook
        Workbook wb = new XSSFWorkbook();
        // creation helper
        CreationHelper helper = wb.getCreationHelper();
        CellStyle wrapStyle = wb.createCellStyle();
        wrapStyle.setWrapText(true);
        // create sheet
        Sheet sheet1 = wb.createSheet("Export");
        // create header row and cells
        Row headerRow = sheet1.createRow(0);
        for(int i = 0; i < headers.size(); i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }
        int largeCell = 0;
        // fields
        Field[] fields = theClass.getDeclaredFields();
        for(int i=0; i < data.size(); i++){
            Row row = sheet1.createRow(i+1);
            row.setRowStyle(wrapStyle);
            for(int j = 0; j < fields.length; j++){
                fields[j].setAccessible(true);
                Cell cell = row.createCell(j);
                if(fields[j].getType().getName().equals("java.util.Date")){
                    CellStyle cellStyle = wb.createCellStyle();
                    cellStyle.setDataFormat(helper.createDataFormat().getFormat("m/dd/yyyy"));
                    cell.setCellValue((Date)fields[j].get(data.get(i)));
                    cell.setCellStyle(cellStyle);
                }else if(fields[j].getType().getName().equals("java.lang.String")){
                        if(fields[j].get(data.get(i)) != null && ((String)fields[j].get(data.get(i))).length() > 50){
                            largeCell = j;
                            System.out.println("content length > 50");
                            sheet1.setColumnWidth(j, 18000);
                            cell.setCellValue((String)fields[j].get(data.get(i)));
                            cell.setCellStyle(wrapStyle);
                        }else{
                            cell.setCellValue((String)fields[j].get(data.get(i)));
                        }
                }else if(fields[j].getType().getName().equals("java.lang.long")){
                    cell.setCellValue((long)fields[j].get(data.get(i)));
                }else if(fields[j].getType().getName().equals("java.lang.integer")){
                    cell.setCellValue((int)fields[j].get(data.get(i)));
                }else if(fields[j].getType().getName().equals("java.math.BigDecimal")){
                    cell.setCellValue(((BigDecimal)fields[j].get(data.get(i))).floatValue());
                }else if(fields[j].getType().getName().equals("java.math.BigInteger")){
                    cell.setCellValue(((BigInteger)fields[j].get(data.get(i))).intValue());
                }
            }
        }
        // set the columns to auto size
        Sheet sheet = wb.getSheetAt(0);
        for(int i = 0; i < fields.length; i++){
            if(i == largeCell){
                continue;
            }
            sheet.autoSizeColumn(i);
        }
        // write to stream
        ExternalContext ec = FacesContext.getCurrentInstance().getExternalContext();
        HttpServletResponse response = (HttpServletResponse) ec.getResponse();
        OutputStream out = response.getOutputStream();
        response.setHeader("Content-Type", "application/excel");
        response.setHeader("Content-Disposition", "inline; filename=\"" + filename + "\"");

        wb.write(out);
        wb.close();
        FacesContext.getCurrentInstance().responseComplete();
    }
}
