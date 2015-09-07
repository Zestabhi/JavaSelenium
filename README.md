

 import java.io.File;
 import java.io.IOException;
 import java.util.ArrayList;
 import java.util.Date;
 import java.util.List;
 
 import jxl.Cell;
 import jxl.DateCell;
 import jxl.LabelCell;
 import jxl.NumberCell;
 import jxl.Sheet;
 import jxl.Workbook;
 import jxl.read.biff.BiffException;
 
public class ReadExcel {
 
  private String inputFile;
 
  public void setInputFile(String inputFile) {
     this.inputFile = inputFile;
   }
 
  public void read() throws IOException  {
     File inputWorkbook = new File(inputFile);
     Workbook w = null;
     List<String> nameRowList = null;
     List<Double> marksRowList = null;
     List<Date> dateRowList = null;
     LabelCell lc;//for cell having String data
     NumberCell nc;//for cell having Numeric data
     DateCell dc;//for cell having date
     
     try {
       w = Workbook.getWorkbook(inputWorkbook);
       // Get the first sheet
       Sheet sheet = w.getSheet(0);
 
      // Loop over columns  
       for (int j = 0; j < sheet.getColumns(); j++) {
       Cell cell = sheet.getCell(j, 0);
       System.out.println("\n"+"Column Name => "+cell.getContents());//To print column values
       
       if(cell.getContents().equalsIgnoreCase("Name")){
        nameRowList = new ArrayList<String>();
        // Loop over rows
        for (int i = 0; i < sheet.getRows()-1; i++) {
         Cell styleColorVal = sheet.getCell(j, i+1);
         lc = (LabelCell)styleColorVal;
         nameRowList.add(i, lc.getString());
         System.out.println(cell.getContents()+(i+1)+" :: "+nameRowList.get(i).toString()); 
         //To print the row values for Name column
        }        
       }
       
       if(cell.getContents().equalsIgnoreCase("Marks")){
        marksRowList = new ArrayList<Double>();
        // Loop over rows
        for (int i = 0; i < sheet.getRows()-1; i++) {
         Cell styleColorVal = sheet.getCell(j, i+1);
         nc = (NumberCell)styleColorVal;
         marksRowList.add(i, nc.getValue());
         System.out.println(cell.getContents()+(i+1)+" :: "+marksRowList.get(i).toString()); 
         //To print the row values for Marks column
        }        
       }
       
       if(cell.getContents().equalsIgnoreCase("Date")){
        dateRowList = new ArrayList<Date>();
        // Loop over rows
        for (int i = 0; i < sheet.getRows()-1; i++) {
         Cell styleColorVal = sheet.getCell(j, i+1);
         dc = (DateCell)styleColorVal;
         dateRowList.add(i, dc.getDate());
         System.out.println(cell.getContents()+(i+1)+" :: "+dateRowList.get(i).toString()); 
         //To print the row values for Date column
        }        
       }
             
       
       }
     } catch (BiffException e) {
      e.printStackTrace();
     } catch (Exception ex){
      ex.printStackTrace();
     }finally{
      w.close();
     }
   }
 
  public static void main(String[] args) throws IOException {
     ReadExcel rd = new ReadExcel();
     rd.setInputFile("D:/Example.xls");
     rd.read();
   } 
 } 
 
