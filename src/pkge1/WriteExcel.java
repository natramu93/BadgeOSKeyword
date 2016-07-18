package pkge1;

import java.io.File;
import java.io.IOException;

import jxl.Workbook; 
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class WriteExcel { 
 public static final String NIL = "";
 public static final String RESULTEXCEL = "E:\\excel files\\Autoresult.xls"; 
 public static final String IOEXP = "IO Exception is thrown";
 public static final String WRITEEXP = "Write Exception is thrown";

 public void writeResult(String... args) {

  String[] strLbl = new String[8];
  int len = args.length;
  if (len == 4) {
   for (int i = 0; i < len; i++) {
    strLbl[i] = args[i];
   }
   for (int j = 4; j < 8; j++)
    strLbl[j] = NIL;
  } else if (len == 8) {
   for (int i = 0; i < len; i++) {
    strLbl[i] = args[i];
   }
  }

   WriteExcel we = new WriteExcel();
   Workbook wb = null;
   WritableWorkbook workbook = null;
   try {
    wb = Workbook.getWorkbook(new File(strLbl[0]));
    workbook = Workbook.createWorkbook(new File(
      strLbl[0]), wb);
   } catch (BiffException | IOException e) {
    // TODO Auto-generated catch block
    e.printStackTrace();
   }
   
   // Cellformatting
   // Format with bold	
   
   try {
    WritableCellFormat cellFormat = null;
    WritableFont cellFont = new WritableFont(WritableFont.TIMES, 12);
    cellFont.setBoldStyle(WritableFont.BOLD);
  
   WritableCellFormat cellFormat1 = new WritableCellFormat(cellFont);
    cellFormat1.setBorder(jxl.format.Border.ALL,
      jxl.format.BorderLineStyle.THIN, jxl.format.Colour.BLACK);
   // format without bold
   WritableFont cellFont1 = new WritableFont(WritableFont.TIMES, 12);
   WritableCellFormat cellFormat11 = new WritableCellFormat(cellFont1);
    cellFormat11.setBorder(jxl.format.Border.ALL,
      jxl.format.BorderLineStyle.THIN, jxl.format.Colour.BLACK);

   if (len == 4) {
    cellFormat = cellFormat1;
   } else if (len == 8) {
    cellFormat = cellFormat11;
   }

   WritableFont cellFont2 = new WritableFont(WritableFont.TIMES, 12);
    cellFont2.setColour(Colour.GREEN);
   WritableCellFormat cellFormat2 = new WritableCellFormat(cellFont2);
    cellFormat2.setBorder(jxl.format.Border.ALL,
      jxl.format.BorderLineStyle.THIN, jxl.format.Colour.BLACK);
   WritableFont cellFont3 = new WritableFont(WritableFont.TIMES, 12);
    cellFont3.setColour(Colour.RED);
   WritableCellFormat cellFormat3 = new WritableCellFormat(cellFont3);
    cellFormat3.setBorder(jxl.format.Border.ALL,
      jxl.format.BorderLineStyle.THIN, jxl.format.Colour.BLACK);
   WritableSheet sheet = workbook.getSheet(0);
   int rows = sheet.getRows();
   //  System.out.println(rows);
   Label label;
   for (int n = 1; n < 8; n++) {

    if (n == 7) {
     if (strLbl[n].equalsIgnoreCase("Pass")) {
      label = new Label(n, rows, strLbl[n], cellFormat2);
     } else if (strLbl[n].equalsIgnoreCase("Fail")) {
      label = new Label(n, rows, strLbl[n], cellFormat3);
     } else {
      label = new Label(n, rows, strLbl[n], cellFormat);
     }
    } else {
     label = new Label(n, rows, strLbl[n], cellFormat);
    }
     sheet.addCell(label);
    } 
   workbook.write();
   workbook.close();
   }catch (WriteException e) {
    // TODO Auto-generated catch block
    we.writeResult(RESULTEXCEL, WRITEEXP, NIL, NIL);
   }catch (IOException e) {
    // TODO Auto-generated catch block
    we.writeResult(RESULTEXCEL, WRITEEXP, NIL, NIL);
   }
 }

}