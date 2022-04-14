import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Objects;

public class Main {
    public static void main(String[] args) throws IOException {
        ///////�������//////
        String kaspiUrl = "C:/Rest/Test/KaspiUPL.xlsx";
        String restUrl = "C:/Rest/Test/130422.xls";
        String midUrl = "C:/Rest/Test/ParsingMid2.xls";

        ///////����������� � ����� ��������////////////
        FileInputStream fis = new FileInputStream(restUrl);
        Workbook rest = new HSSFWorkbook(fis);
        HSSFSheet restSheet = (HSSFSheet) rest.getSheetAt(0);

        /////////////���� ����� ����� ��������/////////
        int restRows = restSheet.getLastRowNum();
        ArrayList restList = new ArrayList();
        for (int i = 0; i < restRows; i++) {
            try{switch (restSheet.getRow(i).getCell(1).getCellType()){
                case STRING -> {
                    restList.add(restSheet.getRow(i).getCell(1).getStringCellValue());
                    break;
                }
                case NUMERIC -> {
                    restList.add(Objects.toString((int)restSheet.getRow(i).getCell(1).getNumericCellValue()));
                    break;
                }
                case BLANK, ERROR,_NONE, BOOLEAN, FORMULA -> {
                    restList.add(Objects.toString(restSheet.getRow(i).getCell(1).getStringCellValue()));
                    break;
                }
            }}catch (NullPointerException e){
                restList.add(" ");

            }


        }

        if(restList.size()==restRows){
            System.out.println("���� �������� ������ �����, ����� ����� � ����� ��������: " + restRows + " ���������: " + restList.size());
        }



        /////////������ ����� ��������//////////////////
        ArrayList ReservedCounter = new ArrayList();
        for (int i = 0; i < restRows; i++) {
            try{switch (restSheet.getRow(i).getCell(2).getCellType()){
                case STRING -> ReservedCounter.add(restSheet.getRow(i).getCell(2).getStringCellValue());
                case ERROR, NUMERIC, BLANK, FORMULA, BOOLEAN, _NONE -> {}}}catch (NullPointerException exception){
                continue;
            }
        }

        int restStart = ReservedCounter.indexOf("�������� �����")+4;
        int restFinish = ReservedCounter.indexOf("������")+6;

        ///////////�������� ����� ���������////////////////
        ArrayList restSkuList = new ArrayList();



        for (int i = restStart; i < restFinish; i++) {
            try{

            switch(restSheet.getRow(i).getCell(1).getCellType())
            {       case NUMERIC:
                        restSkuList.add(Objects.toString((int)restSheet.getRow(i).getCell(1).getNumericCellValue()));break;
                    case STRING:
                        restSkuList.add(Objects.toString(restSheet.getRow(i).getCell(1).getStringCellValue()));break;
                case FORMULA,BOOLEAN,_NONE, ERROR, BLANK:
            }
            }catch (NullPointerException e){}
        }

        System.out.println("�������� ����� ��������(����� "+restSkuList.size()+"): "+restSkuList);
        System.out.println("������ �������: " + restSkuList.get(0) + " ���������: " + restSkuList.get(restSkuList.size()-1));

        ///////////////����������� � MID///////////////////////
        FileInputStream fisMid = new FileInputStream(midUrl);
        Workbook mid = new HSSFWorkbook(fisMid);
        HSSFSheet midSheet = (HSSFSheet) mid.getSheetAt(0);


        int midRows = midSheet.getLastRowNum();
        String actualPrice;
        for (int i = 1; i <= midRows; i++) {
            if(Objects.equals(midSheet.getRow(i).getCell(3).getStringCellValue(), "MusicPark")){
                if (midSheet.getRow(i).getCell(4).getNumericCellValue()==0){
                    String skuMid = midSheet.getRow(i).getCell(0).getStringCellValue();
                   try{double trPrice = restSheet.getRow(restList.indexOf(skuMid)).getCell(6).getNumericCellValue();
                       int trdPrice = Integer.valueOf((int) trPrice);
                       midSheet.getRow(i).createCell(5).setCellValue(trdPrice);
                   }catch (NullPointerException exception){}
                }
                

            }

        }
        FileOutputStream fos = new FileOutputStream(midUrl);
        mid.write(fos);
        fos.close();















    }
}