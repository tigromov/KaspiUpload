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
import java.sql.SQLOutput;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Objects;

public class Main {
    public static void main(String[] args) throws IOException {
        ///////�������//////
        String kaspiUrl = "C:/Rest/Test/KaspiUPL.xlsx";
        String restUrl = "C:/Rest/Test/290422.xls";
        String midUrl = "C:/Rest/Test/2604mid.xls";

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

        /////////////////////////////////////////////////////������ �� �������/////////
        ArrayList restReservedList = new ArrayList();
        for (int i = 0; i < restRows; i++) {
            try{switch (restSheet.getRow(i).getCell(2).getCellType()){
                case STRING -> {
                    restReservedList.add(restSheet.getRow(i).getCell(2).getStringCellValue());
                    break;
                }
                case NUMERIC -> {
                    restReservedList.add(Objects.toString((int)restSheet.getRow(i).getCell(2).getNumericCellValue()));
                    break;
                }
                case BLANK, ERROR,_NONE, BOOLEAN, FORMULA -> {
                    restReservedList.add(Objects.toString(restSheet.getRow(i).getCell(2).getStringCellValue()));
                    break;
                }
            }}catch (NullPointerException e){
                restReservedList.add(" ");

            }


        }

        int rsrSave = restRows-restReservedList.indexOf("������");
        System.out.println("���������� ������������ � �������: " + rsrSave);
        ///////////////////////////////////////////////////////////////////////////////

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
        int actualPrice;
        int inPrice;
        int trdPrice;
        int nxtPrice;
        int lowestPrice;
        int specialPrice;
        String skuMid;
        String skuRest;
        String yes = "yes";
        String no = "no";
        String name;
        String brand;


        for (int i = 1; i <= midRows; i++) {
            ////////////���� �� ������ ����� MusicPark � ������ ��� ��������� - �� ���� �������/////////////
            if(Objects.equals(midSheet.getRow(i).getCell(3).getStringCellValue(), "MusicPark")){
                if (midSheet.getRow(i).getCell(4).getNumericCellValue()==0){
                    skuMid = midSheet.getRow(i).getCell(0).getStringCellValue();
                   try{ trdPrice = (int)restSheet.getRow(restList.indexOf(skuMid)).getCell(6).getNumericCellValue();
                       trdPrice = (int)(trdPrice*1.15);
                       midSheet.getRow(i).createCell(5).setCellValue(trdPrice);}
                   catch (NullPointerException exception){}
                }
                ///////���� �� ������ ����� MusicPark, �� ���� ��� ��������, �� ������ ���� ���� ���� ���������� �� 2%,
                ///////���� �������, ���� ���� ���������� ���� �������
                else{
                    try{
                        nxtPrice = (int)midSheet.getRow(i).getCell(4).getNumericCellValue();
                        skuMid = midSheet.getRow(i).getCell(0).getStringCellValue();
                        inPrice = (int) restSheet.getRow(restList.indexOf(skuMid)).getCell(4).getNumericCellValue();
                        trdPrice = (int)restSheet.getRow(restList.indexOf(skuMid)).getCell(6).getNumericCellValue();
                        trdPrice = (int)(trdPrice*1.15);
                        if(inPrice < nxtPrice& nxtPrice<=trdPrice){ actualPrice = (int) (nxtPrice*0.98);}
                            else{ actualPrice = trdPrice; }
                        midSheet.getRow(i).createCell(5).setCellValue(actualPrice);}

                    catch (NullPointerException e){
                    //System.out.println("��� ���� � �������: " + midSheet.getRow(i).getCell(1).getStringCellValue() + "(����������� � ������� �� ������ ������)");
                }


                }
                  ////////////////���� ��������� �������� ���� ����, �� ������ ��� ���� �� 2%, ���� ��� �� ��������� �����
                //////////////////+13%, ���� ���������, �� ������ �������


            }else{try{
                lowestPrice = (int)midSheet.getRow(i).getCell(2).getNumericCellValue();
                skuMid = midSheet.getRow(i).getCell(0).getStringCellValue();
                inPrice = (int) restSheet.getRow(restList.indexOf(skuMid)).getCell(4).getNumericCellValue();
                trdPrice = (int)restSheet.getRow(restList.indexOf(skuMid)).getCell(6).getNumericCellValue();

                if(lowestPrice > (inPrice*1.13)){
                    if(lowestPrice>trdPrice){actualPrice = trdPrice;midSheet.getRow(i).createCell(5).setCellValue(actualPrice);}
                    actualPrice=(int)(lowestPrice*0.98);
                    midSheet.getRow(i).createCell(5).setCellValue(actualPrice);
                }else{
                    System.out.println("���� ���������� ���� ������������� �� �����: " + midSheet.getRow(i).getCell(1).getStringCellValue());
                    midSheet.getRow(i).createCell(5).setCellValue(trdPrice);
                }
            }catch (NullPointerException e){
//                System.out.println("��� ���� � �������: " + midSheet.getRow(i).getCell(1).getStringCellValue() + "(����������� � ������� �� ������ ������)");
            }

        }

        }
        FileOutputStream fos = new FileOutputStream(midUrl);
        mid.write(fos);
        fos.close();


///////////////////������������ �����////////////
        FileInputStream fisKaspi = new FileInputStream(kaspiUrl);
        XSSFWorkbook kspbook = new XSSFWorkbook(fisKaspi);
        XSSFSheet kaspiSheet = (XSSFSheet) kspbook.getSheetAt(0);

///////////////////////////////////////////////////////////
        ArrayList midSkuList = new ArrayList();
        for (int i = 1; i < midRows+1; i++) {
            midSkuList.add(midSheet.getRow(i).getCell(0).getStringCellValue());


        }




        for (int i = 1; i < restList.size()-rsrSave; i++) {
            if (restSkuList.contains(restList.get(i))){
                /////�������///////
                int kaspiLastRaw = kaspiSheet.getLastRowNum()+1;
                skuRest = Objects.toString(restList.get(i));
                kaspiSheet.createRow(kaspiLastRaw).createCell(0).setCellValue(skuRest);

                /////��������//////

                name = restSheet.getRow(i).getCell(2).getStringCellValue();
                kaspiSheet.getRow(kaspiLastRaw).createCell(1).setCellValue(String.valueOf(name));

                /////�����////////
                int index = name.indexOf(' ');
                brand = name.substring(0, index);
                kaspiSheet.getRow(kaspiLastRaw).createCell(2).setCellValue(String.valueOf(brand));

                /////////////��������//////////////


                
                /////////����//////////
                try{if (brand.equals("GregBennett")){
                    specialPrice = (int)restSheet.getRow(i).getCell(5).getNumericCellValue();
                    kaspiSheet.getRow(kaspiLastRaw).createCell(3).setCellValue(specialPrice);
                    System.out.println("GregBennet ����� ����� " + specialPrice + name);
                }else{
                actualPrice = (int)midSheet.getRow(midSkuList.indexOf(skuRest)+1).getCell(5).getNumericCellValue();
                kaspiSheet.getRow(kaspiLastRaw).createCell(3).setCellValue(actualPrice);}}
                catch (NullPointerException e){
                    if (brand.equals("GregBennett")){
                        specialPrice = (int)restSheet.getRow(i).getCell(5).getNumericCellValue();
                        kaspiSheet.getRow(kaspiLastRaw).createCell(3).setCellValue(specialPrice);
                        System.out.println("GregBennet ����� ����� " + specialPrice + name);
                    }else{

                    actualPrice = (int)restSheet.getRow(restList.indexOf(skuRest)).getCell(6).getNumericCellValue();
                    kaspiSheet.getRow(kaspiLastRaw).createCell(3).setCellValue(actualPrice);
                }}


                ///////////yes///////////
                kaspiSheet.getRow(kaspiLastRaw).createCell(4).setCellValue(yes);
                kaspiSheet.getRow(kaspiLastRaw).createCell(5).setCellValue(no);
                kaspiSheet.getRow(kaspiLastRaw).createCell(6).setCellValue(no);
                kaspiSheet.getRow(kaspiLastRaw).createCell(7).setCellValue(no);
                kaspiSheet.getRow(kaspiLastRaw).createCell(8).setCellValue(no);

            }



        }

        FileOutputStream fosksp = new FileOutputStream(kaspiUrl);
        kspbook.write(fosksp);
        fosksp.close();










    }
}
