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
        ///////сссылки//////
        String kaspiUrl = "C:/Rest/Test/KaspiUPL.xlsx";
        String restUrl = "C:/Rest/Test/130422.xls";
        String midUrl = "C:/Rest/Test/ParsingMid2.xls";

        ///////подключение к файлу остатков////////////
        FileInputStream fis = new FileInputStream(restUrl);
        Workbook rest = new HSSFWorkbook(fis);
        HSSFSheet restSheet = (HSSFSheet) rest.getSheetAt(0);

        /////////////скан ячеек файла остатков/////////
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
            System.out.println("Файл остатков считан верно, всего строк в файле остатков: " + restRows + " Считалось: " + restList.size());
        }



        /////////анализ файла остатков//////////////////
        ArrayList ReservedCounter = new ArrayList();
        for (int i = 0; i < restRows; i++) {
            try{switch (restSheet.getRow(i).getCell(2).getCellType()){
                case STRING -> ReservedCounter.add(restSheet.getRow(i).getCell(2).getStringCellValue());
                case ERROR, NUMERIC, BLANK, FORMULA, BOOLEAN, _NONE -> {}}}catch (NullPointerException exception){
                continue;
            }
        }

        int restStart = ReservedCounter.indexOf("Основной склад")+4;
        int restFinish = ReservedCounter.indexOf("РЕЗЕРВ")+6;

        ///////////создание листа артикулов////////////////
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

        System.out.println("Артикулы файла остатков(всего "+restSkuList.size()+"): "+restSkuList);
        System.out.println("Первый артикул: " + restSkuList.get(0) + " Последний: " + restSkuList.get(restSkuList.size()-1));

        ///////////////подключение к MID///////////////////////
        FileInputStream fisMid = new FileInputStream(midUrl);
        Workbook mid = new HSSFWorkbook(fisMid);
        HSSFSheet midSheet = (HSSFSheet) mid.getSheetAt(0);


        int midRows = midSheet.getLastRowNum();
        int actualPrice;
        int inPrice;
        int trdPrice;
        int nxtPrice;
        int lowestPrice;
        String skuMid;
        for (int i = 1; i <= midRows; i++) {
            ////////////если на первом месте MusicPark и больше нет продавцов - то цена Розница/////////////
            if(Objects.equals(midSheet.getRow(i).getCell(3).getStringCellValue(), "MusicPark")){
                if (midSheet.getRow(i).getCell(4).getNumericCellValue()==0){
                    skuMid = midSheet.getRow(i).getCell(0).getStringCellValue();
                   try{ trdPrice = (int)restSheet.getRow(restList.indexOf(skuMid)).getCell(6).getNumericCellValue();
                       trdPrice = (int)(trdPrice*1.15);
                       midSheet.getRow(i).createCell(5).setCellValue(trdPrice);}
                   catch (NullPointerException exception){}
                }
                ///////если на первом месте MusicPark, но есть еще продавцы, то ставит цену либо ниже конкурента на 2%,
                ///////либо розницу, если цена конкурента выше розницы
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
                    //System.out.println("уже снят с продажи: " + midSheet.getRow(i).getCell(1).getStringCellValue() + "(отсутствует в наличии на складе Алмата)");
                }


                }


            }else{try{
                lowestPrice = (int)midSheet.getRow(i).getCell(2).getNumericCellValue();
                skuMid = midSheet.getRow(i).getCell(0).getStringCellValue();
                inPrice = (int) restSheet.getRow(restList.indexOf(skuMid)).getCell(4).getNumericCellValue();
                trdPrice = (int)restSheet.getRow(restList.indexOf(skuMid)).getCell(6).getNumericCellValue();
                trdPrice = (int)(trdPrice*1.15);
                if(lowestPrice > (inPrice*1.13)){
                    if(lowestPrice>trdPrice){actualPrice = trdPrice;midSheet.getRow(i).createCell(5).setCellValue(actualPrice);}
                    actualPrice=(int)(lowestPrice*0.98);
                    midSheet.getRow(i).createCell(5).setCellValue(actualPrice);
                }else{
                    System.out.println("Цена конкурента ниже себестоимости на товар: " + midSheet.getRow(i).getCell(1).getStringCellValue());
                    midSheet.getRow(i).createCell(5).setCellValue(trdPrice);
                }
            }catch (NullPointerException e){
//                System.out.println("уже снят с продажи: " + midSheet.getRow(i).getCell(1).getStringCellValue() + "(отсутствует в наличии на складе Алмата)");
            }

        }

        }
        FileOutputStream fos = new FileOutputStream(midUrl);
        mid.write(fos);
        fos.close();















    }
}
