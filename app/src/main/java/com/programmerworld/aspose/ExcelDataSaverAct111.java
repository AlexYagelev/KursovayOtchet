package com.programmerworld.aspose;

import android.os.Environment;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
//package com.programmerworld.aspose;

public class ExcelDataSaverAct111 {


    String fileName;
     String folderName;
    public static String outputString;


    public ExcelDataSaverAct111(String fileName, String folderName,  String s8) {
        this.fileName = fileName;
        this.folderName =folderName;
        this.outputString = s8.replace("/", "-");


    }
    public String getMyVariable(){
        return outputString ;
    }

    public void saveData1(String s1, String s2, String s3, String s4, String s5, String s6, String s7) throws Exception {

        // Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);


        Cell cell = worksheet.getCells().get("A1");
        cell.setValue("Эксплуатрующее общество");

        cell = worksheet.getCells().get("B1");
        cell.setValue("Подразделение заказщика");

        cell = worksheet.getCells().get("C1");
        cell.setValue("Наименование объекта");

        cell = worksheet.getCells().get("D1");
        cell.setValue("Вид работ");

        cell = worksheet.getCells().get("E1");
        cell.setValue("Содержание работ");

        cell = worksheet.getCells().get("F1");
        cell.setValue("Краткие Т/Х");

        cell = worksheet.getCells().get("G1");
        cell.setValue("xuz");







        // Заполняем таблицу данными
        cell = worksheet.getCells().get("A2");
        cell.setValue(s1);

        cell = worksheet.getCells().get("B2");
        cell.setValue(s2);

        cell = worksheet.getCells().get("C2");
        cell.setValue(s3);

        cell = worksheet.getCells().get("D2");
        cell.setValue(s4);

        cell = worksheet.getCells().get("E2");
        cell.setValue(s5);

        cell = worksheet.getCells().get("F2");
        cell.setValue(s6);

        cell = worksheet.getCells().get("G2");
        cell.setValue(s7);
         outputString = s7.replace("/", "-");
        File folder = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), outputString);
        if (!folder.exists()) {
            folder.mkdir();
        }
        // Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory+ "/" +outputString + "/" + outputString+ ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }


}