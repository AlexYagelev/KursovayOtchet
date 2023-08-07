package com.programmerworld.aspose;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelDataSaverActS1 {private String fileName;

    public ExcelDataSaverActS1(String fileName) {
        this.fileName = fileName;
    }

    public void saveData12(String t1,String z1,String s1, String t2,String z2,String s2, String t3,String z3,String s3) throws Exception {

// Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download/example.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        Cell cell = worksheet.getCells().get("ad1");
        cell.setValue(" ");

        cell = worksheet.getCells().get("af1");
        cell.setValue("Заводской номер прибора");

        cell = worksheet.getCells().get("ag1");
        cell.setValue("Свидетельство о поверке  ");

        cell = worksheet.getCells().get("ah1");
        cell.setValue("Действительно до ");

        cell = worksheet.getCells().get("ai1");
        cell.setValue("Наименование прибора");

        cell = worksheet.getCells().get("aj1");
        cell.setValue("Заводской номер прибора");

        cell = worksheet.getCells().get("ak1");
        cell.setValue("Свидетельство о поверке  ");

        cell = worksheet.getCells().get("al1");
        cell.setValue("Действительно до ");

        cell = worksheet.getCells().get("am1");
        cell.setValue("Наименование прибора");

        cell = worksheet.getCells().get("an1");
        cell.setValue("Заводской номер прибора");

        cell = worksheet.getCells().get("ao1");
        cell.setValue("Свидетельство о поверке  ");

        cell = worksheet.getCells().get("ap1");
        cell.setValue("Действительно до ");





// Заполняем таблицу данными
        cell = worksheet.getCells().get("ad2");
        cell.setValue(t1);

        cell = worksheet.getCells().get("ae2");
        cell.setValue(z1);

        cell = worksheet.getCells().get("af2");
        cell.setValue(s1);



        cell = worksheet.getCells().get("ah2");
        cell.setValue(t2);

        cell = worksheet.getCells().get("ai2");
        cell.setValue(z2);

        cell = worksheet.getCells().get("aj2");
        cell.setValue(s2);


        cell = worksheet.getCells().get("al2");
        cell.setValue(t3);

        cell = worksheet.getCells().get("am2");
        cell.setValue(z3);

        cell = worksheet.getCells().get("an2");
        cell.setValue(s3);



// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + fileName;
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}
