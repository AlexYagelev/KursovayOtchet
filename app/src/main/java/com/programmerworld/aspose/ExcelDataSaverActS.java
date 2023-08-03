package com.programmerworld.aspose;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelDataSaverActS {private String fileName;

    public ExcelDataSaverActS(String fileName) {
        this.fileName = fileName;
    }

    public void saveData12(String t1,String z1,String s1, String d1, String t2,String z2,String s2, String d2,String t3,String z3,String s3,String d3) throws Exception {

// Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download/example.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        Cell cell = worksheet.getCells().get("O1");
        cell.setValue("Наименование прибора");

        cell = worksheet.getCells().get("P1");
        cell.setValue("Заводской номер прибора");

        cell = worksheet.getCells().get("Q1");
        cell.setValue("Свидетельство о поверке  ");

        cell = worksheet.getCells().get("R1");
        cell.setValue("Действительно до ");

         cell = worksheet.getCells().get("S1");
        cell.setValue("Наименование прибора");

        cell = worksheet.getCells().get("T1");
        cell.setValue("Заводской номер прибора");

        cell = worksheet.getCells().get("U1");
        cell.setValue("Свидетельство о поверке  ");

        cell = worksheet.getCells().get("W1");
        cell.setValue("Действительно до ");

         cell = worksheet.getCells().get("X1");
        cell.setValue("Наименование прибора");

        cell = worksheet.getCells().get("Y1");
        cell.setValue("Заводской номер прибора");

        cell = worksheet.getCells().get("Z1");
        cell.setValue("Свидетельство о поверке  ");

        cell = worksheet.getCells().get("AA1");
        cell.setValue("Действительно до ");





// Заполняем таблицу данными
        cell = worksheet.getCells().get("O2");
        cell.setValue(t1);

        cell = worksheet.getCells().get("P2");
        cell.setValue(z1);

        cell = worksheet.getCells().get("Q2");
        cell.setValue(s1);

        cell = worksheet.getCells().get("R2");
        cell.setValue(d1);

        cell = worksheet.getCells().get("S2");
        cell.setValue(t2);

        cell = worksheet.getCells().get("T2");
        cell.setValue(z2);

        cell = worksheet.getCells().get("U2");
        cell.setValue(s2);

        cell = worksheet.getCells().get("W2");
        cell.setValue(d2);

        cell = worksheet.getCells().get("X2");
        cell.setValue(t3);

        cell = worksheet.getCells().get("Y2");
        cell.setValue(z3);

        cell = worksheet.getCells().get("Z2");
        cell.setValue(s3);

        cell = worksheet.getCells().get("AA2");
        cell.setValue(d3);


// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + fileName;
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}
