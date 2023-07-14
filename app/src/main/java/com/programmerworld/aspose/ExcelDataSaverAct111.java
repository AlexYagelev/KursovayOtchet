package com.programmerworld.aspose;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
//package com.programmerworld.aspose;

public class ExcelDataSaverAct111 {

    private String fileName;

    public ExcelDataSaverAct111(String fileName) {
        this.fileName = fileName;
    }

    public void saveData1(String s1, String s2, String s3, String s4, String s5, String s6, String s7) throws Exception {

        // Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook();
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


        // Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + fileName;
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}