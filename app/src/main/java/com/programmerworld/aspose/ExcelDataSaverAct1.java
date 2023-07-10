package com.programmerworld.aspose;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelDataSaverAct1 {
    private String fileName;

    public ExcelDataSaverAct1(String fileName) {
        this.fileName = fileName;
    }

    public void saveData_1(String TypeWork, String text2, String text3, String text4, String text5, String text6 ) throws Exception {

        // Создаем новый файл Excel с помощью Aspose.Cells
        //File file = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), "example.xlsx");
       // Uri fileUri = Uri.fromFile(file);

        Workbook workbook = new Workbook("Download/gd.xl/example.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);



        Cell cell = worksheet.getCells().get("H1");
        cell.setValue("Номинальный диаметр");

         cell = worksheet.getCells().get("I1");
        cell.setValue("Номинальное давление");

         cell = worksheet.getCells().get("J1");
        cell.setValue("Температура рабочей среды");

         cell = worksheet.getCells().get("K1");
        cell.setValue("Рабочая среда");

         cell = worksheet.getCells().get("L1");
        cell.setValue("Герметичность затвора");

         cell = worksheet.getCells().get("M1");
        cell.setValue("Тип присоединения к трубопроводу");







        // Заполняем таблицу данными


        cell = worksheet.getCells().get("H2");
        cell.setValue(TypeWork);

        cell = worksheet.getCells().get("I2");
        cell.setValue(text2);

        cell = worksheet.getCells().get("J2");
        cell.setValue(text3);

        cell = worksheet.getCells().get("K2");
        cell.setValue(text4);

        cell = worksheet.getCells().get("L2");
        cell.setValue(text5);

        cell = worksheet.getCells().get("M2");
        cell.setValue(text6);

        // Сохраняем файл Excel
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + fileName;
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}
