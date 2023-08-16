package com.programmerworld.aspose;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelDataSaverAct21b {

    private String fileName;

    public ExcelDataSaverAct21b(String fileName) {
        this.fileName = fileName;
    }

    public void saveData(String nameObject, String uslObz, String nameFactory, String yearGo, String yearLGo) throws Exception {

// Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download"+"/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString+".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        Cell cell = worksheet.getCells().get("o1");
        cell.setValue("Рабочая температура среды, ℃");

        cell = worksheet.getCells().get("p1");
        cell.setValue("Марка материала корпуса");

        cell = worksheet.getCells().get("q1");
        cell.setValue("Вместимость, м³");

        cell = worksheet.getCells().get("r1");
        cell.setValue("Схема подключения сосуда в установку");

        cell = worksheet.getCells().get("s1");
        cell.setValue("Климатическое исполнение");



// Заполняем таблицу данными
        cell = worksheet.getCells().get("o2");
        cell.setValue(nameObject);

        cell = worksheet.getCells().get("p2");
        cell.setValue(uslObz);

        cell = worksheet.getCells().get("q2");
        cell.setValue(nameFactory);

        cell = worksheet.getCells().get("r2");
        cell.setValue(yearGo);

        cell = worksheet.getCells().get("s2");
        cell.setValue(yearLGo);



// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory +"/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString+".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}

