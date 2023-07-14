package com.programmerworld.aspose;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
//package com.programmerworld.aspose;

public class ExcelDataSaver {

    private String fileName;

    public ExcelDataSaver(String fileName) {
        this.fileName = fileName;
    }

    public void saveData(String manufacturerNumber, String serialNumber, String productionDate, String place, String location, String operationDate) throws Exception {

        // Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download/example.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);


        Cell cell = worksheet.getCells().get("H1");
        cell.setValue("Завод-Изготовитель номер");

        cell = worksheet.getCells().get("I1");
        cell.setValue("Заводской номер");

        cell = worksheet.getCells().get("J1");
        cell.setValue("Дата изготовления");

        cell = worksheet.getCells().get("K1");
        cell.setValue("Место Эксплуатации");

        cell = worksheet.getCells().get("L1");
        cell.setValue("Регистрационный номер");

        cell = worksheet.getCells().get("M1");
        cell.setValue("Дата ввода в эксплуатацию");







        // Заполняем таблицу данными
        cell = worksheet.getCells().get("H2");
        cell.setValue(manufacturerNumber);

        cell = worksheet.getCells().get("I2");
        cell.setValue(serialNumber);

        cell = worksheet.getCells().get("J2");
        cell.setValue(place);

        cell = worksheet.getCells().get("K2");
        cell.setValue(productionDate);

        cell = worksheet.getCells().get("L2");
        cell.setValue(location);

        cell = worksheet.getCells().get("M2");
        cell.setValue(operationDate);


        // Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + fileName;
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}