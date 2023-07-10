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
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.getWorksheets().get(0);


        Cell cell = worksheet.getCells().get("B1");
        cell.setValue("Завод-Изготовитель номер");

        cell = worksheet.getCells().get("C1");
        cell.setValue("Заводской номер");

        cell = worksheet.getCells().get("D1");
        cell.setValue("Дата изготовления");

        cell = worksheet.getCells().get("E1");
        cell.setValue("Место Эксплуатации");

        cell = worksheet.getCells().get("F1");
        cell.setValue("Регистрационный номер");

        cell = worksheet.getCells().get("G1");
        cell.setValue("Дата ввода в эксплуатацию");







        // Заполняем таблицу данными
        cell = worksheet.getCells().get("B2");
        cell.setValue(manufacturerNumber);

        cell = worksheet.getCells().get("C2");
        cell.setValue(serialNumber);

        cell = worksheet.getCells().get("E2");
        cell.setValue(place);

        cell = worksheet.getCells().get("D2");
        cell.setValue(productionDate);

        cell = worksheet.getCells().get("F2");
        cell.setValue(location);

        cell = worksheet.getCells().get("G2");
        cell.setValue(operationDate);

        // Сохраняем файл Excel
        String outputDirectory = "/Внутренний общий накопитель/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + fileName;
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}