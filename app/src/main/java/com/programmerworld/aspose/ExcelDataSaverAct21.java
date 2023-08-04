package com.programmerworld.aspose;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelDataSaverAct21 {

    private String fileName;

    public ExcelDataSaverAct21(String fileName) {
        this.fileName = fileName;
    }

    public void saveData(String nameObject, String uslObz, String nameFactory, String yearGo, String yearLGo, String rabDav, String nomD) throws Exception {

// Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download/example.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        Cell cell = worksheet.getCells().get("H1");
        cell.setValue("Наименование ТУ");

        cell = worksheet.getCells().get("I1");
        cell.setValue("Наименование завода-изготовителя");

        cell = worksheet.getCells().get("J1");
        cell.setValue("Год изготовления");

        cell = worksheet.getCells().get("K1");
        cell.setValue("Год ввода в эксплуатацию");

        cell = worksheet.getCells().get("L1");
        cell.setValue("Рабочее давление, кгс/см2 (МПа)");

        cell = worksheet.getCells().get("M1");
        cell.setValue("Номинальный наружный или внутренний диаметр, мм");

        cell = worksheet.getCells().get("N1");
        cell.setValue("Наименование рабочей среды");

// Заполняем таблицу данными
        cell = worksheet.getCells().get("H2");
        cell.setValue(nameObject);

        cell = worksheet.getCells().get("I2");
        cell.setValue(uslObz);

        cell = worksheet.getCells().get("J2");
        cell.setValue(nameFactory);

        cell = worksheet.getCells().get("K2");
        cell.setValue(yearGo);

        cell = worksheet.getCells().get("L2");
        cell.setValue(yearLGo);

        cell = worksheet.getCells().get("M2");
        cell.setValue(rabDav);

        cell = worksheet.getCells().get("N2");
        cell.setValue(nomD);

// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + fileName;
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}

