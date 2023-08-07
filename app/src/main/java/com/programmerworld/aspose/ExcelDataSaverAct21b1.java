package com.programmerworld.aspose;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelDataSaverAct21b1 {

    private String fileName;

    public ExcelDataSaverAct21b1(String fileName) {
        this.fileName = fileName;
    }

    public void saveData(String nameObject, String uslObz, String nameFactory, String yearGo, String yearLGo,String a, String b, String c, String d, String i) throws Exception {

// Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download/example.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        Cell cell = worksheet.getCells().get("t1");
        cell.setValue("Наименование элемента сосуда");

        cell = worksheet.getCells().get("u1");
        cell.setValue("Количество");

        cell = worksheet.getCells().get("v1");
        cell.setValue("Диаметр");

        cell = worksheet.getCells().get("w1");
        cell.setValue("Высота");

        cell = worksheet.getCells().get("x1");
        cell.setValue("Толщина стенки");

        cell = worksheet.getCells().get("y1");
        cell.setValue("Марка стали");

        cell = worksheet.getCells().get("z1");
        cell.setValue("ГОСТ или ТУ");

        cell = worksheet.getCells().get("aa1");
        cell.setValue("Вид сварки");

        cell = worksheet.getCells().get("ab1");
        cell.setValue("Электрод");

        cell = worksheet.getCells().get("ac1");
        cell.setValue("Метод НК");



// Заполняем таблицу данными
        cell = worksheet.getCells().get("t2");
        cell.setValue(nameObject);

        cell = worksheet.getCells().get("u2");
        cell.setValue(uslObz);

        cell = worksheet.getCells().get("v2");
        cell.setValue(nameFactory);

        cell = worksheet.getCells().get("w2");
        cell.setValue(yearGo);

        cell = worksheet.getCells().get("x2");
        cell.setValue(yearLGo);

        cell = worksheet.getCells().get("y2");
        cell.setValue(a);

        cell = worksheet.getCells().get("z2");
        cell.setValue(b);

        cell = worksheet.getCells().get("aa2");
        cell.setValue(c);

        cell = worksheet.getCells().get("ab2");
        cell.setValue(d);

        cell = worksheet.getCells().get("ac2");
        cell.setValue(i);



// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + fileName;
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}

