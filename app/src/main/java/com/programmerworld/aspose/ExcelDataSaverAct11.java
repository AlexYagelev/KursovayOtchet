
package com.programmerworld.aspose;

        import com.aspose.cells.Cell;
        import com.aspose.cells.SaveFormat;
        import com.aspose.cells.Workbook;
        import com.aspose.cells.Worksheet;

public class ExcelDataSaverAct11 {
    private String fileName;

    public ExcelDataSaverAct11(String fileName) {
        this.fileName = fileName;
    }

    public void saveData_1(String TypeWork, String text2, String text3, String text4, String text5, String text6 ) throws Exception {

        // Создаем новый файл Excel с помощью Aspose.Cells
        //File file = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), "example.xlsx");
        // Uri fileUri = Uri.fromFile(file);

        Workbook workbook = new Workbook("/sdcard/Download/example.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);



        Cell cell = worksheet.getCells().get("T1");
        cell.setValue("Гидравлические характеристики");

        cell = worksheet.getCells().get("U1");
        cell.setValue("Стойкость к внешним воздействиям");

        cell = worksheet.getCells().get("V1");
        cell.setValue("Климатическое исполнение");

        cell = worksheet.getCells().get("W1");
        cell.setValue("Показатели надежности безопасности");

        cell = worksheet.getCells().get("X1");
        cell.setValue("Вид привода");

        cell = worksheet.getCells().get("Y1");
        cell.setValue("Масса");







        // Заполняем таблицу данными


        cell = worksheet.getCells().get("T2");
        cell.setValue(TypeWork);

        cell = worksheet.getCells().get("U2");
        cell.setValue(text2);

        cell = worksheet.getCells().get("V2");
        cell.setValue(text3);

        cell = worksheet.getCells().get("W2");
        cell.setValue(text4);

        cell = worksheet.getCells().get("X2");
        cell.setValue(text5);

        cell = worksheet.getCells().get("Y2");
        cell.setValue(text6);

        // Сохраняем файл Excel
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + fileName;
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}
