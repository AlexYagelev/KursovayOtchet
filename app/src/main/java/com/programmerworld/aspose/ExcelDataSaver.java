package com.programmerworld.aspose;

import android.Manifest;
import android.content.Intent;
import android.net.Uri;
import android.os.Environment;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.ContextCompat;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelDataSaver {
    private AppCompatActivity activity;
    private String fileName;

    public ExcelDataSaver(String fileName) {
        this.activity = activity;
        this.fileName = fileName;
    }

    public void saveData(String manufacturerNumber, String serialNumber, String productionDate, String place, String location, String operationDate) throws Exception {
        int permissionCheck = ContextCompat.checkSelfPermission(activity, Manifest.permission.WRITE_EXTERNAL_STORAGE);


            // Создаем новый файл Excel с помощью Aspose.Cells
            Workbook workbook = new Workbook(String.valueOf(getDocumentUri()));
            Worksheet worksheet = workbook.getWorksheets().get(0);

            // Ваш остальной код сохранения данных

            String outputDirectory = Environment.DIRECTORY_DOWNLOADS;
            String outputFile = outputDirectory + "/" + fileName;

            // Ваш остальной код сохранения файла


        }


    private Uri getDocumentUri() {
        Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        intent.putExtra(Intent.EXTRA_TITLE, fileName);

        activity.startActivityForResult(intent, 2);

        // Ваш код для обработки ответа получения URI после запуска активности выбора файла

        // Возвращаем заглушку как временное решение
        return Uri.EMPTY;
    }

    public void handleActivityResult(int requestCode, int resultCode, Intent data) {
        if (requestCode == 2 && resultCode == AppCompatActivity.RESULT_OK) {
            if (data != null && data.getData() != null) {
                Uri uri = data.getData();

                // Используйте полученный URI для сохранения файла
            }
        }
    }
}