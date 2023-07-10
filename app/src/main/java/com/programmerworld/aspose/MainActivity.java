package com.programmerworld.aspose;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Bundle;
import android.widget.EditText;
import android.widget.Spinner;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.OutputStream;

public class MainActivity extends AppCompatActivity {
    private static final int REQUEST_WRITE_EXTERNAL_STORAGE = 1;
    private static final String FILE_NAME = "output.xlsx";
    private Spinner spinner; private Uri uri;
    private Workbook workbook;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        findViewById(R.id.button).setOnClickListener(v -> {
            if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    == PackageManager.PERMISSION_GRANTED) {
// Разрешение уже предоставлено, можно создавать файл
                // createExcelFile();
save();






                spinner = findViewById(R.id.spinner);
                 int selectedItemPosition = spinner.getSelectedItemPosition();
                if (selectedItemPosition == 0) {
                    // переключаем на другую активность для выбора 1
                    Intent intent1 = new Intent(MainActivity.this, Activity_type_1.class);
                    //intent1.putExtra("uri", uri.toString());

                    startActivity(intent1);
                } else if (selectedItemPosition == 1) {
                    // переключаем на другую активность для выбора 2
                    Intent intent1 = new Intent(MainActivity.this, Activity_type_2.class);
                    startActivity(intent1);
                }





            }
            else {
// Запрос разрешения
                ActivityCompat.requestPermissions(MainActivity.this,
                        new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},
                        REQUEST_WRITE_EXTERNAL_STORAGE);
            }
        });
    }

    private void save(){
        ExcelDataSaver dataSaver = new ExcelDataSaver("example.xlsx");
        try {



            EditText editText1 = findViewById(R.id.editText2);
            String text1 = editText1.getText().toString();




            // Вставить значение в первую ячейку


            EditText editText = findViewById(R.id.editText_1);
            String text = editText.getText().toString();



            EditText editText2 = findViewById(R.id.editText3);
            String text2 = editText2.getText().toString();




            EditText editText3 = findViewById(R.id.editText5);
                    String text3 = editText3.getText().toString();





                        EditText editText4 = findViewById(R.id.editText6);
                        String text4 = editText4.getText().toString();





                        EditText editText5 = findViewById(R.id.editText7);
                        String text5 = editText5.getText().toString();













            dataSaver.saveData(text, text1, text2,text3, text4, text5);
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }

    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (requestCode == REQUEST_WRITE_EXTERNAL_STORAGE) {
            if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
// Разрешение предоставлено, можно создавать файл
                createExcelFile();
            } else {
                Toast.makeText(this, "Permission denied", Toast.LENGTH_SHORT).show();
            }
        }
    }



    private ActivityResultLauncher<String> createFileLauncher = registerForActivityResult(
            new ActivityResultContracts.CreateDocument(),
            result -> {
                if (result != null) {
                    uri = result;

                    try {
                         workbook = new Workbook();
                        Worksheet worksheet = workbook.getWorksheets().get(0);
                        /* Cells cells = worksheet.getCells();
                        Cell cell = cells.get("Б1");
                        cell.setValue("1");*/


                        // Вставить значение в первую ячейку
                        Cell firstCell1 = worksheet.getCells().get("B1");
                        firstCell1.setValue("Завод-Изготовитель номер");

                        EditText editText1 = findViewById(R.id.editText2);
                        String text1 = editText1.getText().toString();

                        Cell cell1 = worksheet.getCells().get("B2");
                        cell1.setValue(text1);


                        // Вставить значение в первую ячейку
                        Cell firstCell = worksheet.getCells().get("C1");
                        firstCell.setValue("заводской номер");

                        EditText editText = findViewById(R.id.editText_1);
                        String text = editText.getText().toString();

                        Cell cell = worksheet.getCells().get("C2");
                        cell.setValue(text);

                        // Вставить значение в первую ячейку
                        Cell firstCell2 = worksheet.getCells().get("D1");
                        firstCell2.setValue("Дата изготовления");

                        EditText editText2 = findViewById(R.id.editText3);
                        String text2 = editText2.getText().toString();

                        Cell cell2 = worksheet.getCells().get("D2");
                        cell2.setValue(text2);

                        // Вставить значение в первую ячейку
                     /*

                        EditText editText3 = findViewById(R.id.editText4);
                        String text3 = editText3.getText().toString();



                        // Вставить значение в первую ячейку
                     /*   Cell firstCell4 = worksheet.getCells().get("F1");
                        firstCell4.setValue("Место эксплуатирования");

                        EditText editText4 = findViewById(R.id.editText5);
                        String text4 = editText4.getText().toString();

                        Cell cell4 = worksheet.getCells().get("F2");
                        cell4.setValue(text4);

                        // Вставить значение в первую ячейку
                        Cell firstCell5 = worksheet.getCells().get("G1");
                        firstCell5.setValue("Дата ввода в эксплуатацию");

                        EditText editText5 = findViewById(R.id.editText6);
                        String text5 = editText5.getText().toString();

                        Cell cell5 = worksheet.getCells().get("G2");
                        cell5.setValue(text5);

*/




                        // Вставить значение во вторую ячейку
                        /*Cell secondCell = worksheet.getCells().get("B2");
                        secondCell.setValue("Значение во второй ячейке");
*/

                         OutputStream outputStream = getContentResolver().openOutputStream(uri);
                        if (outputStream != null) {
                            workbook.save(outputStream, SaveFormat.XLSX);
                            Toast.makeText(this, "File saved successfully: " + uri.toString(), Toast.LENGTH_SHORT).show();
                        } else {
                            Toast.makeText(this, "Failed to save file", Toast.LENGTH_SHORT).show();
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                        Toast.makeText(this, "Failed to save file: " + e.getMessage(), Toast.LENGTH_SHORT).show();
                    }
                } /*else {
                    Toast.makeText(this, "Failed to create file", Toast.LENGTH_SHORT).show();
                }*/
            }
    );

    private void createExcelFile() {
        final String FILE_NAME = "example.xlsx";


        Intent intent_e = new Intent(this, Activity_type_1.class);

        
        
        Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        intent.putExtra(Intent.EXTRA_TITLE, FILE_NAME);
        //createFileLauncher.launch(FILE_NAME);

        /*spinner = findViewById(R.id.spinner);
        int selectedItemPosition = spinner.getSelectedItemPosition();
        if (selectedItemPosition == 0) {
            // переключаем на другую активность для выбора 1
            Intent intent1 = new Intent(MainActivity.this, Activity_type_1.class);
           //intent1.putExtra("uri", uri.toString());

            startActivity(intent1);
        } else if (selectedItemPosition == 1) {
            // переключаем на другую активность для выбора 2
            Intent intent1 = new Intent(MainActivity.this, Activity_type_2.class);
            startActivity(intent1);
        }*/
    }}