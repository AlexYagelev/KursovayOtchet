package com.programmerworld.aspose.sosudi;

import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.ContextCompat;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.programmerworld.aspose.ExcelDataSaverActS;
import com.programmerworld.aspose.R;

import java.util.ArrayList;
import java.util.List;

public class Activity_type_2_sosudi extends AppCompatActivity {
    private EditText editText;
    private Button button;
    private int cellNumber = 0;
    private TextView counterText;
    private int counter = 0;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type2_sosudi);
        findViewById(R.id.button1s).setOnClickListener(v -> {
            if (ContextCompat.checkSelfPermission(Activity_type_2_sosudi.this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    == PackageManager.PERMISSION_GRANTED) {
                save();
                Intent intent = new Intent(Activity_type_2_sosudi.this, Activity_type_2_sosudi1.class);
                startActivity(intent);
            }
        });

        findViewById(R.id.button2s).setOnClickListener(v -> {
            if (ContextCompat.checkSelfPermission(Activity_type_2_sosudi.this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    == PackageManager.PERMISSION_GRANTED) {
                save();
                Activity_type_2_sosudi.this.finish();

                // Вернуться на предыдущий фрагмент
                getSupportFragmentManager().popBackStack();
            }
        });


        editText = findViewById(R.id.editText13);
        button = findViewById(R.id.button5);

        button.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                saveTextToExcel();
            }
        });
        counterText = findViewById(R.id.counterText);

        // при старте устанавливаем текст
        counterText.setText("Счетчик: " + counter);
    }


    private void saveTextToExcel() {
        String text = editText.getText().toString().trim();

        // String filePath = getPreferences(MODE_PRIVATE).getString("excel_file", "");

        try {
            // Открываем xlsx файл
            Workbook workbook = new Workbook("/sdcard/Download/example.xlsx");

            // Получаем первый лист
            Worksheet worksheet = workbook.getWorksheets().get(0);

            // Записываем текст в ячейку с номером cellNumber
            worksheet.getCells().get(cellNumber, 27).setValue(text);

            // Увеличиваем счетчик ячейки
            cellNumber++;

            // Сохраняем xlsx файл
            workbook.save("/sdcard/Download/example.xlsx");
            counter++;

            Toast.makeText(this, "Запись No" + counter + " сохранена", Toast.LENGTH_SHORT).show();
            counterText.setText("Точка: " + counter);
            editText.setText("");

        } catch (Exception e) {
            showError("Не удалось сохранить в Excel", e);
        }
    }

    private void showError(String message, Exception e) {
        // Показываем сообщение об ошибке
    }


    private List<List<String>> loadDataFromExcel3() throws Exception {
        // Загрузка данных из Excel с помощью Aspose.Cells

        Workbook workbook = new Workbook("/sdcard/Download/Книга1.xlsx");

        // Получаем первый лист
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Создадим список для данных
        List<List<String>> data = new ArrayList<>();

        // Проходим по всем строкам
        for (int rowIndex = 0; rowIndex < worksheet.getCells().getMaxDataRow()+1; rowIndex++) {
            // Данные из одной строки
            List<String> rowData = new ArrayList<>();

            // Проходим по всем ячейкам в строке
            for (int colIndex = 0; colIndex < worksheet.getCells().getMaxDataColumn()+1; colIndex++) {
                Cell cell = worksheet.getCells().get(rowIndex, colIndex);

                // Добавляем данные ячейки в список строки
                rowData.add(cell.getStringValue());
            }
            // Добавляем список данных строки в общий список
            data.add(rowData);
        }    return data;
    }



    private void save(){
        ExcelDataSaverActS dataSaver = new ExcelDataSaverActS("example.xlsx");
        try {
            EditText editText = findViewById(R.id.editText1);
            String text = editText.getText().toString();

            EditText editText1 = findViewById(R.id.editText2);
            String text1 = editText1.getText().toString();

// Вставить значение в первую ячейку



            EditText editText2 = findViewById(R.id.editText3);
            String text2 = editText2.getText().toString();

            EditText editText3 = findViewById(R.id.editText4);
            String text3 = editText3.getText().toString();

            EditText editText4 = findViewById(R.id.editText5);
            String text4 = editText4.getText().toString();

            EditText editText5 = findViewById(R.id.editText6);
            String text5 = editText5.getText().toString();

            EditText editText6 = findViewById(R.id.editText7);
            String text6 = editText6.getText().toString();
            EditText editText7 = findViewById(R.id.editText8);
            String text7 = editText7.getText().toString();

// Вставить значение в первую ячейку



            EditText editText8 = findViewById(R.id.editText9);
            String text8 = editText8.getText().toString();

            EditText editText9 = findViewById(R.id.editText10);
            String text9 = editText9.getText().toString();

            EditText editText10 = findViewById(R.id.editText11);
            String text10 = editText10.getText().toString();

            EditText editText11 = findViewById(R.id.editText12);
            String text11 = editText11.getText().toString();



            dataSaver.saveData12(text, text1, text2, text3, text4, text5, text6,text7,text8,text9,text10,text11 );
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }

    }
