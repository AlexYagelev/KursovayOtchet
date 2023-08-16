package com.programmerworld.aspose;


import android.content.Intent;
import android.os.Bundle;
import android.widget.EditText;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.util.ArrayList;
import java.util.List;

public class Activity_type_2b1 extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type2b1);

        findViewById(R.id.button1).setOnClickListener(v -> {

                save();
                Intent intent = new Intent(Activity_type_2b1.this, Activity_type_2_sosudi1.class);
                startActivity(intent);

        });

        findViewById(R.id.button2).setOnClickListener(v -> {

                save();
                Activity_type_2b1.this.finish();

                // Вернуться на предыдущий фрагмент
                getSupportFragmentManager().popBackStack();

        });
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
        ExcelDataSaverAct21b1 dataSaver = new ExcelDataSaverAct21b1("example.xlsx");
        try {
            EditText editText = findViewById(R.id.editText16);
            String text = editText.getText().toString();

            EditText editText1 = findViewById(R.id.editText17);
            String text1 = editText1.getText().toString();

// Вставить значение в первую ячейку



            EditText editText2 = findViewById(R.id.editText18);
            String text2 = editText2.getText().toString();

            EditText editText3 = findViewById(R.id.editText19);
            String text3 = editText3.getText().toString();

            EditText editText4 = findViewById(R.id.editText20);
            String text4 = editText4.getText().toString();

            EditText editText5 = findViewById(R.id.editText21);
            String text5 = editText5.getText().toString();

            EditText editText6 = findViewById(R.id.editText22);
            String text6 = editText6.getText().toString();

// Вставить значение в первую ячейку



            EditText editText7 = findViewById(R.id.editText23);
            String text7 = editText7.getText().toString();

            EditText editText8 = findViewById(R.id.editText24);
            String text8 = editText8.getText().toString();

            EditText editText9 = findViewById(R.id.editText25);
            String text9 = editText9.getText().toString();



            dataSaver.saveData(text, text1, text2, text3, text4, text5, text6, text7, text8, text9);
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }
}
