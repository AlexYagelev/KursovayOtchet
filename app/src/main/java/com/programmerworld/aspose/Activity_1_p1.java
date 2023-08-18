package com.programmerworld.aspose;


import android.content.Intent;
import android.os.Bundle;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.util.ArrayList;
import java.util.List;

public class Activity_1_p1 extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_activity1_p1);

        findViewById(R.id.button1).setOnClickListener(v -> {

            //save();
            Intent intent = new Intent(Activity_1_p1.this, Activity_1_p2.class);
            startActivity(intent);
            // Activity_type_2b4.this.finish();
        });
        findViewById(R.id.button3).setOnClickListener(v -> {

            save();

            Toast.makeText(this, "Данные сохранены", Toast.LENGTH_SHORT).show();




        });

        findViewById(R.id.button2).setOnClickListener(v -> {

            //save();
            Activity_1_p1.this.finish();

            // Вернуться на предыдущий фрагмент
            getSupportFragmentManager().popBackStack();

        });
    }

    private void clearFields() {

        EditText editText1 = findViewById(R.id.editText1);
        editText1.setText("");
        EditText editText2 = findViewById(R.id.editText2);
        editText2.setText("");
        EditText editText3 = findViewById(R.id.editText3);
        editText3.setText("");
        EditText editText4 = findViewById(R.id.editText4);
        editText4.setText("");
        EditText editText5 = findViewById(R.id.editText5);

        editText5.setText("");
        EditText editText6 = findViewById(R.id.editText6);

        editText6.setText("");
        EditText editText7 = findViewById(R.id.editText7);

        editText7.setText("");
        EditText editText8 = findViewById(R.id.editText8);

        editText8.setText("");
        EditText editText9 = findViewById(R.id.editText9);

        editText9.setText("");
        EditText editText10 = findViewById(R.id.editText10);

        editText10.setText("");
        EditText editText11 = findViewById(R.id.editText11);

        editText11.setText("");

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
        ExcelDataSaverAct21d dataSaver = new ExcelDataSaverAct21d("example.xlsx");
        try {


            TextView Text = findViewById(R.id.Text1);
            String text ="#1_2_1_1"+ Text.getText().toString() ;

            //user.setU1((text));

            EditText editText1 = findViewById(R.id.editText1);
            String text1 = editText1.getText().toString();

            User user =new User();
            user.setU1((text));
            user.setUU1((text1));

// Вставить значение в первую ячейку

            TextView Text2 = findViewById(R.id.Text2);
            String text2 = Text2.getText().toString();
            user.setU2((text2));


            EditText editText2 = findViewById(R.id.editText2);
            String text3 = editText2.getText().toString();
            user.setUU2((text3));

            TextView Text3 =findViewById(R.id.Text0);
            String text5 ="#1_2_1_2"+  Text3.getText().toString();
            user.setU3((text5));


            EditText editText3 = findViewById(R.id.editText3);
            String text6 = editText3.getText().toString();
            user.setUU3(String.valueOf(text6));






            dataSaver.saveData2_3(user);
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }
}
