package com.programmerworld.aspose;


import android.content.Context;
import android.content.Intent;
import android.os.Bundle;
import android.view.MotionEvent;
import android.view.View;
import android.view.inputmethod.InputMethodManager;
import android.widget.EditText;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.programmerworld.aspose.sosudi.Activity_type_2b;

import java.util.ArrayList;
import java.util.List;

public class Activity_type_2 extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type2);

        View rootView = getWindow().getDecorView().getRootView();
        rootView.setOnTouchListener(new View.OnTouchListener() {
            @Override
            public boolean onTouch(View v, MotionEvent event) {
                hideKeyboard();
                return false;
            }


        });
        findViewById(R.id.button3).setOnClickListener(v -> {

            save();save1();


        });
        findViewById(R.id.button1).setOnClickListener(v -> {


                Intent intent = new Intent(Activity_type_2.this, Activity_type_2b.class);
                startActivity(intent);
            // Activity_type_2.this.finish();
        });

        findViewById(R.id.button2).setOnClickListener(v -> {


                Activity_type_2.this.finish();

                // Вернуться на предыдущий фрагмент
                getSupportFragmentManager().popBackStack();

        });
    }
    private void hideKeyboard() {
        View view = getCurrentFocus();
        if (view != null) {
            InputMethodManager imm = (InputMethodManager) getSystemService(Context.INPUT_METHOD_SERVICE);
            imm.hideSoftInputFromWindow(view.getWindowToken(), 0);
        }
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
        ExcelDataSaverAct21 dataSaver = new ExcelDataSaverAct21("example.xlsx");
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

            dataSaver.saveData(text, text1, text2, text3, text4, text5, text6);


            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }
    private void save1(){
        ExcelDataSaverAct21b dataSaver = new ExcelDataSaverAct21b("example.xlsx");
        try {
            EditText editText = findViewById(R.id.editText26);
            String text = editText.getText().toString();

            EditText editText1 = findViewById(R.id.editText27);
            String text1 = editText1.getText().toString();

// Вставить значение в первую ячейку



            EditText editText2 = findViewById(R.id.editText28);
            String text2 = editText2.getText().toString();

            EditText editText3 = findViewById(R.id.editText29);
            String text3 = editText3.getText().toString();

            EditText editText4 = findViewById(R.id.editText30);
            String text4 = editText4.getText().toString();



            dataSaver.saveData(text, text1, text2, text3, text4);
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }
}
