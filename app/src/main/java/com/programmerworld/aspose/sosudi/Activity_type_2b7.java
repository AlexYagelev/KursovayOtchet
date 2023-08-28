package com.programmerworld.aspose.sosudi;


import android.content.Intent;
import android.os.Bundle;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.programmerworld.aspose.Activity_type_SHB1;
import com.programmerworld.aspose.ExcelDataSaverAct21d;
import com.programmerworld.aspose.R;
import com.programmerworld.aspose.User;

import java.util.ArrayList;
import java.util.List;

public class Activity_type_2b7 extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type2b7);

        findViewById(R.id.button1).setOnClickListener(v -> {

                    //save();
                Intent intent = new Intent(this, Activity_type_SHB1.class);
            intent.setFlags(Intent.FLAG_ACTIVITY_SINGLE_TOP);
          startActivity(intent);
            //   Activity_type_2b7.this.finish();
                });
        findViewById(R.id.button3).setOnClickListener(v -> {

            save();

            Toast.makeText(this, "Данные сохранены", Toast.LENGTH_SHORT).show();




        });

        findViewById(R.id.button2).setOnClickListener(v -> {

            //save();
            Activity_type_2b7.this.finish();

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
            User user =new User();

            TextView Text = findViewById(R.id.Text0);
            String text = Text.getText().toString();
            TextView Text1_1 = findViewById(R.id.Text1);
            String text1 = Text1_1.getText().toString();

            //user.setU1((text));



            user.setU1((text));
            user.setU2((text1));

            TextView Text1_2 = findViewById(R.id.Text00);
            user.setUU1( Text1_2.getText().toString());
            EditText editText1_2 = findViewById(R.id.editText1);
            user.setUU2(  editText1_2.getText().toString());

            TextView Text1_3 = findViewById(R.id.Text000);
            user.setUUU1( Text1_3.getText().toString());
            EditText editText1_3 = findViewById(R.id.editText111);
            user.setUUU2(  editText1_3.getText().toString());




            //user.setU1((text));



            TextView Text2_1 = findViewById(R.id.Text2);
            user.setU3((Text2_1.getText().toString()));
            EditText editText2_2 = findViewById(R.id.editText2);
            user.setUU3(  editText2_2.getText().toString());
            EditText editText2_3 = findViewById(R.id.editText22);
            user.setUUU3(  editText2_3.getText().toString());

            TextView Text3_1 = findViewById(R.id.Text3);
            user.setU4((Text3_1.getText().toString()));
            EditText editText3_2 = findViewById(R.id.editText3);
            user.setUU4(  editText3_2.getText().toString());
            EditText editText3_3 = findViewById(R.id.editText33);
            user.setUUU4(  editText3_3.getText().toString());


            TextView Text4_1 = findViewById(R.id.Text4);
            user.setU5((Text4_1.getText().toString()));
            EditText editText4_2 = findViewById(R.id.editText4);
            user.setUU5(  editText4_2.getText().toString());
            EditText editText4_3 = findViewById(R.id.editText44);
            user.setUUU5(  editText4_3.getText().toString());

            TextView Text5_1 = findViewById(R.id.Text5);
            user.setU6((Text5_1.getText().toString()));
            EditText editText5_2 = findViewById(R.id.editText5);
            user.setUU6(  editText5_2.getText().toString());
            EditText editText5_3 = findViewById(R.id.editText55);
            user.setUUU6(  editText5_3.getText().toString());



            dataSaver.saveData2_5(user);
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }
}
