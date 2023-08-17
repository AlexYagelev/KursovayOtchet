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

public class Activity_type_2b2 extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type2b2);

        findViewById(R.id.button1).setOnClickListener(v -> {

            //save();
            Intent intent = new Intent(Activity_type_2b2.this, Activity_type_2_sosudi1.class);
            startActivity(intent);

        });
        findViewById(R.id.button3).setOnClickListener(v -> {

            save();

            Toast.makeText(this, "Данные сохранены", Toast.LENGTH_SHORT).show();




        });

        findViewById(R.id.button2).setOnClickListener(v -> {

            //save();
            Activity_type_2b2.this.finish();

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
        ExcelDataSaverAct21b2 dataSaver = new ExcelDataSaverAct21b2("example.xlsx");
        try {


            TextView Text = findViewById(R.id.Text0);
            String text_1 = Text.getText().toString();

            //user.setU1((text));

            TextView Text1_1 = findViewById(R.id.Text00);
            String text1_1 = Text1_1.getText().toString();

            TextView Text1_2 = findViewById(R.id.Text000);
            String text1_2 = Text1_2.getText().toString();

            TextView Text1_3 = findViewById(R.id.Text0000);
            String text1_3 = Text1_3.getText().toString();

            User user =new User();
            user.setU17((text_1));
            user.setUU17((text1_1));
            user.setUUU17((text1_2));
            user.setUUUU17((text1_3));

// Вставить значение в первую ячейку

            TextView Text2_1 = findViewById(R.id.Text1);


            //user.setU1((text));

            EditText editText1_1 = findViewById(R.id.editText1_1);
            EditText editText1_2 = findViewById(R.id.editText1_2);
            EditText editText1_3 = findViewById(R.id.editText1_3);
            user.setU1(( Text2_1.getText().toString()));
            user.setUU1(( editText1_1.getText().toString()));
            user.setUUU1((editText1_2.getText().toString()));
            user.setUUUU1((            editText1_3.getText().toString()));


            TextView Text2 = findViewById(R.id.Text2);
            String text2 = Text2.getText().toString();
            user.setU2((text2));
            EditText editText2 = findViewById(R.id.editText2_1);
            String text3 = editText2.getText().toString();
            user.setUU2((text3));
            EditText editText2_2 = findViewById(R.id.editText2_2);
            EditText editText2_3 = findViewById(R.id.editText2_3);
            user.setUUU2((editText2_2.getText().toString()));
            user.setUUUU2((editText2_3.getText().toString()));



            TextView Text3 = findViewById(R.id.Text3);
            String text5 = Text3.getText().toString();
            user.setU3((text5));
            EditText editText3 = findViewById(R.id.editText3_1);
            String text6 = editText3.getText().toString();
            user.setUU3(String.valueOf(text6));
            EditText editText3_2 = findViewById(R.id.editText3_2);
            EditText editText3_3 = findViewById(R.id.editText3_3);
            user.setUUU3((editText3_2.getText().toString()));
            user.setUUUU3((editText3_3.getText().toString()));


            TextView Text4 = findViewById(R.id.Text4);
            user.setU4(Text4.getText().toString());
            EditText editText4 = findViewById(R.id.editText4_1);
            user.setUU4(editText4.getText().toString());
            EditText editText4_2 = findViewById(R.id.editText4_2);
            EditText editText4_3 = findViewById(R.id.editText4_3);
            user.setUUU4((editText4_2.getText().toString()));
            user.setUUUU4((editText4_3.getText().toString()));



            TextView Text5 = findViewById(R.id.Text5);
            user.setU5(Text5.getText().toString());
            EditText editText5 = findViewById(R.id.editText5_1);
            user.setUU5(editText5.getText().toString());
            EditText editText5_2 = findViewById(R.id.editText5_2);
            EditText editText5_3 = findViewById(R.id.editText5_3);
            user.setUUU5((editText5_2.getText().toString()));
            user.setUUUU5((editText5_3.getText().toString()));



            TextView Text6 = findViewById(R.id.Text6);
            user.setU6(Text6.getText().toString());
            EditText editText6 = findViewById(R.id.editText6_1);
            user.setUU6(editText6.getText().toString());
            EditText editText6_2 = findViewById(R.id.editText6_2);
            EditText editText6_3 = findViewById(R.id.editText6_3);
            user.setUUU6((editText6_2.getText().toString()));
            user.setUUUU6((editText6_3.getText().toString()));


// Вставить значение в первую ячейку



            TextView Text7 = findViewById(R.id.Text7);
            user.setU7(Text7.getText().toString());
            EditText editText7 = findViewById(R.id.editText7_1);
            user.setUU7(editText7.getText().toString());
            EditText editText7_2 = findViewById(R.id.editText7_2);
            EditText editText7_3 = findViewById(R.id.editText7_3);
            user.setUUU7((editText7_2.getText().toString()));
            user.setUUUU7((editText7_3.getText().toString()));


            TextView Text8 = findViewById(R.id.Text8);
            user.setU8(Text8.getText().toString());
            EditText editText8 = findViewById(R.id.editText8_1);
            user.setUU8(editText8.getText().toString());
            EditText editText8_2 = findViewById(R.id.editText8_2);
            EditText editText8_3 = findViewById(R.id.editText8_3);
            user.setUUU8((editText8_2.getText().toString()));
            user.setUUUU8((editText8_3.getText().toString()));



            TextView Text9 = findViewById(R.id.Text9);
            user.setU9(Text9.getText().toString());
            EditText editText9 = findViewById(R.id.editText9_1);
            user.setUU9(editText9.getText().toString());
            EditText editText9_2 = findViewById(R.id.editText9_2);
            EditText editText9_3 = findViewById(R.id.editText9_3);
            user.setUUU9((editText9_2.getText().toString()));
            user.setUUUU9((editText9_3.getText().toString()));



            TextView Text10 = findViewById(R.id.Text10);
            user.setU10(Text10.getText().toString());
            EditText editText10 = findViewById(R.id.editText10_1);
            user.setUU10(editText10.getText().toString());
            EditText editText10_2 = findViewById(R.id.editText10_2);
            EditText editText10_3 = findViewById(R.id.editText10_3);
            user.setUUU10((editText10_2.getText().toString()));
            user.setUUUU10((editText10_3.getText().toString()));



            TextView Text11 = findViewById(R.id.Text11);
            user.setU11(Text11.getText().toString());
            EditText editText11 = findViewById(R.id.editText11_1);
            user.setUU11(editText11.getText().toString());
            EditText editText11_2 = findViewById(R.id.editText11_2);
            EditText editText11_3 = findViewById(R.id.editText11_3);
            user.setUUU11((editText11_2.getText().toString()));
            user.setUUUU11((editText11_3.getText().toString()));



            TextView Text12 = findViewById(R.id.Text12);
            user.setU11(Text12.getText().toString());
            EditText editText12 = findViewById(R.id.editText12_1);
            user.setUU11(editText12.getText().toString());
            EditText editText12_2 = findViewById(R.id.editText12_2);
            EditText editText12_3 = findViewById(R.id.editText12_3);
            user.setUUU12((editText12_2.getText().toString()));
            user.setUUUU12((editText12_3.getText().toString()));


            TextView Text13 = findViewById(R.id.Text13);
            user.setU11(Text13.getText().toString());
            EditText editText13 = findViewById(R.id.editText13_1);
            user.setUU11(editText13.getText().toString());
            EditText editText13_2 = findViewById(R.id.editText13_2);
            EditText editText13_3 = findViewById(R.id.editText13_3);
            user.setUUU13((editText13_2.getText().toString()));
            user.setUUUU13((editText13_3.getText().toString()));




            TextView Text14 = findViewById(R.id.Text14);
            user.setU11(Text14.getText().toString());
            EditText editText14 = findViewById(R.id.editText14_1);
            user.setUU11(editText14.getText().toString());
            EditText editText14_2 = findViewById(R.id.editText14_2);
            EditText editText14_3 = findViewById(R.id.editText14_3);
            user.setUUU14((editText14_2.getText().toString()));
            user.setUUUU14((editText14_3.getText().toString()));


            TextView Text15 = findViewById(R.id.Text15);
            user.setU11(Text15.getText().toString());
            EditText editText15 = findViewById(R.id.editText15_1);
            user.setUU11(editText15.getText().toString());
            EditText editText15_2 = findViewById(R.id.editText15_2);
            EditText editText15_3 = findViewById(R.id.editText15_3);
            user.setUUU15((editText15_2.getText().toString()));
            user.setUUUU15((editText15_3.getText().toString()));

            TextView Text16 = findViewById(R.id.Text16);
            user.setU11(Text16.getText().toString());
            EditText editText16 = findViewById(R.id.editText16_1);
            user.setUU11(editText16.getText().toString());
            EditText editText16_2 = findViewById(R.id.editText16_2);
            EditText editText16_3 = findViewById(R.id.editText16_3);
            user.setUUU16((editText16_2.getText().toString()));
            user.setUUUU16((editText16_3.getText().toString()));






            dataSaver.saveData(user);
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }
}
