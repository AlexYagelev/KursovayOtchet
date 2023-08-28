package com.programmerworld.aspose.TPA.pr1;

import android.content.Intent;
import android.os.Bundle;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.programmerworld.aspose.ExcelDataSaverAct21d;
import com.programmerworld.aspose.R;
import com.programmerworld.aspose.User;

import java.util.ArrayList;
import java.util.List;


public class path1  extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_pr1);

        findViewById(R.id.button1).setOnClickListener(v -> {

            //save();
            Intent intent = new Intent(this, path2.class);
            startActivity(intent);
            //     Activity_type_2b.this.finish();
        });
        findViewById(R.id.button3).setOnClickListener(v -> {

            save();


        });

        findViewById(R.id.button2).setOnClickListener(v -> {

            // save();
            this.finish();

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

            TextView Text6_1 = findViewById(R.id.Text6);
            user.setU7((Text6_1.getText().toString()));
            EditText editText6_2 = findViewById(R.id.editText6);
            user.setUU7(  editText6_2.getText().toString());
            EditText editText6_3 = findViewById(R.id.editText66);
            user.setUUU7(  editText6_3.getText().toString());

            TextView Text7_1 = findViewById(R.id.Text7);
            user.setU8((Text7_1.getText().toString()));
            EditText editText7_2 = findViewById(R.id.editText7);
            user.setUU8(  editText7_2.getText().toString());
            EditText editText7_3 = findViewById(R.id.editText77);
            user.setUUU8(  editText7_3.getText().toString());

            TextView Text8_1 = findViewById(R.id.Text8);
            user.setU9((Text8_1.getText().toString()));
            EditText editText8_2 = findViewById(R.id.editText8);
            user.setUU9(  editText8_2.getText().toString());
            EditText editText8_3 = findViewById(R.id.editText88);
            user.setUUU9(  editText8_3.getText().toString());

            TextView Text9_1 = findViewById(R.id.Text9);
            user.setU10((Text9_1.getText().toString()));
            EditText editText9_2 = findViewById(R.id.editText9);
            user.setUU10(  editText9_2.getText().toString());
            EditText editText9_3 = findViewById(R.id.editText99);
            user.setUUU10(  editText9_3.getText().toString());


            TextView Text10_1 = findViewById(R.id.Text10);
            user.setU11((Text10_1.getText().toString()));
            EditText editText10_2 = findViewById(R.id.editText10);
            user.setUU11(  editText10_2.getText().toString());
            EditText editText10_3 = findViewById(R.id.editText1010);
            user.setUUU11(  editText10_3.getText().toString());

            dataSaver.saveData2_10_TPA_1(user);
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }
}
