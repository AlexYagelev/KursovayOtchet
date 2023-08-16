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

public class Activity_type_2b extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type2b);

        findViewById(R.id.button1).setOnClickListener(v -> {

                //save();
                Intent intent = new Intent(Activity_type_2b.this, Activity_type_2b1.class);
                startActivity(intent);

        });
        findViewById(R.id.button3).setOnClickListener(v -> {

            save();


        });

        findViewById(R.id.button2).setOnClickListener(v -> {

               // save();
                Activity_type_2b.this.finish();

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
        ExcelDataSaverAct21c dataSaver = new ExcelDataSaverAct21c("example.xlsx");
        try {
            TextView Text = findViewById(R.id.Text0);
            String i1 = Text.getText().toString();

            TextView Text1 = findViewById(R.id.Text00);
            String i11 = Text1.getText().toString();

// Вставить значение в первую ячейку



            TextView Text2 = findViewById(R.id.Text000);
            String i111 = Text2.getText().toString();

            TextView Text3 = findViewById(R.id.Text1);
            String a = Text3.getText().toString();

            EditText editText3 = findViewById(R.id.editText1);
            String aa = editText3.getText().toString();

            EditText editText4 = findViewById(R.id.editText111);
            String aaa = editText4.getText().toString();


            TextView Text5 = findViewById(R.id.Text2);
            String b = Text5.getText().toString();

            EditText editText6 = findViewById(R.id.editText2);
            String bb = editText6.getText().toString();

            EditText editText7 = findViewById(R.id.editText22);
            String bbb = editText7.getText().toString();



            TextView Text8 = findViewById(R.id.Text3);
            String c = Text8.getText().toString();

            EditText editText9 = findViewById(R.id.editText3);
            String cc = editText9.getText().toString();

            EditText editText10 = findViewById(R.id.editText33);
            String ccc = editText10.getText().toString();


            TextView Text10 = findViewById(R.id.Text4);
            String d = Text10.getText().toString();

            EditText editText11 = findViewById(R.id.editText4);
            String dd = editText11.getText().toString();

            EditText editText12 = findViewById(R.id.editText44);
            String ddd = editText12.getText().toString();



            TextView Text13 = findViewById(R.id.Text5);
            String e = Text13.getText().toString();

            EditText editText14 = findViewById(R.id.editText5);
            String ee = editText14.getText().toString();

            EditText editText15 = findViewById(R.id.editText55);
            String eee = editText15.getText().toString();


            TextView Text16 = findViewById(R.id.Text6);
            String f = Text16.getText().toString();

            EditText editText17 = findViewById(R.id.editText6);
            String ff = editText17.getText().toString();

            EditText editText18 = findViewById(R.id.editText66);
            String fff = editText18.getText().toString();




            TextView Text19 = findViewById(R.id.Text7);
            String g = Text19.getText().toString();

            EditText editText20 = findViewById(R.id.editText7);
            String gg = editText20.getText().toString();

            EditText editText21 = findViewById(R.id.editText77);
            String ggg = editText21.getText().toString();


            TextView Text22 = findViewById(R.id.Text8);
            String h = Text22.getText().toString();

            EditText editText23 = findViewById(R.id.editText8);
            String hh = editText23.getText().toString();

            EditText editText24 = findViewById(R.id.editText88);
            String hhh = editText24.getText().toString();



            TextView Text25 = findViewById(R.id.Text9);
            String i = Text25.getText().toString();

            EditText editText26 = findViewById(R.id.editText9);
            String ii = editText26.getText().toString();

            EditText editText27 = findViewById(R.id.editText99);
            String iii = editText27.getText().toString();


            TextView Text28 = findViewById(R.id.Text10);
            String j = Text28.getText().toString();

            EditText editText29 = findViewById(R.id.editText10);
            String jj = editText29.getText().toString();

            EditText editText30 = findViewById(R.id.editText1010);
            String jjj = editText30.getText().toString();

            TextView Text31 = findViewById(R.id.Text11);
            String k = Text31.getText().toString();

            EditText editText32 = findViewById(R.id.editText11);
            String kk = editText32.getText().toString();

            EditText editText33 = findViewById(R.id.editText1111);
            String kkk = editText33.getText().toString();



            TextView Text34 = findViewById(R.id.Text12);
            String l = Text34.getText().toString();

            EditText editText35 = findViewById(R.id.editText12);
            String ll = editText35.getText().toString();

            EditText editText36 = findViewById(R.id.editText1212);
            String lll = editText36.getText().toString();



            dataSaver.saveData( i1,  i11,  i111, a,  aa,  aaa,  b,  bb,  bbb,  c,  cc, ccc,  d,  dd, ddd,  e,  ee, eee,  f,  ff, fff,  g,  gg, ggg,  h,  hh, hhh,  i,  ii,  iii, j,  jj, jjj,  k,  kk, kkk,  l,  ll, lll);

            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }
}
