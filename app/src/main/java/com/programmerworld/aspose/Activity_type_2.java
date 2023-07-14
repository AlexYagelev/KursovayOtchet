package com.programmerworld.aspose;


import android.content.pm.PackageManager;
import android.os.Bundle;
import android.widget.EditText;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.ContextCompat;

public class Activity_type_2 extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type2);

        findViewById(R.id.button1).setOnClickListener(v -> {
            if (ContextCompat.checkSelfPermission(Activity_type_2.this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    == PackageManager.PERMISSION_GRANTED) {
                save();
            }
        });
    }


    private void save() {
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
}
