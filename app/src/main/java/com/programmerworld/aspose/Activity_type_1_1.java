

package com.programmerworld.aspose;

        import android.Manifest;
        import android.content.Intent;
        import android.content.pm.PackageManager;
        import android.os.Bundle;
        import android.widget.EditText;
        import android.widget.Toast;

        import androidx.appcompat.app.AppCompatActivity;
        import androidx.core.app.ActivityCompat;
        import androidx.core.content.ContextCompat;

        import com.aspose.cells.Workbook;

public class Activity_type_1_1 extends AppCompatActivity {
    private static final int REQUEST_WRITE_EXTERNAL_STORAGE = 1;private Workbook workbook;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type11);
        findViewById(R.id.b2).setOnClickListener(v -> {
            if (ContextCompat.checkSelfPermission(Activity_type_1_1.this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    == PackageManager.PERMISSION_GRANTED) {
                save_1();
                Intent intent1 = new Intent(Activity_type_1_1.this, MainActivity2.class);
                //intent1.putExtra("uri", uri.toString());

                startActivity(intent1);
            } else {
                // Запрос разрешения
                ActivityCompat.requestPermissions(Activity_type_1_1.this,
                        new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},
                        REQUEST_WRITE_EXTERNAL_STORAGE);
            }
        });

        findViewById(R.id.button4).setOnClickListener(v -> {

            Intent intent2 = new Intent(Activity_type_1_1.this, Activity_type_1.class);
            //intent1.putExtra("uri", uri.toString());

            startActivity(intent2);

        });

    }


    private void save_1(){
        ExcelDataSaverAct11 dataSaver_1 = new ExcelDataSaverAct11("example.xlsx");
        try {



            EditText editText11 = findViewById(R.id.editText_11);
            String text11 = editText11.getText().toString();

            EditText editText12 = findViewById(R.id.editText_21);
            String text12 = editText12.getText().toString();

            EditText editText13 = findViewById(R.id.editText_31);
            String text13 = editText13.getText().toString();

            EditText editText14 = findViewById(R.id.editText_41);
            String text14 = editText14.getText().toString();

            EditText editText15 = findViewById(R.id.editText_51);
            String text15 = editText15.getText().toString();

            EditText editText16 = findViewById(R.id.editText_61);
            String text16 = editText16.getText().toString();




            // Вставить значение в первую ячейку




            dataSaver_1.saveData_1(text11, text12, text13, text14, text15, text16);
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }

    }
}