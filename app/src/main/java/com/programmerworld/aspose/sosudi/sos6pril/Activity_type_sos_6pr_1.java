


package com.programmerworld.aspose.sosudi.sos6pril;


import android.content.Intent;
import android.os.Bundle;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import com.programmerworld.aspose.ExcelDataSaverAct111;
import com.programmerworld.aspose.ExcelDataSaverAct21d;
import com.programmerworld.aspose.R;
import com.programmerworld.aspose.User;
import com.programmerworld.aspose.sosudi.sos5pril.Activity_type_sos_5pr_2;

public class Activity_type_sos_6pr_1 extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type_sos6pr1);

        findViewById(R.id.button1).setOnClickListener(v -> {

            //save();
            Intent intent = new Intent(Activity_type_sos_6pr_1.this, Activity_type_sos_5pr_2.class);
            startActivity(intent);
            //   Activity_type_2b1.this.finish();
        });
        findViewById(R.id.button3).setOnClickListener(v -> {

            save();

            Toast.makeText(this, "Данные сохранены", Toast.LENGTH_SHORT).show();




        });

        findViewById(R.id.button2).setOnClickListener(v -> {

            //save();
            Activity_type_sos_6pr_1.this.finish();

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






    private void save(){
        ExcelDataSaverAct21d dataSaver = new ExcelDataSaverAct21d("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        try {


            TextView Text = findViewById(R.id.Text1);
            String text = Text.getText().toString();

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




            dataSaver.saveData2_2_6pr(user);clearFields();
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }
    }
}
