package com.programmerworld.aspose;

import android.content.Intent;
import android.os.Bundle;
import android.widget.CheckBox;

import androidx.appcompat.app.AppCompatActivity;

public class Activity_type_SHB1 extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type_shb);
        findViewById(R.id.button1).setOnClickListener(v -> {


            Intent intent = new Intent(Activity_type_SHB1.this, Activity_type_2.class);
            startActivity(intent);
            CheckBox checkbox = findViewById(R.id.checkbox1);
            checkbox.setChecked(true);

        });
    }
}