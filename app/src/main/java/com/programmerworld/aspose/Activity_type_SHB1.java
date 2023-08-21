package com.programmerworld.aspose;

import android.content.Intent;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;

import androidx.appcompat.app.AppCompatActivity;

import com.programmerworld.aspose.sos3pril.Activity_type_sos_3pr_1;
import com.programmerworld.aspose.sos4pril.Activity_type_sos_4pr_1;


public class Activity_type_SHB1  extends AppCompatActivity {


    private Button button;
    private MainViewModel viewModel;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type_shb);


        button = findViewById(R.id.button1);




        button.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                Intent intent = new Intent(Activity_type_SHB1.this, Activity_type_2.class);
                startActivity(intent);

            }
        });
        findViewById(R.id.button2).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                Intent intent = new Intent(Activity_type_SHB1.this, Activity_1_p1.class);
                startActivity(intent);

            }
        });


        findViewById(R.id.button3).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                Intent intent = new Intent(Activity_type_SHB1.this, Activity_type_sos_3pr_1.class);
                startActivity(intent);

            }
        });findViewById(R.id.button4).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                Intent intent = new Intent(Activity_type_SHB1.this, Activity_type_sos_4pr_1.class);
                startActivity(intent);

            }
        });


    }}
