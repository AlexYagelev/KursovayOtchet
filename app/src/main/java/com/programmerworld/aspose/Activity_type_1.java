package com.programmerworld.aspose;

import android.os.Bundle;

import androidx.appcompat.app.AppCompatActivity;

import com.aspose.cells.Workbook;

public class Activity_type_1 extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type1);
        Workbook workbook = (Workbook) getIntent().getSerializableExtra("workbook");
    }
}