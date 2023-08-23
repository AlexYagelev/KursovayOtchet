package com.programmerworld.aspose;

import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.Color;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.view.View;
import android.widget.Button;

import androidx.appcompat.app.AppCompatActivity;

import com.programmerworld.aspose.sos3pril.Activity_type_sos_3pr_1;
import com.programmerworld.aspose.sos4pril.Activity_type_sos_4pr_1;
import com.programmerworld.aspose.sos5pril.Activity_type_sos_5pr_1;
import com.programmerworld.aspose.sos6pril.Activity_type_sos_6pr_1;
import com.programmerworld.aspose.sos7pril.Activity_type_sos_7pr_1;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
public class Activity_type_SHB1  extends AppCompatActivity {


    private Button button;
    private MainViewModel viewModel;
    private static final int CAMERA_REQUEST = 1;

    private Button button3;



    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        setContentView(R.layout.activity_type_shb);
        findViewById(R.id.button1).setBackgroundColor(Color.parseColor("#0079C2"));
        findViewById(R.id.button2).setBackgroundColor(Color.parseColor("#0079C2"));
        findViewById(R.id.button3).setBackgroundColor(Color.parseColor("#0079C2"));
        findViewById(R.id.button4).setBackgroundColor(Color.parseColor("#0079C2"));
        findViewById(R.id.button5).setBackgroundColor(Color.parseColor("#0079C2"));
        findViewById(R.id.button6).setBackgroundColor(Color.parseColor("#0079C2"));
        findViewById(R.id.button7).setBackgroundColor(Color.parseColor("#0079C2"));
        // findViewById(R.id.button8).setBackgroundColor(Color.parseColor("#0079C2"));

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

        int photoNum = 1;
        findViewById(R.id.button3).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                Intent intent = new Intent(Activity_type_SHB1.this, Activity_type_sos_3pr_1.class);
                startActivity(intent);

            }
        });
        findViewById(R.id.button4).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                Intent intent = new Intent(Activity_type_SHB1.this, Activity_type_sos_4pr_1.class);
                startActivity(intent);

            }
        });

        findViewById(R.id.button5).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                Intent intent = new Intent(Activity_type_SHB1.this, Activity_type_sos_5pr_1.class);
                startActivity(intent);

            }
        });
        findViewById(R.id.button6).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                Intent intent = new Intent(Activity_type_SHB1.this, Activity_type_sos_6pr_1.class);
                startActivity(intent);

            }
        });
        findViewById(R.id.button7).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                Intent intent = new Intent(Activity_type_SHB1.this, Activity_type_sos_7pr_1.class);
                startActivity(intent);

            }
        });



            button3 = findViewById(R.id.button8);

           /* button3.setOnClickListener(new View.OnClickListener() {
                @Override
                public void onClick(View v) {
                    //Intent cameraIntent = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);

                    Intent intent;
                    intent = new Intent(Activity_type_SHB1.this,MainActivity2.class);
                    startActivity(intent);

                    //startActivityForResult(cameraIntent, CAMERA_REQUEST);

                }
            });*/
        button3 = findViewById(R.id.button8);

       button3.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                openCamera();
            }
        });


    }    private void openCamera() {
        Intent cameraIntent = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
        startActivityForResult(cameraIntent, CAMERA_REQUEST);
    }
    public int photoNum = 1;
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {

        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == CAMERA_REQUEST && resultCode == RESULT_OK) {

            String photoName;
            photoNum++;
            photoName = "photo_" + photoNum + ".jpg";

            // сохранение фото с именем photoName


            if (requestCode == CAMERA_REQUEST && resultCode == RESULT_OK) {
                Bitmap photo = (Bitmap) data.getExtras().get("data");

                // Исправил ошибку в имени директории
                File picturesDirectory = getExternalFilesDir(Environment.DIRECTORY_PICTURES);

                File imageFile = new File("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString+ "/" + photoName );

                try {
                    OutputStream stream = new FileOutputStream(imageFile);
                    photo.compress(Bitmap.CompressFormat.JPEG, 100, stream);
                    stream.flush();
                      stream.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }

            photoNum++;
        }
    }







}
    /*private void openCamera() {
        Intent cameraIntent = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
        startActivityForResult(cameraIntent, CAMERA_REQUEST);
    }
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        int photosTaken = 0;
        // Добавил проверку на resultCode
        if (requestCode == CAMERA_REQUEST && resultCode == RESULT_OK) {
            Bitmap photo = (Bitmap) data.getExtras().get("data");

            // Исправил ошибку в имени директории
            File picturesDirectory = getExternalFilesDir(Environment.DIRECTORY_PICTURES);
            File imageFile = new File("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString+ "/" + "data.jpg" );

            try {
                OutputStream stream = new FileOutputStream(imageFile);
                photo.compress(Bitmap.CompressFormat.JPEG, 100, stream);
                stream.flush();
                stream.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
       }*/
