package com.programmerworld.aspose;

import android.Manifest;
import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.Settings;
import android.util.Log;
import android.view.MotionEvent;
import android.view.View;
import android.view.inputmethod.InputMethodManager;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.AutoCompleteTextView;
import android.widget.Button;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.FileNotFoundException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.stream.Collectors;

public class Activity_type_1_2 extends AppCompatActivity {
    private static final int REQUEST_CODE = 100 ;
    private Spinner eoSpinner;
    private Spinner objectNameSpinner;
    private Spinner objectSpinner;
    private Spinner workSpinner;
    private Spinner dataSpinner;
    private static final int REQUEST_READ_STORAGE = 100;
private String tag = "Жизненный цикл";
    private List<String> eoValues = new ArrayList<>();
    private List<String> objectNameValues = new ArrayList<>();
    private List<String> objectValues = new ArrayList<>();
    private List<String> workValues = new ArrayList<>();
    private List<String> dataValues = new ArrayList<>();
    private AutoCompleteTextView autoCompleteTextView;
    private Spinner spinner;
    private Spinner spinner6;

    private Button button2;
    private Spinner spinner2;
    private Spinner spinner3;
    private Spinner spinner33;
    private Spinner spinner4;
    private Spinner spinner5;
    private TextView tvPermission;
    private Button btnPermission;
    private static final int PERMISSION_STORAGE = 101;
    String ST;
    private DeviceExplorerViewImpl explorerView;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.ak);
Log.i(tag, "onCreate()");
        View rootView = getWindow().getDecorView().getRootView();
        rootView.setOnTouchListener(new View.OnTouchListener() {
            @Override
            public boolean onTouch(View v, MotionEvent event) {
                hideKeyboard();
                return false;
            }
        });
         tvPermission = findViewById(R.id.tvPermission);
         btnPermission = findViewById(R.id.btnPermission);
        if (PermissionUtils.hasPermissions(this)) {
            tvPermission.setText("Разрешение получено");
        } else {
            tvPermission.setText("Разрешение не предоставлено");
        }btnPermission.setOnClickListener(new View.OnClickListener() {
            @Override public void onClick(View v) {
                if (PermissionUtils.hasPermissions(Activity_type_1_2.this)) return;
                PermissionUtils.requestPermissions(Activity_type_1_2.this, PERMISSION_STORAGE);
            }
        });



        try {explorerView = new DeviceExplorerViewImpl(this);

            // Создаем объект для передачи в метод
            TransferSupport support = new TransferSupport();

            // Вызываем метод импорта данных
            explorerView.importData(support);
              clip();
            spinner = findViewById(R.id.spinner8);
            spinner2 = findViewById(R.id.spinner2);
            clip3();

            spinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                @Override
                public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                    spinner2.setEnabled(position != 0);
                }

                @Override
                public void onNothingSelected(AdapterView<?> parent) {
                }
            });

            clip4();
            clip5();
            clip6();
            clip7();
            clip8();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    @Override
    protected void onStart() {
        super.onStart();

        Toast.makeText(getApplicationContext(), "onStart()", Toast.LENGTH_SHORT).show();
        Log.i(tag, "onStart()");
    }

    @Override
    protected void onResume() {
        super.onResume();

        Toast.makeText(getApplicationContext(), "onResume()", Toast.LENGTH_SHORT).show();
        Log.i(tag, "onResume()");
    }

    @Override
    protected void onPause() {
        super.onPause();

        Toast.makeText(getApplicationContext(), "onPause()", Toast.LENGTH_SHORT).show();
        Log.i(tag, "onPause()");
    }

    @Override
    protected void onStop() {
        super.onStop();

        Toast.makeText(getApplicationContext(), "onStop()", Toast.LENGTH_SHORT).show();
        Log.i(tag, "onStop()");
    }

    @Override
    protected void onRestart() {
        super.onRestart();

        Toast.makeText(getApplicationContext(), "onRestart()", Toast.LENGTH_SHORT).show();
        Log.i(tag, "onRestart()");
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();

        Toast.makeText(getApplicationContext(), "onDestroy()", Toast.LENGTH_SHORT).show();
        Log.i(tag, "onDestroy()");
    }

    private void hideKeyboard() {
        View view = getCurrentFocus();
        if (view != null) {
            InputMethodManager imm = (InputMethodManager) getSystemService(Context.INPUT_METHOD_SERVICE);
            imm.hideSoftInputFromWindow(view.getWindowToken(), 0);
        }
    }


    public class DeviceExplorerViewImpl {Context context;

        public DeviceExplorerViewImpl(Context context) {
            this.context = context;
        }

        private void showError(String error) {
            Toast.makeText(context, error, Toast.LENGTH_LONG).show();
        }

        // Метод импорта данных
        public void importData(TransferSupport support) {

            // Проверяем, поддерживает ли компонент перетаскивание
            if(!support.isDrop()) {
                // Если нет - выходим из метода
                return;
            }

            try {
                // Получаем информацию о перетаскивании
                DropLocation dropLocation = support.getDropLocation();

                // Дальнейший код импорта данных
                // используя dropLocation

            } catch (IllegalStateException e) {

                // Обрабатываем ошибку
                showError("Невозможно импортировать данные");

            }

        }

    }

    // Класс-помощник для работы с перетаскиванием
    class TransferSupport {

        public boolean isDrop() {
            // проверка поддержки перетаскивания компонентом
            return true;
        }

        public DropLocation getDropLocation() {
            // код получения информации о местоположении drop
            return new DropLocation(0, 0);
        }

    }

    // Класс хранения информации о местоположении
    class DropLocation {

        int x;
        int y;

        public DropLocation(int x, int y) {
            this.x = x;
            this.y = y;
        }

    }
    private static final int PERMISSION_REQUEST_CODE = 1;
    private void checkPermissions() throws Exception {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.READ_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED ||
                    ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
                ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.READ_EXTERNAL_STORAGE, Manifest.permission.WRITE_EXTERNAL_STORAGE}, PERMISSION_REQUEST_CODE);
            }else {
                clip();
                spinner = findViewById(R.id.spinner8);
                spinner2 = findViewById(R.id.spinner2);
                clip3();

                spinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                    @Override
                    public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                        spinner2.setEnabled(position != 0);
                    }

                    @Override
                    public void onNothingSelected(AdapterView<?> parent) {
                    }
                });

                clip4();
                clip5();
                clip6();
                clip7();
                clip8();
            }

        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, String[] permissions, int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (requestCode == PERMISSION_REQUEST_CODE) {
            if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED && grantResults[1] == PackageManager.PERMISSION_GRANTED) {
                // Разрешения получены, выполните необходимые операции
                try { clip();
                spinner = findViewById(R.id.spinner8);
                spinner2 = findViewById(R.id.spinner2);
                clip3();

                spinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                    @Override
                    public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                        spinner2.setEnabled(position != 0);
                    }

                    @Override
                    public void onNothingSelected(AdapterView<?> parent) {
                    }
                });

                clip4();
                clip5();
                clip6();
                clip7();

                    clip8();
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            } else {
                // Разрешения не получены, обработайте это соответствующим образом
            }
        }
    }


// Устанавливаем слушатель выбора элементов для первого спиннера

// Создаем ArrayAdapter с помощью массива строк и стандартного внешнего вида


// Проверяем значения всех пяти спинеров при создании активности
            /*     if (spinner1.getSelectedItemPosition() == 0 ||
                    spinner2.getSelectedItemPosition() == 0 ||
                    spinner3.getSelectedItemPosition() == 0 ||
                    spinner4.getSelectedItemPosition() == 0 ||
                    spinner5.getSelectedItemPosition() == 0) {
                // Если хотя бы одно значение является первым значением, делаем кнопку невидимой
                button.setVisibility(View.INVISIBLE);
            } else {
                // Если все значения выбраны, делаем кнопку видимой
                button.setVisibility(View.VISIBLE);
            }

// Устанавливаем слушатель клика на кнопку
            button.setOnClickListener(new View.OnClickListener() {
                @Override
                public void onClick(View v) {
                    // Проверяем, что хотя бы у одного из спинеров выбрано первое значение
                    if (spinner1.getSelectedItemPosition() == 0 ||
                            spinner2.getSelectedItemPosition() == 0 ||
                            spinner3.getSelectedItemPosition() == 0 ||
                            spinner4.getSelectedItemPosition() == 0 ||
                            spinner5.getSelectedItemPosition() == 0) {
                        // Если да, выводим сообщение о незаполненных полях
                        Toast.makeText(Activity_type_1_2.this, "Не все поля заполнены", Toast.LENGTH_SHORT).show();
                    } else {
                        // Если нет, осуществляем переход на другую активность
                        Intent intent = new Intent(Activity_type_1_2.this, Activity_type_2.class);
                        startActivity(intent);
                    }
                }
            });






*/






        /*    Spinner spinner33 = findViewById(R.id.spinner8);
              spinner33.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                  @Override
                  public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                      if (position == 0) {
                          spinner2.setEnabled(false);
                      } else {
                          spinner2.setEnabled(true);
                          try {
                              clip4(); // вызываем метод, который может выбросить исключение
                          } catch (Exception e) {
                              e.printStackTrace(); // обрабатываем исключение
                          }
                      }
                  }

                @Override
                public void onNothingSelected(AdapterView<?> parent) {
                    // не используется
                }
            });if (spinner.getSelectedItemPosition() == 0) {
                spinner2.setEnabled(false);
            } else {
                spinner2.setEnabled(true);
                try {
                    clip4(); // вызываем метод, который может выбросить исключение
                } catch (Exception e) {
                    e.printStackTrace(); // обрабатываем исключение
                }
            }

            /*spinner2.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                @Override
                public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                    String selectedItem = (String) parent.getSelectedItem();
                    if (selectedItem.trim().equals(" ")) {
                        spinner2.setEnabled(false);
                    } else {
                        spinner2.setEnabled(true);
                    }
                }

                @Override
                public void onNothingSelected(AdapterView<?> parent) {
                    // не используется
                }
            });*/



        //AutoCompleteTextView autoCompleteTextView = findViewById(R.id.autoCompleteTextView);

        /*try {
            clip2();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }  */





    private void clip3() throws Exception {


    autoCompleteTextView = findViewById(R.id.autoCompleteTextView);
    spinner = findViewById(R.id.spinner8);

    // Загружаем данные из Excel файла
    List<List<String>> data = null;
        try {
        data = loadDataFromExcel();
    } catch (Exception e) {
        throw new RuntimeException(e);
    }

    // Добавляем обработчик выбора элемента для autoCompleteTextView
    List<List<String>> finalData = data;



        autoCompleteTextView.setOnItemClickListener(new AdapterView.OnItemClickListener() {
        @Override
        public void onItemClick(AdapterView<?> parent, View view, int position, long id) {
            // Получаем выбранный текст
            String selectedText = autoCompleteTextView.getText().toString();
            List<String> filtered = new ArrayList<>();
            List<String> filtered1= new ArrayList<>();//filtered1.add("");
            AutoCompleteTextView autoCompleteTextView1 = findViewById(R.id.autoCompleteTextView);
            ST =autoCompleteTextView1.getText().toString();
            // Фильтруем данные из Excel
             filtered1 = filterByText(finalData, selectedText);
             filtered = filtered1.stream().distinct().collect(Collectors.toList());


            // Устанавливаем отфильтрованные данные в spinner
            setDataToSpinner(filtered);
        }

    });
        spinner2.setEnabled(spinner.getSelectedItemPosition() != 0);



    }






    private void clip4() throws Exception {

        autoCompleteTextView = findViewById(R.id.autoCompleteTextView);

        spinner2 = findViewById(R.id.spinner2);
        spinner = findViewById(R.id.spinner8);
        // Загружаем данные из Excel файла
        List<List<String>> data1 = null;
        try {
            data1 = loadDataFromExcel2();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        // Добавляем обработчик выбора элемента для autoCompleteTextView
        List<List<String>> finalData1 = data1;














        spinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener(){
        @Override
            public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                // Получаем выбранный текст




                String selectedText = spinner.getSelectedItem().toString();




            AutoCompleteTextView autoCompleteTextView1 = findViewById(R.id.autoCompleteTextView);
            ST =autoCompleteTextView1.getText().toString();
                System.out.println(selectedText);
                // Фильтруем данные из Excel
                List<String> filtered1 = filterByText2(finalData1, selectedText, ST);
                List<String> filtered = filtered1.stream().distinct().collect(Collectors.toList());



                // Устанавливаем отфильтрованные данные в spinner
                setDataToSpinnerSecond(filtered);
            }
            @Override
            public void onNothingSelected(AdapterView<?> parent) {
                // Обработчик, вызываемый если ни один элемент не выбран
            }
        });}


    private void clip5() throws Exception {

        autoCompleteTextView = findViewById(R.id.autoCompleteTextView);

        spinner3 = findViewById(R.id.spinner3);
        spinner2 = findViewById(R.id.spinner2);
        spinner = findViewById(R.id.spinner8);
        // Загружаем данные из Excel файла
        List<List<String>> data1 = null;
        try {
            data1 = loadDataFromExcel3();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        // Добавляем обработчик выбора элемента для autoCompleteTextView
        List<List<String>> finalData1 = data1;













        spinner2.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener(){
            @Override
            public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                // Получаем выбранный текст


                String selectedText1 = spinner.getSelectedItem().toString();
                String selectedText2 = spinner2.getSelectedItem().toString();




                // Фильтруем данные из Excel
               // List<String> filtered12 = filterByText3(finalData1, selectedText1);


                String selectedText = spinner2.getSelectedItem().toString();



                System.out.println(selectedText);
                // Фильтруем данные из Excel
                List<String> filtered1 = filterByText3(finalData1, selectedText,selectedText1);
                List<String> filtered = filtered1.stream().distinct().collect(Collectors.toList());



                // Устанавливаем отфильтрованные данные в spinner
                setDataToSpinner3(filtered);
            }
            @Override
            public void onNothingSelected(AdapterView<?> parent) {
                // Обработчик, вызываемый если ни один элемент не выбран
            }
        });}

    private void clip6() throws Exception {

        autoCompleteTextView = findViewById(R.id.autoCompleteTextView);
         spinner4 = findViewById(R.id.spinner4);
        spinner3 = findViewById(R.id.spinner3);
        spinner2 = findViewById(R.id.spinner2);
        // Загружаем данные из Excel файла
        List<List<String>> data1 = null;
        try {
            data1 = loadDataFromExcel3();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        // Добавляем обработчик выбора элемента для autoCompleteTextView
        List<List<String>> finalData1 = data1;














        spinner3.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener(){
            @Override
            public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                // Получаем выбранный текст




                String selectedText = spinner3.getSelectedItem().toString();
                String selectedText1 = spinner2.getSelectedItem().toString();



                System.out.println(selectedText);
                // Фильтруем данные из Excel
                List<String> filtered1 = filterByText4(finalData1, selectedText,selectedText1 );
                List<String> filtered = filtered1.stream().distinct().collect(Collectors.toList());



                // Устанавливаем отфильтрованные данные в spinner
                setDataToSpinner4(filtered);
            }
            @Override
            public void onNothingSelected(AdapterView<?> parent) {
                // Обработчик, вызываемый если ни один элемент не выбран
            }
        });}


    private void clip7() throws Exception {

        autoCompleteTextView = findViewById(R.id.autoCompleteTextView);
        spinner5= findViewById(R.id.spinner5);
        spinner4 = findViewById(R.id.spinner4);
        spinner3 = findViewById(R.id.spinner3);
        spinner2 = findViewById(R.id.spinner2);
        // Загружаем данные из Excel файла
        List<List<String>> data1 = null;
        try {
            data1 = loadDataFromExcel3();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        // Добавляем обработчик выбора элемента для autoCompleteTextView
        List<List<String>> finalData1 = data1;














        spinner4.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener(){
            @Override
            public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                // Получаем выбранный текст



                String selectedText1 = spinner3.getSelectedItem().toString();
                String selectedText = spinner4.getSelectedItem().toString();



                System.out.println(selectedText);
                // Фильтруем данные из Excel
                List<String> filtered1 = filterByText5(finalData1, selectedText,selectedText1);
                List<String> filtered = filtered1.stream().distinct().collect(Collectors.toList());



                // Устанавливаем отфильтрованные данные в spinner
                setDataToSpinner5(filtered);
            }
            @Override
            public void onNothingSelected(AdapterView<?> parent) {
                // Обработчик, вызываемый если ни один элемент не выбран
            }
        });}






    private void clip8() throws Exception {

        autoCompleteTextView = findViewById(R.id.autoCompleteTextView);
        spinner5= findViewById(R.id.spinner5);
        spinner6= findViewById(R.id.spinner6);
        spinner4 = findViewById(R.id.spinner4);
        spinner3 = findViewById(R.id.spinner3);
        spinner2 = findViewById(R.id.spinner2);
        // Загружаем данные из Excel файла
        List<List<String>> data1 = null;
        try {
            data1 = loadDataFromExcel3();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        // Добавляем обработчик выбора элемента для autoCompleteTextView
        List<List<String>> finalData1 = data1;














        spinner5.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener(){
            @Override
            public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                // Получаем выбранный текст



                String selectedText1 = spinner4.getSelectedItem().toString();
                String selectedText = spinner5.getSelectedItem().toString();




                // Фильтруем данные из Excel
                List<String> filtered1 = filterByText6(finalData1, selectedText,selectedText1);
                List<String> filtered = filtered1.stream().distinct().collect(Collectors.toList());



                // Устанавливаем отфильтрованные данные в spinner
                setDataToSpinner6(filtered);
            }
            @Override
            public void onNothingSelected(AdapterView<?> parent) {
                // Обработчик, вызываемый если ни один элемент не выбран
            }
        });}




    private List<List<String>> loadDataFromExcel() throws Exception {
        List<List<String>> data = null;


            // Если разрешение уже есть, загрузить данные
            try {
                // Загрузка данных из Excel с помощью Aspose.Cells

                Workbook workbook = new Workbook("/storage/emulated/0/Download/Книга1.xlsx");

                // Получаем первый лист
                Worksheet worksheet = workbook.getWorksheets().get(0);

                // Создадим список для данных
                data = new ArrayList<>();

                // Проходим по всем строкам
                for (int rowIndex = 0; rowIndex < worksheet.getCells().getMaxDataRow(); rowIndex++) {
                    // Данные из одной строки
                    List<String> rowData = new ArrayList<>();

                    // Проходим по всем ячейкам в строке
                    for (int colIndex = 0; colIndex < worksheet.getCells().getMaxDataColumn(); colIndex++) {
                        Cell cell = worksheet.getCells().get(rowIndex, colIndex);

                        // Добавляем данные ячейки в список строки
                        rowData.add(cell.getStringValue());
                    }

                    // Добавляем список данных строки в общий список
                    data.add(rowData);
                }


            } catch (FileNotFoundException e) {

                // Обработка ошибки отсутствия файла
                Toast.makeText(this, "Файл не найден", Toast.LENGTH_SHORT).show();

            }

        return data;
    }
    private List<List<String>> loadDataFromExcel2() throws Exception {
        // Загрузка данных из Excel с помощью Aspose.Cells

        Workbook workbook = new Workbook("/sdcard/Download/Книга1.xlsx");

        // Получаем первый лист
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Создадим список для данных
        List<List<String>> data = new ArrayList<>();

        // Проходим по всем строкам
        for (int rowIndex = 0; rowIndex < worksheet.getCells().getMaxDataRow(); rowIndex++) {
            // Данные из одной строки
            List<String> rowData = new ArrayList<>();

            // Проходим по всем ячейкам в строке
            for (int colIndex = 0; colIndex < worksheet.getCells().getMaxDataColumn(); colIndex++) {
                Cell cell = worksheet.getCells().get(rowIndex, colIndex);

                // Добавляем данные ячейки в список строки
                rowData.add(cell.getStringValue());
            }

            // Добавляем список данных строки в общий список
            data.add(rowData);
        }

        return data;
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
        }

        return data;
    }

    private List<String> filterByText(List<List<String>> data, String text) {
        List<String> filtered1 = new ArrayList<>();
        //filtered1.add(" ");


        for(List<String> row : data) {
            if(row.get(0).contains(text)) {
                filtered1.add(row.get(1));
            }
        }
        return filtered1;
    }

    private List<String> filterByText2(List<List<String>> data, String text, String text1) {
        List<String> filtered1 = new ArrayList<>();
       // filtered1.add(" ");



        for(List<String> row : data) {
            if(row.get(1).contains(text)&&row.get(0).contains(text1)) {
                filtered1.add(row.get(6));
            }

        }
        return filtered1;
    }

    private List<String> filterByText3(List<List<String>> data, String text,String text1) {
        List<String> filtered1 = new ArrayList<>();
        // filtered1.add(" ");



        for(List<String> row : data) {
            if(row.get(6).contains(text) && row.get(1).contains(text1) ) {
                filtered1.add(row.get(8));
            }

        }
        return filtered1;
    }
    private List<String> filterByText4(List<List<String>> data, String text,String text1) {
        List<String> filtered1 = new ArrayList<>();
        // filtered1.add(" ");



        for(List<String> row : data) {
            if(row.get(8).contains(text) && row.get(6).contains(text1) ) {
                filtered1.add(row.get(9));
            }

        }
        return filtered1;
    }


    private List<String> filterByText5(List<List<String>> data, String text,String text1) {
        List<String> filtered1 = new ArrayList<>();
        // filtered1.add(" ");



        for(List<String> row : data) {
            if(row.get(9).contains(text) && row.get(8).contains(text1)) {
                filtered1.add(row.get(10));
            }

        }
        return filtered1;
    }

    private List<String> filterByText6(List<List<String>> data, String text,String text1) {
        List<String> filtered1 = new ArrayList<>();
        // filtered1.add(" ");



        for(List<String> row : data) {
            if(row.get(10).contains(text) && row.get(9).contains(text1)) {
                filtered1.add(row.get(13));
            }

        }
        return filtered1;
    }


    private void setDataToSpinner(List<String> data) {
        data.add(0,"Сделайте выбор:");
        ArrayAdapter<String> adapter = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, data);

        spinner.setAdapter(adapter);
    }
     private void setDataToSpinnerSecond(List<String> data) {
         data.add(0,"Сделайте выбор:");
        ArrayAdapter<String> adapter2 = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, data);
        spinner2.setAdapter(adapter2);
    }
    private void setDataToSpinner3(List<String> data) {
        data.add(0,"Сделайте выбор:");
        ArrayAdapter<String> adapter3 = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, data);

        spinner3.setAdapter(adapter3);
    }
    private void setDataToSpinner4(List<String> data) {
        data.add(0,"Сделайте выбор:");
        ArrayAdapter<String> adapter4 = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, data);

        spinner4.setAdapter(adapter4);
    }

    private void setDataToSpinner5(List<String> data) {
        data.add(0,"Сделайте выбор:");
        ArrayAdapter<String> adapter4 = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, data);

        spinner5.setAdapter(adapter4);
    }

    private void setDataToSpinner6(List<String> data) {
        data.add(0,"Сделайте выбор:");
        ArrayAdapter<String> adapter5 = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, data);

        spinner6.setAdapter(adapter5);






        button2 = findViewById(R.id.button2);
// Устанавливаем обработчик события выбора элемента в Spinner
       spinner6.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
            @Override
            public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {

// Если выбран не первый элемент, то разблокируем кнопку
                if (position != 0) {
                    button2.setEnabled(true);
                } else {
                    button2.setEnabled(false);
                }
            }

            @Override
            public void onNothingSelected(AdapterView<?> parent) {
// Ничего не делаем
            }
        });

// Устанавливаем обработчик нажатия на кнопку
     button2.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
// Получаем выбранный элемент в Spinner
                save1();
                //saveW();

                 spinner33 = findViewById(R.id.spinner3);
                String selectedText = spinner33.getSelectedItem().toString();
                if (selectedText.equals("Обследование сосудов и аппаратов")) {
                    Intent intent = new Intent(Activity_type_1_2.this, Activity_type_SHB1.class);
                   // intent.setFlags(Intent.FLAG_ACTIVITY_SINGLE_TOP);
                    startActivity(intent);
                } else if (selectedText.equals("Обследование ТПА")) {
                   Intent intent = new Intent(Activity_type_1_2.this, MainActivity.class);
                   startActivity(intent);
                }
            }
        });
    }








  /*  private void spin() throws Exception {
        // Загрузка сводной таблицы из файла Excel
        Workbook workbook = new Workbook("/sdcard/Download/Книга1.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);
        PivotTable pivotTable = worksheet.getPivotTables().get(0);

// Получение значений для первого списка из первого столбца сводной таблицы
        ArrayList<String> firstListValues = new ArrayList<String>();
        PivotField firstField = pivotTable.getRowFields().get(0);
        for(int i = 0; i < firstField.getPivotItems().getCount(); i++) {
            firstListValues.add(firstField.getPivotItems().get(i).getName());
        }

// Создание ArrayAdapter для первого списка
        ArrayAdapter<String> firstListAdapter = new ArrayAdapter<String>(this, android.R.layout.simple_list_item_1, firstListValues);

// Добавление слушателя событий на первый список
        Spinner firstListSpinner = (Spinner) findViewById(R.id.first_list_spinner);
        firstListSpinner.setAdapter(firstListAdapter);
        /*firstListSpinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
             @Override
                       public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                // Получение выбранного значения из первого списка
                String selectedValue = parent.getItemAtPosition(position).toString();

                // Загрузка данных для второго списка из Excel

                Worksheet worksheet = workbook.getWorksheets().get(0);
                Cells cells = worksheet.getCells();

                // Фильтрация данных в соответствии с выбранным значением из первого списка
                ArrayList<String> secondListValues = new ArrayList<String>();
                for(int i = 0; i < cells.getMaxDataRow(); i++) {
                    if(cells.get(i, 0).getValue().toString().equals(selectedValue)) {
                        secondListValues.add(cells.get(i, 1).getValue().toString());
                    }
                }

                // Создание ArrayAdapter для второго списка
                ArrayAdapter<String> secondListAdapter = new ArrayAdapter<String>(Activity_type_1_2.this, android.R.layout.simple_list_item_1, secondListValues);

                // Отображение второго списка на экране
                Spinner secondListSpinner = (Spinner) findViewById(R.id.second_list_spinner);
                secondListSpinner.setAdapter(secondListAdapter);
            }

            @Override
            public void onNothingSelected(AdapterView<?> parent) {
                // Do nothing
            }
        }}*/

    private void clip() throws Exception {

        // Получить путь к файлу Excel
        String filePath = "/sdcard/Download/Книга1.xlsx";

// Создать объект Workbook из файла Excel
        Workbook workbook = new Workbook(filePath);

// Получить лист с данными
        Worksheet worksheet = workbook.getWorksheets().get(0);

// Создать список строк для автодополнения
        List<String> autocompleteList1 = new ArrayList<>();
        List<String> autocompleteList= new ArrayList<>();
        autocompleteList.add("");





// Пройти по всем строкам и добавить название фильма в список
        for (int row = 1; row <= worksheet.getCells().getMaxDataRow(); row++) {
            Cell cell = worksheet.getCells().get(row, 0);
            String filmName = cell.getStringValue();
            autocompleteList1.add(filmName);
        }
        autocompleteList = autocompleteList1.stream().distinct().collect(Collectors.toList());

        // Создать адаптер для AutoCompleteTextView
        ArrayAdapter<String> adapter =
                new ArrayAdapter<>(this, android.R.layout.simple_dropdown_item_1line, autocompleteList);

// Установить адаптер для AutoCompleteTextView
        AutoCompleteTextView autoCompleteTextView1 = findViewById(R.id.autoCompleteTextView);
        ST =autoCompleteTextView1.getText().toString();
        autoCompleteTextView1.setAdapter(adapter);


//clip3();
    }
     private void clip2() throws Exception {List<String> secondListValues = new ArrayList<>();

// Загрузка данных из Excel-файла
        Workbook workbook = new Workbook("/sdcard/Download/Книга1.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();

// Получаем список элементов AutoCompleteTextView
        AutoCompleteTextView autoCompleteTextView = findViewById(R.id.autoCompleteTextView);
        String[] firstListValues = autoCompleteTextView.getText().toString().split(",");
         //EditText editText4 = findViewById(R.id.editText_111);
         //String[] firstListValues = editText4.getText().toString().split(",");


// Перебираем ячейки первого столбца Excel
        for (int i = 0; i < cells.getMaxDataRow() + 1; i++) {
            Cell cell = cells.get(i, 0);
            String value = cell.getStringValue();

            // Проверяем, содержит ли значение ячейки первого столбца текст из AutoCompleteTextView
            for (String firstListValue : firstListValues) {
                if (value.contains(firstListValue.trim())) {
                    // Если да, то добавляем соответствующее значение из второго столбца в список secondListValues
                    Cell secondColumnCell = cells.get(i, 1);
                    String secondColumnValue = secondColumnCell.getStringValue();
                    secondListValues.add(secondColumnValue);
                    break;
                }
            }
        }

// Создаем ArrayAdapter для заполнения второго выпадающего списка
        ArrayAdapter<String> secondListAdapter = new ArrayAdapter<>(this, android.R.layout.simple_list_item_1, secondListValues);

// Присваиваем ArrayAdapter второму выпадающему списку
        Spinner secondListSpinner = findViewById(R.id.spinner8);
        secondListSpinner.setAdapter(secondListAdapter);

    }







    /*private void clip2() throws Exception {

// Загрузка данных из Excel-файла
        Workbook workbook = new Workbook("/sdcard/Download/Книга1.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();
        EditText editText = findViewById(R.id.editText_111);
        String searchText = editText.getText().toString();


        List<String> list = new ArrayList<String>();
        for (int i = 0; i < worksheet.getCells().getMaxDataRow() + 1; i++) {
            String cellValueA = worksheet.getCells().get("A" + (i + 1)).getStringValue();
            String cellValueB = worksheet.getCells().get("B" + (i + 1)).getStringValue();
            if (cellValueA.contains(searchText)) {
                list.add(cellValueB);
            }
        }
        ArrayAdapter<String> adapter = new ArrayAdapter<String>(Activity_type_1_2.this, android.R.layout.simple_spinner_item, list);
        Spinner secondListSpinner = findViewById(R.id.spinner8);
        secondListSpinner.setAdapter(adapter);
    }*/

    String fName;


    private void save1(){
        String s8 = spinner6.getSelectedItem().toString();

        String s00 = s8.replace("/", "");
        ExcelDataSaverAct111 dataSaver1 = new ExcelDataSaverAct111(fileName, folderName,s8);
        try {



             autoCompleteTextView = findViewById(R.id.autoCompleteTextView);
            String s1 = autoCompleteTextView.getText().toString();

             spinner = findViewById(R.id.spinner8);
            String s2 = spinner.getSelectedItem().toString();

             spinner2 = findViewById(R.id.spinner2);
            String s3 = spinner2.getSelectedItem().toString();

             spinner3 = findViewById(R.id.spinner3);
            String s4 = spinner3.getSelectedItem().toString();

             spinner4 = findViewById(R.id.spinner4);
            String s5 = spinner4.getSelectedItem().toString();

             spinner5 = findViewById(R.id.spinner5);
            String s6 = spinner5.getSelectedItem().toString();

             spinner6 = findViewById(R.id.spinner6);
            String s7 = spinner6.getSelectedItem().toString();
            fName = s7;




            // Вставить значение в первую ячейку











            dataSaver1.saveData1(s1, s2, s3,s4, s5, s6, s7);
            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }

    }


    private void saveW(){
        String s8 = spinner6.getSelectedItem().toString();

        String s00 = s8.replace("/", "");

        try {



            autoCompleteTextView = findViewById(R.id.autoCompleteTextView);
            String s1 = autoCompleteTextView.getText().toString();

            spinner = findViewById(R.id.spinner8);
            String s2 = spinner.getSelectedItem().toString();

            spinner2 = findViewById(R.id.spinner2);
            String s3 = spinner2.getSelectedItem().toString();

            spinner3 = findViewById(R.id.spinner3);
            String s4 = spinner3.getSelectedItem().toString();

            spinner4 = findViewById(R.id.spinner4);
            String s5 = spinner4.getSelectedItem().toString();

            spinner5 = findViewById(R.id.spinner5);
            String s6 = spinner5.getSelectedItem().toString();

            spinner6 = findViewById(R.id.spinner6);
            String s7 = spinner6.getSelectedItem().toString();
            fName = s7;




            // Вставить значение в первую ячейку












            Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
        }

    }




    String folderName =" " ;
    String fileName ="2";
    public String getCurrentDate() {
        SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy", Locale.getDefault());
        return sdf.format(new Date());
    }

    public String getCurrentDateTime() {
        spinner6 = findViewById(R.id.spinner6);
        String s7 = spinner6.getSelectedItem().toString();
        return (s7);
    }



    public class PermissionUtils {
        public static boolean hasPermissions(Context context) {
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
                return Environment.isExternalStorageManager();
            } else if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
                return ContextCompat.checkSelfPermission(context, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                        == PackageManager.PERMISSION_GRANTED;
            } else {
                return true;
            }
        }

        public static void requestPermissions(Activity activity, int requestCode) {
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
                try {
                    Intent intent = new Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION);
                    intent.addCategory("android.intent.category.DEFAULT");
                    intent.setData(Uri.parse(String.format("package:%s", activity.getPackageName())));
                    activity.startActivityForResult(intent, requestCode);
                } catch (Exception e) {
                    Intent intent = new Intent();
                    intent.setAction(Settings.ACTION_MANAGE_ALL_FILES_ACCESS_PERMISSION);
                    activity.startActivityForResult(intent, requestCode);
                }
            } else {
                ActivityCompat.requestPermissions(activity,
                        new String[] { Manifest.permission.WRITE_EXTERNAL_STORAGE },
                        requestCode);
            }
        }
    }

}

