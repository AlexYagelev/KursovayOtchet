package com.programmerworld.aspose;

import android.os.Bundle;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.AutoCompleteTextView;
import android.widget.Spinner;

import androidx.appcompat.app.AppCompatActivity;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class Activity_type_1_2 extends AppCompatActivity {private Spinner eoSpinner;
    private Spinner objectNameSpinner;
    private Spinner objectSpinner;
    private Spinner workSpinner;
    private Spinner dataSpinner;


    private List<String> eoValues = new ArrayList<>();
    private List<String> objectNameValues = new ArrayList<>();
    private List<String> objectValues = new ArrayList<>();
    private List<String> workValues = new ArrayList<>();
    private List<String> dataValues = new ArrayList<>();
    private AutoCompleteTextView autoCompleteTextView;
    private Spinner spinner;
    private Spinner spinner2;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_type12);
        /*try { spin();} catch (Exception e) {
            throw new RuntimeException(e);
        }
      */  try {
            clip();
            clip3();
             clip4();

        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        //AutoCompleteTextView autoCompleteTextView = findViewById(R.id.autoCompleteTextView);

        /*try {
            clip2();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }  */

    }
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

            // Фильтруем данные из Excel
            List<String> filtered1 = filterByText(finalData, selectedText);
            List<String> filtered = filtered1.stream().distinct().collect(Collectors.toList());


            // Устанавливаем отфильтрованные данные в spinner
            setDataToSpinner(filtered);
        }
    });}


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
                System.out.println(selectedText);
                // Фильтруем данные из Excel
                List<String> filtered1 = filterByText2(finalData1, selectedText);
                List<String> filtered = filtered1.stream().distinct().collect(Collectors.toList());


                // Устанавливаем отфильтрованные данные в spinner
                setDataToSpinnerSecond(filtered);
            }
            @Override
            public void onNothingSelected(AdapterView<?> parent) {
                // Обработчик, вызываемый если ни один элемент не выбран
            }
        });}





    private List<List<String>> loadDataFromExcel() throws Exception {
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

    private List<String> filterByText(List<List<String>> data, String text) {
        List<String> filtered1 = new ArrayList<>();
        for(List<String> row : data) {
            if(row.get(0).contains(text)) {
                filtered1.add(row.get(1));
            }
        }
        return filtered1;
    }

    private List<String> filterByText2(List<List<String>> data, String text) {
        List<String> filtered1 = new ArrayList<>();
        for(List<String> row : data) {
            if(row.get(1).contains(text)) {
                filtered1.add(row.get(6));
            }
        }
        return filtered1;
    }

    private void setDataToSpinner(List<String> data) {
        ArrayAdapter<String> adapter = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, data);
        spinner.setAdapter(adapter);
    }
     private void setDataToSpinnerSecond(List<String> data) {
        ArrayAdapter<String> adapter2 = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, data);
        spinner2.setAdapter(adapter2);
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


// Пройти по всем строкам и добавить название фильма в список
        for (int row = 1; row <= worksheet.getCells().getMaxDataRow(); row++) {
            Cell cell = worksheet.getCells().get(row, 0);
            String filmName = cell.getStringValue();
            autocompleteList1.add(filmName);
        }
        List<String> autocompleteList = autocompleteList1.stream().distinct().collect(Collectors.toList());

        // Создать адаптер для AutoCompleteTextView
        ArrayAdapter<String> adapter =
                new ArrayAdapter<>(this, android.R.layout.simple_dropdown_item_1line, autocompleteList);

// Установить адаптер для AutoCompleteTextView
        AutoCompleteTextView autoCompleteTextView = findViewById(R.id.autoCompleteTextView);
        autoCompleteTextView.setAdapter(adapter);



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
}

