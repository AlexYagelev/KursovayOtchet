package com.programmerworld.aspose;

import androidx.lifecycle.MutableLiveData;
import androidx.lifecycle.ViewModel;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelDataSaverAct21d extends ViewModel {


    private MutableLiveData<User> user = new MutableLiveData<>();
    private final String fileName;
    private int cellsCounter;
    public ExcelDataSaverAct21d(String fileName) {
        this.fileName = fileName;
    }
    public void saveData1(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU4());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU4().toString());
            ++count;
        }


        if(!user.getUU5().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU5());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU5().toString());
            ++count;
        }


        if(!user.getUU6().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU6());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU6().toString());
            ++count;
        }


        if(!user.getUU7().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU7());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU7().toString());
            ++count;
        }



        if(!user.getUU8().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU8());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU8().toString());
            ++count;
        }


        if(!user.getUU9().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU9());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU9().toString());
            ++count;
        }


        if(!user.getUU10().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU10());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU10().toString());
            ++count;
        }





// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }public void saveData4(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU4());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU4().toString());
            ++count;
        }


        if(!user.getUU5().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU5());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU5().toString());
            ++count;
        }


        if(!user.getUU6().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU6());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU6().toString());
            ++count;
        }






// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
    public void saveData1_6pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(15,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(16, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU4());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU4().toString());
            ++count;
        }


        if(!user.getUU5().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU5());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU5().toString());
            ++count;
        }


        if(!user.getUU6().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU6());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU6().toString());
            ++count;
        }


        if(!user.getUU7().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU7());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU7().toString());
            ++count;
        }



        if(!user.getUU8().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU8());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU8().toString());
            ++count;
        }






// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
    public void saveData2_3(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }





// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
    public void saveData4_6pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(15,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(16, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU4());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU4().toString());
            ++count;
        }


        if(!user.getUU5().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU5());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU5().toString());
            ++count;
        }


        if(!user.getUU6().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU6());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU6().toString());
            ++count;
        }






// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }

    public void saveData2_3_2pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(3, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(4, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(3, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(4, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }





// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }



    public void saveData2_2_3pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(6,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(7, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(6, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(7, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }


// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }


    public void saveData2_2(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }


// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }

    public void saveData2_3_1(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }











// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
    public void saveData2_3_4pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(9,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(10, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(9, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(10, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(9, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(10, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }











// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
    public void saveData2_2_5pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(12,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(13, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(12, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(13, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }


// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
    public void saveData2_2_6pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(15,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(16, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(15, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(16, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }


// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }

    public void saveData2_2_7pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(18,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(19, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(18, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(19, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }


// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }

    public void saveData2_2_1(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }


// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }

    public void saveData2_2_4pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(9,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(10, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(9, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(10, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }


// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }

    public void saveData2_4(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU4());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU4().toString());
            ++count;
        }









// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }

    public void saveData2_4_3pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(6,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(7, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(6, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(7, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(6, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(7, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(6, count);
            cell.setValue(user.getU4());
            Cell cell1 = worksheet.getCells().get(7, count);
            cell1.setValue(user.getUU4().toString());
            ++count;
        }









// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
    public void saveData2(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU4());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU4().toString());
            ++count;
        }


        if(!user.getUU5().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU5());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU5().toString());
            ++count;
        }






// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }

    public void saveData2_7pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(18,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(19, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(18, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(19, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(18, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(19, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(18, count);
            cell.setValue(user.getU4());
            Cell cell1 = worksheet.getCells().get(19, count);
            cell1.setValue(user.getUU4().toString());
            ++count;
        }


        if(!user.getUU5().isEmpty()) {
            Cell cell = worksheet.getCells().get(18, count);
            cell.setValue(user.getU5());
            Cell cell1 = worksheet.getCells().get(19, count);
            cell1.setValue(user.getUU5().toString());
            ++count;
        }






// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }


    public void saveData2_5(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getU2().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU2().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU2().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUUU2().toString());
            ++count;
        }


        if(!user.getU3().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU3().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU3().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU4().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU4().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUUU4().toString());
            ++count;
        }

        if(!user.getU5().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU5().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU5().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUUU5().toString());
            ++count;
        }


        if(!user.getU6().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU6().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU6().toString());
            ++count;
            cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUUU6().toString());
            ++count;
        }





// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }


    public void saveData2_5_2pr(User user) throws Exception { Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getU2().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getU2().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUU2().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUUU2().toString());
            ++count;
        }


        if(!user.getU3().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getU3().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUU3().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getU4().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUU4().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUUU4().toString());
            ++count;
        }

        if(!user.getU5().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getU5().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUU5().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUUU5().toString());
            ++count;
        }


        if(!user.getU6().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getU6().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUU6().toString());
            ++count;
            cell11 = worksheet.getCells().get(3,count);
            cell11.setValue((user.getUUU1()).toString());
            cell12 = worksheet.getCells().get(4, count);
            cell12.setValue(user.getUUU6().toString());
            ++count;
        }





// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
    public void saveData(User user) throws Exception {

// Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU1()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU2());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU2().toString());
            ++count;
        }

        if(!user.getUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU3());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU3().toString());
            ++count;
        }

        if(!user.getUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU4());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU4().toString());
            ++count;
        }


        if(!user.getUU5().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU5());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU5().toString());
            ++count;
        }


        if(!user.getUU6().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU6());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU6().toString());
            ++count;
        }


        if(!user.getUU7().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU7());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU7().toString());
            ++count;
        }



        if(!user.getUU8().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU8());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU8().toString());
            ++count;
        }


        if(!user.getUU9().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU9());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU9().toString());
            ++count;
        }


        if(!user.getUU10().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU10());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU10().toString());
            ++count;
        }

        if(!user.getUU11().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getU11());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUU11().toString());
            ++count;
        }



// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}

