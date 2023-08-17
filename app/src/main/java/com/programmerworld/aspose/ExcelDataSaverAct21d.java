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

