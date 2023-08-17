package com.programmerworld.aspose;

import androidx.lifecycle.MutableLiveData;
import androidx.lifecycle.ViewModel;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelDataSaverAct21b2 extends ViewModel {


    private MutableLiveData<User> user = new MutableLiveData<>();
    private final String fileName;
    private int cellsCounter;
    public ExcelDataSaverAct21b2(String fileName) {
        this.fileName = fileName;
    }

    public void saveData(User user) throws Exception {

// Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download" + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int  count = worksheet.getCells().getMaxDataColumn()+1;


// Заполняем таблицу данными
        if(!user.getU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU1().toString());
            ++count;}


        if(!user.getUU1().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU1().toString());
            ++count;}


        if(!user.getUUU1().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU1().toString());
            ++count;
        }

        if(!user.getUUUU1().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU1().toString());
            ++count;
        }


        if(!user.getU2().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU2().toString());
            ++count;}


        if(!user.getUU2().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU2().toString());
            ++count;}


        if(!user.getUUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU2().toString());
            ++count;
        }

        if(!user.getUUUU2().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU2().toString());
            ++count;
        }


        if(!user.getU3().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU3().toString());
            ++count;}


        if(!user.getUU3().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU3().toString());
            ++count;}


        if(!user.getUUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU3().toString());
            ++count;
        }

        if(!user.getUUUU3().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU3().toString());
            ++count;
        }


        if(!user.getU4().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU4().toString());
            ++count;}


        if(!user.getUU4().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU4().toString());
            ++count;}


        if(!user.getUUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU4().toString());
            ++count;
        }

        if(!user.getUUUU4().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU4().toString());
            ++count;
        }



        if(!user.getU5().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU5().toString());
            ++count;}


        if(!user.getUU5().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU5().toString());
            ++count;}


        if(!user.getUUU5().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU5().toString());
            ++count;
        }

        if(!user.getUUUU5().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU5().toString());
            ++count;
        }


        if(!user.getU6().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU6().toString());
            ++count;}


        if(!user.getUU6().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU6().toString());
            ++count;}


        if(!user.getUUU6().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU6().toString());
            ++count;
        }

        if(!user.getUUUU6().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU6().toString());
            ++count;
        }

        if(!user.getU7().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU7().toString());
            ++count;}


        if(!user.getUU7().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU7().toString());
            ++count;}


        if(!user.getUUU7().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU7().toString());
            ++count;
        }

        if(!user.getUUUU7().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU7().toString());
            ++count;
        }
        if(!user.getU8().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU8().toString());
            ++count;}


        if(!user.getUU8().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU8().toString());
            ++count;}


        if(!user.getUUU8().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU8().toString());
            ++count;
        }

        if(!user.getUUUU8().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU8().toString());
            ++count;
        }

        if(!user.getU9().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU9().toString());
            ++count;}


        if(!user.getUU9().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU9().toString());
            ++count;}


        if(!user.getUUU9().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU9().toString());
            ++count;
        }

        if(!user.getUUUU9().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU9().toString());
            ++count;
        }

        if(!user.getU10().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU10().toString());
            ++count;}


        if(!user.getUU10().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU10().toString());
            ++count;}


        if(!user.getUUU10().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU10().toString());
            ++count;
        }

        if(!user.getUUUU10().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU10().toString());
            ++count;
        }
        if(!user.getU11().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU11().toString());
            ++count;}


        if(!user.getUU11().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU11().toString());
            ++count;}


        if(!user.getUUU11().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU11().toString());
            ++count;
        }

        if(!user.getUUUU11().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU11().toString());
            ++count;
        }


        if(!user.getU12().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU12().toString());
            ++count;}


        if(!user.getUU12().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU12().toString());
            ++count;}


        if(!user.getUUU12().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU12().toString());
            ++count;
        }

        if(!user.getUUUU12().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU12().toString());
            ++count;
        }


        if(!user.getU13().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU13().toString());
            ++count;}


        if(!user.getUU13().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU13().toString());
            ++count;}


        if(!user.getUUU13().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU13().toString());
            ++count;
        }

        if(!user.getUUUU13().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU13().toString());
            ++count;
        }

        if(!user.getU14().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU14().toString());
            ++count;}


        if(!user.getUU14().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU14().toString());
            ++count;}


        if(!user.getUUU14().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU14().toString());
            ++count;
        }

        if(!user.getUUUU14().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU14().toString());
            ++count;
        }
        if(!user.getU15().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU15().toString());
            ++count;}


        if(!user.getUU15().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU15().toString());
            ++count;}


        if(!user.getUUU15().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU15().toString());
            ++count;
        }

        if(!user.getUUUU15().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU15().toString());
            ++count;
        }
        if(!user.getU16().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getU16().toString());
            ++count;}


        if(!user.getUU16().isEmpty()) {
            Cell cell11 = worksheet.getCells().get(0,count);
            cell11.setValue((user.getUU17()).toString());
            Cell cell12 = worksheet.getCells().get(1, count);
            cell12.setValue(user.getUU16().toString());
            ++count;}


        if(!user.getUUU16().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUU16().toString());
            ++count;
        }

        if(!user.getUUUU16().isEmpty()) {
            Cell cell = worksheet.getCells().get(0, count);
            cell.setValue(user.getUUUU17());
            Cell cell1 = worksheet.getCells().get(1, count);
            cell1.setValue(user.getUUUU16().toString());
            ++count;
        }


// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory + "/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString + ".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}

