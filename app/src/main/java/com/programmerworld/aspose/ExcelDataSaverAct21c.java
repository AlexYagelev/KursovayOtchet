
package com.programmerworld.aspose;

        import com.aspose.cells.Cell;
        import com.aspose.cells.SaveFormat;
        import com.aspose.cells.Workbook;
        import com.aspose.cells.Worksheet;

 public class ExcelDataSaverAct21c {

    private String fileName;



     public ExcelDataSaverAct21c(String fileName) {
        this.fileName = fileName;
    }

    public void saveData(String i1, String i11, String i111,String a, String aa, String aaa, String b, String bb, String bbb, String c, String cc,String ccc, String d, String dd,String ddd, String e, String ee,String eee, String f, String ff,String fff, String g, String gg,String ggg, String h, String hh,String hhh, String i, String ii, String iii,String j, String jj,String jjj, String k, String kk,String kkk, String l, String ll,String lll) throws Exception {
        int tpm1 = 20;
// Создаем новый файл Excel с помощью Aspose.Cells
        Workbook workbook = new Workbook("/sdcard/Download"+"/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString+".xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);

        for(int i_1 = 0; i_1<12;i_1++){
            Cell cell = worksheet.getCells().get(0,tpm1);
            cell.setValue(i1);
            ++tpm1;

            cell = worksheet.getCells().get(0,tpm1);
            cell.setValue(i11);
            ++tpm1;

            cell = worksheet.getCells().get(0,tpm1);
            cell.setValue(i111);
            ++tpm1;
        }







// Заполняем таблицу данными


        int tpm = 20;

        Cell cell = worksheet.getCells().get(1,tpm);
        cell.setValue(a);
         ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(aa);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(aaa);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(b);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(bb);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(bbb);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(c);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(cc);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(ccc);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(d);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(dd);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(ddd);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(e);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(ee);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(eee);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(f);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(ff);
        ++tpm;
        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(fff);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(g);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(gg);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(ggg);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(h);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(hh);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(hhh);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(i);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(ii);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(iii);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(j);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(jj);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(jjj);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(k);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(kk);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(kkk);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(l);
        ++tpm;

        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(ll);
        ++tpm;
        cell = worksheet.getCells().get(1,tpm);
        cell.setValue(lll);
        ++tpm;




// Сохраняем файл Exc
        String outputDirectory = "/sdcard/Download"; // Место для сохранения файла
        String outputFile = outputDirectory +"/" + ExcelDataSaverAct111.outputString + "/" + ExcelDataSaverAct111.outputString+".xlsx";
        workbook.save(outputFile, SaveFormat.XLSX);
    }
}

