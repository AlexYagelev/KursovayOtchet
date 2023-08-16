/*package com.programmerworld.aspose;

import com.aspose.words.Document;
import com.aspose.words.FindReplaceOptions;

public class W1 {

    public static void main(String[] args) throws Exception {

        // Путь к исходному DOCX файлу
        String inputPath = "/sdcard/Download/101.docx";
        String inputPath1 = "/sdcard/Download/1011.docx";
        

        // Открываем документ
        Document doc = new Document(inputPath);

        // Создаем опции поиска и замены
        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(true);

        // Ищем строку "старая строка" и заменяем на "новая строка"
        doc.getRange().replace("сосуд, работающий под давлением", "новая строка", options);

        // Сохраняем измененный DOCX файл
        doc.save(inputPath1);



    }

}*/