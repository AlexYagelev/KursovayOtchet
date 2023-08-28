package com.programmerworld.aspose.TPA.pr1
import android.annotation.SuppressLint
import android.content.Intent
import android.os.Bundle
import android.widget.EditText
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import com.aspose.cells.Workbook
import com.aspose.cells.Worksheet
import com.aspose.cells.a.b.a.l9m.e
import com.programmerworld.aspose.ExcelDataSaverAct21d
import com.programmerworld.aspose.R
import com.programmerworld.aspose.User
import com.programmerworld.aspose.sosudi.Activity_type_2b2

private fun String.printStackTrace() {
    TODO("Not yet implemented")
}

class path2 : AppCompatActivity() {
    @SuppressLint("MissingInflatedId")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_pr2)



                findViewById<TextView>(R.id.button1).setOnClickListener { v ->
                    val intent = Intent(this, Activity_type_2b2::class.java)
                    startActivity(intent)
                }

                findViewById<TextView>(R.id.button3).setOnClickListener { v ->
                    save()
                    Toast.makeText(this, "Данные сохранены", Toast.LENGTH_SHORT).show()
                    clearFields()
                }

                findViewById<TextView>(R.id.button2).setOnClickListener { v ->
                    finish()
                    supportFragmentManager.popBackStack()
                }
            }

            private fun clearFields() {
                findViewById<EditText>(R.id.editText1).setText("")
                findViewById<EditText>(R.id.editText2).setText("")
                findViewById<EditText>(R.id.editText3).setText("")
                findViewById<EditText>(R.id.editText4).setText("")
                findViewById<EditText>(R.id.editText5).setText("")
                findViewById<EditText>(R.id.editText6).setText("")
                findViewById<EditText>(R.id.editText7).setText("")
                findViewById<EditText>(R.id.editText8).setText("")
                findViewById<EditText>(R.id.editText9).setText("")
                findViewById<EditText>(R.id.editText10).setText("")
                findViewById<EditText>(R.id.editText11).setText("")
            }

            @Throws(Exception::class)
            private fun loadDataFromExcel3(): MutableList<MutableList<String>> {
                // Загрузка данных из Excel с помощью Aspose.Cells
                val workbook = Workbook("/sdcard/Download/Книга1.xlsx")
                // Получаем первый лист
                val worksheet: Worksheet = workbook.worksheets[0]
                // Создадим список для данных
                val data: MutableList<MutableList<String>> = ArrayList()
                // Проходим по всем строкам
                for (rowIndex in 0..worksheet.cells.maxDataRow) {
                    // Данные из одной строки
                    val rowData: MutableList<String> = ArrayList()
                    // Проходим по всем ячейкам в строке
                    for (colIndex in 0..worksheet.cells.maxDataColumn) {
                        val cell = worksheet.cells[rowIndex, colIndex]
                        // Добавляем данные ячейки в список строки
                        rowData.add(cell.stringValue)
                    }
                    // Добавляем список данных строки в общий список
                    data.add(rowData)
                }
                return data
            }

            private fun save() {
                val dataSaver = ExcelDataSaverAct21d("example.xlsx")
                try {
                    val text = findViewById<TextView>(R.id.Text1).text.toString()
                    val text1 = findViewById<EditText>(R.id.editText1).text.toString()
                    val user = User()
                    user.setU1(text)
                    user.setUU1 (text1)


                    user.setU2(findViewById<TextView>(R.id.Text2).text.toString())
                    user.setUU2(findViewById<EditText>(R.id.editText2).text.toString())
                    user.setU3(findViewById<TextView>(R.id.Text3).text.toString())
                    user.setUU3(findViewById<EditText>(R.id.editText3).text.toString())
                    user.setU4(findViewById<TextView>(R.id.Text4).text.toString())
                    user.setUU4(findViewById<EditText>(R.id.editText4).text.toString())
                    user.setU5(findViewById<TextView>(R.id.Text5).text.toString())
                    user.setUU5(findViewById<EditText>(R.id.editText5).text.toString())
                    user.setU6(findViewById<TextView>(R.id.Text6).text.toString())
                    user.setUU6(findViewById<EditText>(R.id.editText6).text.toString())
                    user.setU7(findViewById<TextView>(R.id.Text7).text.toString())
                    user.setUU7(findViewById<EditText>(R.id.editText7).text.toString())
                    dataSaver.saveData(user)
                    // Остальная часть кода сохранения данных в Excel
                    Toast.makeText(this, "Данные сохранены успешно", Toast.LENGTH_SHORT).show()
                } catch (e: Exception) {
                    e.printStackTrace();
                    Toast.makeText(this, "Ошибка при сохранении данных", Toast.LENGTH_SHORT).show();
                }
            }
        }