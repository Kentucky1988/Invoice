using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.ComponentModel;
using System.Windows;

namespace Накладние
{
    class ViewModel
    {
        string textTextBlock;//текст который передается текстблоку

        public string TextTextBlock
        {
            get { return textTextBlock; }
            set { textTextBlock = value; }
        }

        FileInfo[] files; //масив файлов с расширением .mmo

        public int FileArray()//создание папок
        {
            int z = 1;//номер пункта в текстбоксе

            //string array = null;//текст который передается текстбоксу

            if (!Directory.Exists("Накладные поставщика"))
            {
                TextTextBlock += $"{z++}. Папка \"Накладные поставщика\" отсутствует , она буде создана\n\n";
                Directory.CreateDirectory("Накладные поставщика");//создает папку если она не существует
            }

            if (!Directory.Exists("Накладные"))
            {
                TextTextBlock += $"{z++}. Папка \"Папка \"Накладные\" отсутствует , она буде создана\n\n";
                Directory.CreateDirectory("Накладные");//создает папку если она не существует
            }

            DirectoryInfo directory = new DirectoryInfo("Накладные поставщика");
            files = directory.GetFiles("*.mmo");//масив файлов с расширением .mmo в папке "Накладные поставщика"   

            if (files.Length == 0)//проверка количества файлов .mmo в папке
            {
                TextTextBlock += $"{z++}. В папке \"Накладные поставщика\" не найдено файлов с расширением mmo\n\n";
            }
            else
            {
                TextTextBlock += $"{z++}. В папке \"Накладные поставщика\" найдено {files.Length} файла\n\n";
            }

            return files.Length;//колицество файлов .mmo в папке
        }

        public void Main(BackgroundWorker backgroundWorker)
        {
            Document open = new Document();//клас который создает шапку документа
            Excel.Application ObjExcel = new Excel.Application();//создать приложение Excel
            Excel.Workbook oldWorkBook, newWorkBook;
            Excel.Worksheet oldWorkSheet, newWorkSheet;

            ObjExcel.Visible = false; //скрыть открытие документа

            int position = 0;//количество строк
            double sum = 0;//сумма цен с НДС            

            for (int i = 0; i < files.Length; i++)
            {
                newWorkBook = ObjExcel.Workbooks.Add(1);//количество листов в новом документе          
                newWorkSheet = newWorkBook.Sheets[1];//в каком листе нового документа делать запись
                newWorkSheet.Range[newWorkSheet.Columns[1], newWorkSheet.Columns[11]].ColumnWidth = 8;//ширина стобцов "А:K" = 8  
                newWorkSheet.Columns[10].ColumnWidth = 8.43;
                newWorkSheet.PageSetup.LeftMargin = 35;//отступ слева
                newWorkSheet.PageSetup.RightMargin = 10;//отступ справа            

                position = 0; sum = 0;
                oldWorkBook = ObjExcel.Workbooks.Open(files[i].FullName); //открываем наш файл       
                oldWorkSheet = oldWorkBook.Worksheets[1];//с каким листом работать

                for (int x = 0; oldWorkSheet.Cells[x + 4, 2].Value != null; x++)
                {
                    position = x + 6;//робочая строка
                    newWorkSheet.Cells[position, 1] = oldWorkSheet.Cells[x + 4, 2];//названия 
                    if (oldWorkSheet.Cells[2, 4].Value == "ТОВ Фірма \"ЛАКС\"")
                    {
                        newWorkSheet.Cells[position, 5] = oldWorkSheet.Cells[x + 4, 16].Value / 1000;//количество                  
                        newWorkSheet.Cells[position, 6] = oldWorkSheet.Cells[x + 4, 20].Value / 10000;//цена 1 шт. 
                    }
                    else
                    {
                        newWorkSheet.Cells[position, 5] = oldWorkSheet.Cells[x + 4, 16];//количество                  
                        newWorkSheet.Cells[position, 6] = oldWorkSheet.Cells[x + 4, 20];//цена 1 шт. 
                    }
                    newWorkSheet.Cells[position, 7] = newWorkSheet.Cells[position, 5].Value * newWorkSheet.Cells[position, 6].Value;//цена всего                 
                    newWorkSheet.Cells[position, 8] = oldWorkSheet.Cells[x + 4, 9];//Кофециент наценки НДС     
                    newWorkSheet.Cells[position, 10] = "=MROUND((RC[-4]*RC[-1])+(RC[-4]*RC[-1]*RC[-2]/100),0.05)";//цена с наценкой  
                    newWorkSheet.Cells[position, 11] = "=RC[-1]*RC[-6]";//сумма всего                       

                    newWorkSheet.Range[newWorkSheet.Cells[position, 5], newWorkSheet.Cells[position, 11]].HorizontalAlignment = Excel.Constants.xlCenter;//выравнивание по центру 
                    newWorkSheet.Range[newWorkSheet.Cells[position, 1], newWorkSheet.Cells[position, 11]].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//создание таблицы
                    newWorkSheet.Range[newWorkSheet.Cells[position, 1], newWorkSheet.Cells[position, 4]].Merge();//обэдинение ячеек

                    sum = sum + (newWorkSheet.Cells[position, 7].Value + (newWorkSheet.Cells[position, 7].Value * (newWorkSheet.Cells[position, 8].Value / 100)));//  цена всего с НДС
                }
                newWorkSheet.Cells[position + 1, 2] = "Сумма з НДС:";          //сумма цен с НДС     
                newWorkSheet.Cells[position + 1, 4] = string.Format("{0:0.00}", sum);
                newWorkSheet.Cells[position + 1, 9] = "Сумма з націнкою:";        //сумма цен с НДС и наценкой           
                newWorkSheet.Cells[position + 1, 11] = $"= SUM(R[-{position - 5}]C: R[-1]C)";
                newWorkSheet.Cells[position + 1, 11].HorizontalAlignment = Excel.Constants.xlCenter;//выравнивание с лева по горизонтали
                newWorkSheet.Cells[position + 2, 2] = "Коефіцієнт націнки:";
                newWorkSheet.Cells[position + 2, 2].Font.Size = 10;//размер шрифта
                newWorkSheet.Cells[position + 2, 4] = "=ROUND(R[-1]C[7]/R[-1]C,2)";
                newWorkSheet.Cells[position + 3, 2] = "% Прибутку:";
                newWorkSheet.Cells[position + 3, 4] = "=IFERROR(ROUND((R[-2]C[7]-R[-2]C)*100/R[-2]C[7],2),0)";
                newWorkSheet.Range[newWorkSheet.Cells[position + 1, 4], newWorkSheet.Cells[position + 3, 4]].HorizontalAlignment = Excel.Constants.xlLeft;//выравнивание с лева по горизонтали

                string nameSupplier = oldWorkSheet.Cells[2, 4].Value;//полное название фирмы поставщика
                string supplier = null; //название фирмы поставщика без знака /"/

                for (int y = 0; y < nameSupplier.Length; y++)//удаление из имени фирмы поставщика знака "
                {
                    if (nameSupplier[y] != '"')
                    {
                        supplier = supplier + nameSupplier[y];
                    }
                }
                open.Table(newWorkSheet, oldWorkSheet, sum, supplier);//создание шапки таблицы     
                newWorkSheet.SaveAs(open.Name(supplier).FullName);//сохранить документ
                open.Number = 0;
                backgroundWorker.ReportProgress(i + 1);//прогрес копирования
            }

            ObjExcel.Quit(); //окончить роботу с Excel 
            MessageBox.Show("Обработка документов успешно оконченна");
        }
    }

    class Document
    {
        int number;
        public int Number { set { number = value; } }
        
        public FileInfo Name(string supplier)//уникальное имя нового документа
        {
            FileInfo file;
            DirectoryInfo directory = new DirectoryInfo("Накладные");
            DateTime data = DateTime.Now;
            do
            {
                number++;
                string name = (supplier + " №" + number + " " + data.Day + "." + data.Month + ".xlsx").ToString();//имя нового документа
                file = new FileInfo($@"{ directory.FullName }\{ name }"); //полный адрес нового документа
            }

            while (file.Exists);//проверка на наличие документа с таким же именем           
            return file;
        }

        public void Table(Excel.Worksheet newWorkSheet, Excel.Worksheet oldWorkSheet, double sum, string supplier)//шапка таблицы
        {
            newWorkSheet.Cells[1, 4] = $"НАКЛАДНА № : {oldWorkSheet.Cells[2, 1].Value} від {string.Format("{0:d}", DateTime.Now)}";
            newWorkSheet.Cells[2, 1] = $"Фірма: {supplier}";
            newWorkSheet.Cells[3, 1] = $"Адрес: {oldWorkSheet.Cells[3, 1].Value}";

            newWorkSheet.Cells[5, 1] = "Назва препарату";
            newWorkSheet.Range[newWorkSheet.Cells[5, 1], newWorkSheet.Cells[5, 4]].Merge();//обэдинение ячеек
            newWorkSheet.Cells[5, 5] = "Кількість, шт.";
            newWorkSheet.Cells[5, 6] = "Ціна 1 шт. с НДС";
            newWorkSheet.Cells[5, 7] = "Ціна всього";
            newWorkSheet.Cells[5, 8] = "Кофецієнт націнки НДС";
            newWorkSheet.Cells[5, 9] = "Кофецієнт націнки";
            newWorkSheet.Cells[5, 10] = "Ціна з націнкою";
            newWorkSheet.Cells[5, 11] = "Сумма з націнкою";

            Excel.Range format = newWorkSheet.Range[newWorkSheet.Cells[5, 1], newWorkSheet.Cells[5, 11]];//диапазон который необходимо форматировать
            format.HorizontalAlignment = Excel.Constants.xlCenter;//выравнивание по центру по горизонтали
            format.VerticalAlignment = Excel.Constants.xlCenter;//выравнивание по центру по вертикали
            format.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//создание таблицы

            format.RowHeight = 40;//высота строки
            format.WrapText = true;//перенос слова
            format.Font.Size = 9;//шрифт           
        }
    }
}
