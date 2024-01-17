using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Datagrid_to_txt_and_csv_converter
{
    public partial class FormConvert : Form
    {
        public FormConvert()
        {
            InitializeComponent();

            // добавление столбцов в dataGridView
            dataGridView.Columns.Add("surname", "Фамилия");
            dataGridView.Columns.Add("name", "Имя");
            dataGridView.Columns.Add("patronymic", "Отчество");

            // генерируем тестовые данные
            for (int i = 0; i < 100; i++)
            {
                dataGridView.Rows.Add(new DataGridViewRow());
                for (int j = 0; j < 3; j++)
                {
                    dataGridView.Rows[i].Cells[j].Value = $"{i}test{j}";
                }
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            // создаем диалог выбора файла
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // путь к выбранному файлу
                string filePath = openFileDialog.FileName;

                // проверяем расширение файла и определяем формат экспорта
                switch (Path.GetExtension(filePath).ToLower())
                {
                    case ".txt":
                        ExportToTextFile(filePath);
                        break;
                    case ".csv":
                        ExportToCSVFile(filePath);
                        break;
                }
            }
        }

        private void ExportToTextFile(string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.GetEncoding(1251)))
            {
                // записываем заголовок таблицы в файл
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    sw.Write($"{column.HeaderText}\t");
                }

                sw.WriteLine();

                // проходим по каждой строке таблицы
                foreach (DataGridViewRow row in dataGridView.Rows)
                {

                    // записываем значение ячейки строки в файл
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        sw.Write($"{row.Cells[i].Value}\t");
                    }
                    sw.WriteLine();

                }
            }
        }

        private void ExportToCSVFile(string filePath)
        {
            // создаем строку
            StringBuilder sb = new StringBuilder();

            // проходим по каждому столбцу таблицы
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                // записываем название столбца в строку
                sb.Append(column.HeaderText + " ");

                // ставим разделитель
                sb.Append(';');
            }

            // делаем перенос на следующую строку
            sb.Length--;
            sb[sb.Length - 1] = '\n';

            // проходим по каждой строке таблицы
            for (int rowIndex = 0; rowIndex < dataGridView.RowCount; rowIndex++)
            {
                // проходим по каждой ячейке строки таблицы
                for (int colIndex = 0; colIndex < dataGridView.ColumnCount; colIndex++)
                {
                    // записываем значение ячейки в строку
                    DataGridViewRow row = dataGridView.Rows[rowIndex];
                    sb.Append(row.Cells[colIndex].Value.ToString() + " ");

                    // ставим разделитель
                    sb.Append(';');
                }

                // делаем перенос на следующую строку
                sb.Length--;
                sb[sb.Length - 1] = '\n';
            }

            // записываес строку в файл
            File.WriteAllText(filePath, sb.ToString(), Encoding.GetEncoding(1251));
        }
    }
}
