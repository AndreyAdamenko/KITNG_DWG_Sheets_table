using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace KITNG_DWG_Sheets_table
{
    public partial class Main : Form
    {
        public List<string> SortedFiles { get; private set; }

        private int dragIndex = -1;
        private bool isDragging = false;

        public Main()
        {
            InitializeComponent();

            // Настройка listBox1 для поддержки drag-and-drop
            listBox1.AllowDrop = true;
            listBox1.MouseDown += new MouseEventHandler(listBox1_MouseDown);
            listBox1.MouseMove += new MouseEventHandler(listBox1_MouseMove);
            listBox1.DragOver += new DragEventHandler(listBox1_DragOver);
            listBox1.DragDrop += new DragEventHandler(listBox1_DragDrop);

            // Подключение обработчика для button2
            button2.Click += new EventHandler(button2_Click);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var fileDialog = new OpenFileDialog
            {
                Filter = "DWG файлы (*.dwg)|*.dwg",
                Multiselect = true,
                Title = "Выберите файлы DWG"
            };

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in fileDialog.FileNames)
                {
                    if (!listBox1.Items.Contains(file))
                    {
                        listBox1.Items.Add(file);
                    }
                }
            }
        }

        // Обработчик для начала перетаскивания
        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBox1.Items.Count == 0)
                return;

            dragIndex = listBox1.IndexFromPoint(e.X, e.Y);
            if (dragIndex != -1)
            {
                isDragging = true;
            }
        }

        // Обработчик для перетаскивания элемента
        private void listBox1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging && dragIndex != -1)
            {
                listBox1.DoDragDrop(listBox1.Items[dragIndex], DragDropEffects.Move);
                isDragging = false;
            }
        }

        // Обработчик для отображения эффекта перетаскивания
        private void listBox1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        // Обработчик для завершения перетаскивания и обновления списка
        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            Point point = listBox1.PointToClient(new Point(e.X, e.Y));
            int index = listBox1.IndexFromPoint(point);

            if (index < 0)
                index = listBox1.Items.Count - 1;

            object data = e.Data.GetData(typeof(string));

            listBox1.Items.Remove(data);
            listBox1.Items.Insert(index, data);
        }

        // Обработчик для кнопки button2
        private void button2_Click(object sender, EventArgs e)
        {
            SortedFiles = new List<string>();
            foreach (var item in listBox1.Items)
            {
                SortedFiles.Add(item.ToString());
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
