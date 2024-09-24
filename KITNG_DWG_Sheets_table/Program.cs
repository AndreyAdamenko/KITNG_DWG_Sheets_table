using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(KITNG_DWG_Sheets_table.Program))]

namespace KITNG_DWG_Sheets_table
{
    public class Program : IExtensionApplication
    {
        [CommandMethod("ChangeKITNGNumAttribute")]
        public void ChangeKITNGNumAttribute()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            // Создаем и показываем форму для выбора и сортировки файлов
            List<string> drawingFiles = null;
            using (Main mainForm = new Main())
            {
                Application.ShowModalDialog(mainForm);

                if (mainForm.DialogResult != DialogResult.OK)
                {
                    ed.WriteMessage("\nОперация отменена пользователем.");
                    return;
                }

                drawingFiles = mainForm.SortedFiles;

                if (drawingFiles == null || drawingFiles.Count == 0)
                {
                    ed.WriteMessage("\nНет выбранных файлов для обработки.");
                    return;
                }
            }

            // Счетчик для номера листа
            int sheetNumber = 1;

            foreach (string file in drawingFiles)
            {
                try
                {
                    // Проверяем, доступен ли файл для чтения
                    if (IsFileLocked(file))
                    {
                        ed.WriteMessage($"\nФайл занят и не может быть открыт: {file}");
                        continue;
                    }

                    // Открываем чертеж
                    Document doc = Application.DocumentManager.Open(file, false);
                    ed.WriteMessage($"\nОткрыт чертеж: {file}");

                    using (doc.LockDocument())
                    {
                        Database db = doc.Database;

                        // Начинаем транзакцию
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            DBDictionary layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                            foreach (DBDictionaryEntry entry in layoutDict)
                            {
                                Layout layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);

                                // Пропускаем модельное пространство
                                if (layout.ModelType) continue;

                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                                foreach (ObjectId id in btr)
                                {
                                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                                    if (ent is BlockReference blockRef)
                                    {
                                        // Проверяем имя блока
                                        BlockTableRecord blockDef = (BlockTableRecord)tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead);
                                        if (blockDef.Name == "KITNGNum")
                                        {
                                            // Изменяем значение атрибута с именем _NUM
                                            foreach (ObjectId attId in blockRef.AttributeCollection)
                                            {
                                                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                                                if (attRef != null && attRef.Tag == "_NUM")
                                                {
                                                    attRef.TextString = sheetNumber.ToString();
                                                }
                                            }
                                        }
                                    }
                                }

                                sheetNumber++; // Увеличиваем номер листа даже если файл не открылся или был пропущен
                            }

                            // Сохраняем изменения
                            tr.Commit();
                        }
                    }

                    // Сохраняем и закрываем документ
                    doc.CloseAndSave(file);
                    ed.WriteMessage($"\nЧертеж сохранен: {file}");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nОшибка при обработке файла {file}: {ex.Message}");
                    // Увеличиваем номер листа, даже если возникла ошибка при обработке
                    sheetNumber++;
                    continue;
                }
            }

            ed.WriteMessage("\nОбработка завершена.");
        }

        // Метод для проверки, занят ли файл
        private bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        public void Initialize()
        {
            // Инициализация при загрузке плагина
        }

        public void Terminate()
        {
            // Действия при выгрузке плагина
        }
    }
}
