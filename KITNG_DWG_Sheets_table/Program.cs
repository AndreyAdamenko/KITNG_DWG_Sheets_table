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
        [CommandMethod("KITNG_Sheets_table_numerate")]
        public void ChangeKITNGNumAttribute()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            // Создаем и показываем форму для выбора и сортировки файлов
            List<string> drawingFiles = null;
            int startNumber = 0;
            using (Main mainForm = new Main())
            {
                Application.ShowModalDialog(mainForm);

                if (mainForm.DialogResult != DialogResult.OK)
                {
                    ed.WriteMessage("\nОперация отменена пользователем.");
                    return;
                }

                startNumber = mainForm.StartNumber;
                drawingFiles = mainForm.SortedFiles;

                if (drawingFiles == null || drawingFiles.Count == 0)
                {
                    ed.WriteMessage("\nНет выбранных файлов для обработки.");
                    return;
                }
            }

            // Счетчики для итогового сообщения
            int totalFiles = 0;
            int totalSheets = 0;
            int skippedFiles = 0;
            int skippedSheets = 0;
            List<string> filesWithoutBlock = new List<string>();

            // Счетчик для номера листа
            int sheetNumber = startNumber;

            // Строка для записи диапазонов номеров листов в файлы
            List<string> fileSheetRanges = new List<string>();

            // Проходим по файлам
            for (int curDwgIndex = 0; curDwgIndex < drawingFiles.Count; curDwgIndex++)
            {
                string file = drawingFiles[curDwgIndex];
                int initialSheetNumber = sheetNumber; // Начальный номер для файла
                string blockAttributeCombined = null; // Переменная для объединенной строки атрибутов

                Database acDb = new Database(false, true);
                bool fileIsChanged = false;
                try
                {
                    // Открываем файл DWG
                    acDb.ReadDwgFile(file, FileOpenMode.OpenForReadAndWriteNoShare, false, "");
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    if (ex.ErrorStatus == ErrorStatus.FileSharingViolation)
                    {
                        ed.WriteMessage($"\nНе удалось обработать файл \"{file}\". Возможно, он открыт другим пользователем.");
                        skippedFiles++;
                        acDb.Dispose();
                        continue;
                    }
                    else
                    {
                        throw ex;
                    }
                }

                using (Transaction tr = acDb.TransactionManager.StartTransaction())
                {
                    try
                    {
                        DBDictionary layoutDict = (DBDictionary)tr.GetObject(acDb.LayoutDictionaryId, OpenMode.ForRead);

                        bool blockFoundInFile = false; // Флаг, указывающий найден ли блок в файле

                        foreach (DBDictionaryEntry entry in layoutDict)
                        {
                            Layout layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);

                            // Пропускаем модельное пространство
                            if (layout.ModelType) continue;

                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                            bool blockFoundInSheet = false;

                            foreach (ObjectId id in btr)
                            {
                                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                                if (ent is BlockReference blockRef)
                                {
                                    // Проверяем имя блока
                                    BlockTableRecord blockDef = (BlockTableRecord)tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead);
                                    if (blockDef.Name == "KITNGNum")
                                    {
                                        blockFoundInSheet = true;
                                        blockFoundInFile = true;
                                        fileIsChanged = true;

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
                                    else if (blockDef.Name == "KITNGMainA" && string.IsNullOrEmpty(blockAttributeCombined))
                                    {
                                        // Если это блок "KITNGMainA" и объединенная строка атрибутов еще не заполнена
                                        string attr1 = "", attr2 = "", attr3 = "";
                                        foreach (ObjectId attId in blockRef.AttributeCollection)
                                        {
                                            AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                            if (attRef != null)
                                            {
                                                if (attRef.Tag == "_GTOST_1_DWGNAME_1") attr1 = attRef.TextString?.Trim();
                                                if (attRef.Tag == "_GTOST_1_DWGNAME_2") attr2 = attRef.TextString?.Trim();
                                                if (attRef.Tag == "_GTOST_1_DWGNAME_3") attr3 = attRef.TextString?.Trim();
                                            }
                                        }

                                        // Проверяем, что хотя бы один из атрибутов не пустой
                                        if (!string.IsNullOrWhiteSpace(attr1) || !string.IsNullOrWhiteSpace(attr2) || !string.IsNullOrWhiteSpace(attr3))
                                        {
                                            blockAttributeCombined = attr1 + attr2 + attr3; // Объединяем атрибуты
                                        }
                                    }
                                }
                            }

                            // Если блок не был найден на листе, увеличиваем количество пропущенных листов
                            if (!blockFoundInSheet)
                            {
                                skippedSheets++;
                            }

                            sheetNumber++; // Увеличиваем номер листа
                            totalSheets++;
                        }

                        if (fileIsChanged)
                        {
                            tr.Commit();
                        }
                        else
                        {
                            tr.Abort();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nОшибка при обработке файла {file}: {ex.Message}");
                        tr.Abort();
                        skippedFiles++;
                        continue;
                    }
                }

                // Если блок KITNGMainA не был найден или атрибуты пусты, используем имя файла
                if (string.IsNullOrEmpty(blockAttributeCombined))
                {
                    blockAttributeCombined = Path.GetFileNameWithoutExtension(file); // Используем имя файла без расширения как запасной вариант
                }

                var fileName = Path.GetFileNameWithoutExtension(file);

                // Формируем строку отчета в зависимости от количества страниц
                int finalSheetNumber = sheetNumber - 1; // Конечный номер листа для файла
                string sheetRange = (initialSheetNumber == finalSheetNumber)
                    ? $"{fileName} - {blockAttributeCombined} - {initialSheetNumber}" // Если одна страница, просто указываем номер
                    : $"{fileName} - {blockAttributeCombined} - {initialSheetNumber}-{finalSheetNumber}"; // Диапазон для нескольких страниц
                fileSheetRanges.Add(sheetRange);

                if (fileIsChanged)
                {
                    try
                    {
                        acDb.SaveAs(file, true, DwgVersion.Newest, acDb.SecurityParameters);
                        ed.WriteMessage($"\nЧертеж сохранен: {file}");
                        totalFiles++; // Увеличиваем количество успешно обработанных файлов
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nОшибка при сохранении файла {file}: {ex.Message}");
                        skippedFiles++;
                    }
                }
                else
                {
                    ed.WriteMessage($"\nИзменений не внесено в файл: {file}");
                }

                acDb.Dispose();
            }

            // Записываем данные в текстовый файл в папку temp
            string tempFilePath = Path.Combine(Path.GetTempPath(), $"KITNG_DWG_Sheets_table_SheetRanges_{DateTime.Now.Ticks}.txt");
            try
            {
                File.WriteAllLines(tempFilePath, fileSheetRanges);
                ed.WriteMessage($"\nРезультаты сохранены в файл: {tempFilePath}");

                // Открываем файл через Shell после создания
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = tempFilePath,
                    UseShellExecute = true // Запуск через Shell
                });
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nОшибка при сохранении результатов в файл: {ex.Message}");
            }

            // Формируем итоговое сообщение
            string resultMessage = "Обработка завершена.\n";

            resultMessage += $"\nФайлов: {totalFiles}";
            resultMessage += $"\nЛистов: {totalSheets}";

            if (skippedFiles > 0)
            {
                resultMessage += $"\nПропущено файлов: {skippedFiles}";
            }

            if (skippedSheets > 0)
            {
                resultMessage += $"\nПропущено листов (нет блока номера): {skippedSheets}";
            }

            if (filesWithoutBlock.Count > 0)
            {
                resultMessage += "\nФайлы, в которых не был найден блок 'KITNGNum':\n";
                resultMessage += string.Join("\n", filesWithoutBlock);
            }

            // Выводим итоговое сообщение
            MessageBox.Show(resultMessage, "Результаты выполнения");
            ed.WriteMessage(resultMessage);
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
