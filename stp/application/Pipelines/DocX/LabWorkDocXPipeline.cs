using core.Enums;
using infrastructure.Context;
using infrastructure.DocX.Lists;
using infrastructure.DocX.MainText;
using infrastructure.DocX.Pictures;
using infrastructure.DocX.Tables;
using infrastructure.Utils.UtilsDocX.Validate;
using logger.Logger;
using core.Interfaces;
using System.Text.RegularExpressions;
using core.Model;


namespace application.Pipelines.DocX
{
    public class LabWorkDocXPipeline : IPipeline
    {
        private readonly DocXContext _context;
        public LabWorkDocXPipeline(DocXContext doc, LabWorkOptions options)
        {
            _context = doc;
            MainTextStyle.allTextStyle(_context.Doc);
        }


        public void StartPipeline()
        {


            foreach (var item in _context.Doc.Paragraphs)
            {
                try
                {

                    if (!ParagraphValidate.IsSpecialParagraph(item))
                    {
                        MainTextStyle.BaseTextStyle(item);
                    }



                    if (item.Pictures.Any())
                    {
                        Pictures.validate(item);
                    }



                    if (item.IsListItem)
                    {
                        ListStyle.ProcessList(item);
                        continue;
                    }


                    else if (Regex.IsMatch(item.Text.Trim().ToLower(), @"^рисунок \d+\.\d+ –", RegexOptions.IgnoreCase))
                    {
                        Logger.Debug($"Подпись рисунка для форматирования найдена : {item.Text} ");
                        Pictures.PictureNameStyle(item);
                        continue;
                    }

                    else if (item.Text.Trim().ToLower().StartsWith("задания для выполнения работы"))
                    {
                        MainTextStyle.HeaderOneLevel(item);
                        continue;
                    }


                    else if (Regex.IsMatch(item.Text.Trim().ToLower(), @"^таблица \d+\.\d+ – ", RegexOptions.IgnoreCase))
                    {
                        Logger.Debug($"Подпись таблицы для форматирования найдена : {item.Text} ");
                        MainTable.TableNameStyle(item);
                        continue;
                    }


                    else if (Regex.IsMatch(item.Text.Trim(), @"^лабораторная\s+работа\s+№?\s*\d+", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(item.Text.Trim(), @"^практическая\s+работа\s+№?\s*\d+", RegexOptions.IgnoreCase)
                        )
                    {
                        MainTextStyle.HeaderOneLevel(item, true);

                    }


                }

                catch (OverflowException ex)
                {
                    Logger.Log($"Пропущено изображение: {ex.Message}");
                    continue;
                }
                catch (Exception ex)
                {
                    Logger.Error(ex.Message);
                }

            }

            foreach (var item in _context.Doc.Tables)
            {
                MainTable.BaseTableStyle(item);

            }

        }


        public void SaveAS(string TargetPath)
        {
            try
            {
                _context.Doc.SaveAs(TargetPath);
            }
            catch (Exception ex)
            {
                Logger.Error($"ошибка сохранения в новый файл {ex.Message} ");
            }

        }
        public void Save()
        {
            try
            {
                _context.Doc.Save();
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка сохранения в текущий файл {ex.Message} ");
            }

        }


    }
}
