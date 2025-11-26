using infrastructure.Context;
using infrastructure.DocX.Style.lists;
using infrastructure.DocX.Style.MainText;
using infrastructure.DocX.Style.Pictures;
using infrastructure.DocX.Style.Tabl;
using infrastructure.utils.utilsDocX.Validate;
using logger.Logger;
using System.Text.RegularExpressions;


namespace DocXFunc.Conveers
{
    public class LabsConveer
    {
        private readonly DocXContext _context;
        public LabsConveer(DocXContext doc)
        {
            _context = doc;
            MainTextStyle.allTextStyle(_context.Doc);
        }

        private void TableConveer()
        {
            foreach(var item in _context.Doc.Tables)
            {
                MainTable.BaseTableStyle(item);
                
            }
        }


        public void Conveer()
        {

           
            foreach ( var item in _context.Doc.Paragraphs)
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
                    TableConveer();

                }

                catch (OverflowException ex)
                {
                    Logger.Log($"Пропущено изображение: {ex.Message}");
                    continue; 
                }
                catch( Exception ex )
                {
                    Logger.Error(ex.Message);
                }
               
            }
            
        }


        public void SaveAS(string TargetPath)
        {
            try
            {
                _context.Doc.SaveAs(TargetPath);
            }
            catch(Exception ex)
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
            catch(Exception ex)
            {
                Logger.Error($"Ошибка сохранения в текущий файл {ex.Message} ");
            }
           
        }

       
    }
}
