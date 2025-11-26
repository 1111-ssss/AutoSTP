using core;
using core.Enums;
using infrastructure.OpenXmlContext;
using logger.Logger;
using infrastructure;
using infrastructure.OpenXML.Metatags;
using infrastructure.OpenXML.Style;
using infrastructure.OpenXML.Margins;
using infrastructure.OpenXML.NamesProcessing;

namespace openXMlFunc.Conveers
{
    public class LabsConveer
    {
        private DocContext _context;
        private PicturesNames _picturesNames;

        public LabsConveer(
            DocContext doc,
            PicturesNames picturesNames
            )
        {
            _context = doc;
            _picturesNames = picturesNames;
        }

        public bool Conveer()
        {
            try
            {
                Margins.SetupPageMargins(_context.Body);
                Logger.Log("OpenXML Conveer: SetupPageMargins is done");
                //Style.Style.EditStyle.CreateBaseStyle(_doc);
                //logger.Logger.Log("OpenXML Conveer: CreateBaseStyle is done");
                _picturesNames.AddPictureNames();
                TableNames.TableName(_context.Body);
                //TableStyle.ApplyBaseTableStyle(_context.Body);
                TextStyle.ApplyBaseTextStyle(_context.Body);
                Logger.Log("OpenXML Conveer: ApplyBaseTextStyle is done");
                Logger.Log("OpenXML Conveer: TableNames is done");
                Metatags.AddMetatags(_context, author: "AutoSTP", description: "Document formatted with private docx formatter script");
                Logger.Log("OpenXML Conveer: AddMultipleStyles is done");
                StylesCreator.AddMultipleStyles(_context.MainDocumentPart, new Dictionary<string, string>() {
                    {"stp1", "1 Automatical"},
                    {"stp2", "2 STP"},
                    {"stp3", "3 Formatter"},
                    {"stp4", "4 Script"},
                    {"stp5", "5 Касевич"},
                    {"stp6", "6 Гурский"}
                });
                Logger.Log("OpenXML Conveer: AddMetatags is done");
            }
            catch (Exception ex)
            {
                Logger.Log($"Formatting - {ex.ToString()}", LoggerState.Error);
                return false;
            }
            finally
            {
                Logger.Log("OpenXML: file formatted.");
            }
            return true;
        }
        public bool Save()
        {
            try
            {
                _context.Doc.Save();
            }
            catch (Exception ex)
            {
                Logger.Log($"Saving - {ex.ToString()}", LoggerState.Error);
                return false;
            }
            finally
            {
                Logger.Log("OpenXML: file saved.");
            }
            return true;
        }
    }
}