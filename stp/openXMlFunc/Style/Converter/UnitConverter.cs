namespace openXMlFunc.Converter
{
    /// <summary>
    /// Класс для конвертации значений
    /// </summary>
    public static class UnitConverter
    {
        /// <summary>
        /// Преобразует пункты (pt) в half-points (используется в FontSize)
        /// </summary>
        public static string PtToHalfPoints(double pt) => ((int)(pt * 2)).ToString();

        /// <summary>
        /// Преобразует сантиметры (cm) в twips (1 cm = 567 twips)
        /// </summary>
        public static string CmToTwips(double cm) => ((int)(cm * 567)).ToString();

        /// <summary>
        /// Преобразует пункты (pt) в twips (1 pt = 20 twips)
        /// </summary>
        public static string PtToTwips(double pt) => ((int)(pt * 20)).ToString();
    }
}