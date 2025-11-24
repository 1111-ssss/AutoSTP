namespace core
{
    public static class Constants
    {

            // Единицы
            public const float MillimetersToPoints = 2.835f;
            public const float CentimetersToPoints = 28.35f;

            // Общие
            public const string FontName = "Times New Roman";
            public const float MainFontSize = 14f;
            public const float MainLineSpacing = 1.0f;
            public const float FirstLineIndent = 1.25f * CentimetersToPoints; // 35.44f

            // Заголовки 1-го уровня
            public const float Heading1Size = 16f;
            public const float Heading1AfterSpacing = 6f;

            // Заголовки 2-го уровня
            public const float Heading2Size = 14f;
            public const float Heading2BeforeSpacing = 6f;

        // Таблицы
            public const float TableFontSize = 14f;
            public const float MinTableRowHeightMm = 8f;
            public const float MinTableRowHeight = 8f * MillimetersToPoints; // 22.68f
            public const float TableSpacing = 6f; // или 6f

            // Иллюстрации
            public const float FigureSpacing = 6f; // или 6f

            // Поля (для Interop)
            public const float LeftMarginMm = 30f;
            public const float RightMarginMm = 15f;
            public const float TopMarginMm = 20f;
            public const float BottomMarginMm = 20f;
        
    }
}
