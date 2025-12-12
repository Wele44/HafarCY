using System.Collections.Generic;
using System.Drawing;

namespace ExcelClassificationEnforcer.Core
{
    public enum ClassificationLevel
    {
        TopSecret,   // سري للغاية
        Secret,      // سري
        Restricted,  // مقيّد
        Public       // عام
    }

    public static class ClassificationConfig
    {
        public const string PropertyName = "DocumentClassification";

        public static readonly Dictionary<ClassificationLevel, string> Labels =
            new Dictionary<ClassificationLevel, string>
            {
                { ClassificationLevel.TopSecret,  "سري للغاية" },
                { ClassificationLevel.Secret,     "سري" },
                { ClassificationLevel.Restricted, "مقيّد" },
                { ClassificationLevel.Public,     "عام" }
            };

        // الألوان المطلوبة (Dark Red / Red / Orange / Green)
        public static readonly Dictionary<ClassificationLevel, Color> Colors =
            new Dictionary<ClassificationLevel, Color>
            {
                { ClassificationLevel.TopSecret,  Color.FromArgb(0x8B, 0x00, 0x00) }, // Dark Red
                { ClassificationLevel.Secret,     Color.FromArgb(0xFF, 0x00, 0x00) }, // Red
                { ClassificationLevel.Restricted, Color.FromArgb(0xFF, 0xA5, 0x00) }, // Orange
                { ClassificationLevel.Public,     Color.FromArgb(0x00, 0x80, 0x00) }  // Green
            };

        public const string FooterSuffix = ""; // غير مستخدم هنا (واترمارك فقط)
    }
}
