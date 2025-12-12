using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelClassificationEnforcer.Core;

namespace ExcelClassificationEnforcer
{
    public partial class ThisAddIn
    {
        private readonly Dictionary<string, bool> _promptedThisSession = new Dictionary<string, bool>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ((Excel.AppEvents_Event)this.Application).WorkbookOpen += Application_WorkbookOpen;
            ((Excel.AppEvents_Event)this.Application).WorkbookBeforeSave += Application_WorkbookBeforeSave;
            ((Excel.AppEvents_Event)this.Application).WorkbookBeforeClose += Application_WorkbookBeforeClose;
            ((Excel.AppEvents_Event)this.Application).NewWorkbook += Application_NewWorkbook;
        }

        private void Application_NewWorkbook(Excel.Workbook wb)
        {
            ResetPromptFlag(wb);
            var cur = ReadClassification(wb);
            if (cur.HasValue) UpdateWatermarks(wb, cur.Value);
        }

        private void Application_WorkbookOpen(Excel.Workbook wb)
        {
            ResetPromptFlag(wb);
            var cur = ReadClassification(wb);
            if (cur.HasValue) UpdateWatermarks(wb, cur.Value);
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook wb, bool SaveAsUI, ref bool Cancel)
        {
            if (!PromptAndEnsureClassification(wb)) { Cancel = true; return; }
            var lvl = ReadClassification(wb);
            if (lvl.HasValue) UpdateWatermarks(wb, lvl.Value);
            MarkPrompted(wb);
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {
            if (WasPrompted(wb)) return;

            bool hasUnsaved = false;
            try { hasUnsaved = !wb.Saved; } catch { }

            if (!hasUnsaved) return;

            if (!PromptAndEnsureClassification(wb)) { Cancel = true; return; }
            var lvl = ReadClassification(wb);
            if (lvl.HasValue) UpdateWatermarks(wb, lvl.Value);
            MarkPrompted(wb);
        }

        private bool PromptAndEnsureClassification(Excel.Workbook wb)
        {
            var current = ReadClassification(wb);

            using (var dlg = new ClassificationDialog(current))
            {
                var res = dlg.ShowDialog();
                if (res != DialogResult.OK || !dlg.Selected.HasValue) return false;
                current = dlg.Selected.Value;
                WriteClassification(wb, current.Value);
            }
            return true;
        }

        internal ClassificationLevel? ReadClassification(Excel.Workbook wb)
        {
            try
            {
                var props = wb.CustomDocumentProperties as Office.DocumentProperties;
                foreach (Office.DocumentProperty p in props)
                    if (p.Name == ClassificationConfig.PropertyName)
                    {
                        ClassificationLevel level;
                        var val = p.Value == null ? null : p.Value.ToString();
                        if (Enum.TryParse(val, out level)) return level;
                    }
            }
            catch { }
            return null;
        }

        internal void WriteClassification(Excel.Workbook wb, ClassificationLevel lvl)
        {
            var props = wb.CustomDocumentProperties as Office.DocumentProperties;
            try { props[ClassificationConfig.PropertyName].Delete(); } catch { }
            props.Add(ClassificationConfig.PropertyName, false,
                      Office.MsoDocProperties.msoPropertyTypeString, lvl.ToString());
        }

        // ====== تحديث الواترماركات ======
        // - واترمارك التصنيف: ملوّن (حسب المستوى) ويظهر مرتين
        // - واترمارك الاسم: رمادي، بلا خلفية، الكتابة نفسها شفافة، ويظهر 3 مرات
        private void UpdateWatermarks(Excel.Workbook wb, ClassificationLevel lvl)
        {
            string label = ClassificationConfig.Labels[lvl];
            int rgbClass = ColorTranslator.ToOle(ClassificationConfig.Colors[lvl]);

            // اسم آخر من عدّل
            string user;
            try { user = this.Application.UserName; if (string.IsNullOrWhiteSpace(user)) user = Environment.UserName; }
            catch { user = Environment.UserName; }

            string classText = "التصنيف : (" + label + ")";
            string userText = "آخر من عدّل: " + user + " — " + DateTime.Now.ToString("yyyy/MM/dd HH:mm");

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                var shapes = ws.Shapes;

                // نظّف آثارنا السابقة
                try
                {
                    for (int i = shapes.Count; i >= 1; i--)
                    {
                        var shp = shapes.Item(i);
                        if (shp.Name != null && (shp.Name.StartsWith("ClassificationWatermark_") || shp.Name.StartsWith("EditorWatermark_")))
                            shp.Delete();
                    }
                }
                catch { }

                // الهوامش
                float leftM = (float)ws.PageSetup.LeftMargin;
                float rightM = (float)ws.PageSetup.RightMargin;
                float topM = (float)ws.PageSetup.TopMargin;
                float bottomM = (float)ws.PageSetup.BottomMargin;

                // ===== 1) واترمارك "التصنيف" (مرتين) =====
                var classPoints = new[]
                {
                    new PointF(leftM + 60f,  topM + 60f),   // أعلى-يسار
                    new PointF(leftM + 320f, topM + 360f)   // منتصف باتجاه أسفل-يمين
                };

                float classFont = 28f;
                float classTrans = 0.65f;
                float rot = 315f;

                for (int i = 0; i < classPoints.Length; i++)
                {
                    var shape = shapes.AddTextEffect(
                        Office.MsoPresetTextEffect.msoTextEffect1,
                        classText,
                        "Segoe UI",
                        classFont,
                        Office.MsoTriState.msoTrue,   // Bold
                        Office.MsoTriState.msoFalse,  // Italic
                        classPoints[i].X,
                        classPoints[i].Y
                    );
                    shape.Name = $"ClassificationWatermark_{i}";
                    shape.Rotation = rot;

                    // لا خلفية للـShape، ولكن نلوّن نص التصنيف عبر Fill (للـWordArt)
                    shape.Fill.Visible = Office.MsoTriState.msoFalse; // تأكيد عدم وجود خلفية
                    shape.Line.Visible = Office.MsoTriState.msoFalse;

                    // لون النص عبر TextFrame2 (أوضح وأدق)
                    var tf = shape.TextFrame2.TextRange.Font.Fill;
                    tf.Visible = Office.MsoTriState.msoTrue;
                    tf.Solid();
                    tf.ForeColor.RGB = rgbClass;
                    tf.Transparency = classTrans;

                    shape.Placement = Excel.XlPlacement.xlFreeFloating;
                    shape.ZOrder(Office.MsoZOrderCmd.msoSendBehindText);
                    shape.LockAspectRatio = Office.MsoTriState.msoTrue;
                }

                // ===== 2) واترمارك "اسم آخر معدّل" (ثلاث مرات) =====
                // رمادي، أصغر، بدون أي خلفية، والكتابة نفسها شفافة
                int rgbGray = ColorTranslator.ToOle(Color.FromArgb(0x88, 0x88, 0x88));
                float editorFont = 18f;        // أصغر من التصنيف
                float editorTrans = 0.80f;     // شفافية أعلى (أخف)
                var editorPoints = new[]
                {
                    new PointF(leftM + 100f, topM + 30f),   // أعلى (وسط الطرف الأيسر قليلًا)
                    new PointF(leftM + 360f, topM + 230f),  // وسط مائل بعيد عن التصنيف
                    new PointF(leftM + 120f, topM + 420f)   // أسفل باتجاه اليسار
                };

                for (int i = 0; i < editorPoints.Length; i++)
                {
                    var shape = shapes.AddTextEffect(
                        Office.MsoPresetTextEffect.msoTextEffect1,
                        userText,
                        "Segoe UI",
                        editorFont,
                        Office.MsoTriState.msoFalse,  // غير عريض
                        Office.MsoTriState.msoFalse,
                        editorPoints[i].X,
                        editorPoints[i].Y
                    );

                    shape.Name = $"EditorWatermark_{i}";
                    shape.Rotation = rot;

                    // تأكيد: لا خلفية للـShape إطلاقًا
                    shape.Fill.Visible = Office.MsoTriState.msoFalse;
                    shape.Line.Visible = Office.MsoTriState.msoFalse;

                    // لون وشفافية "النص" فقط (بدون خلفية)
                    var tf = shape.TextFrame2.TextRange.Font.Fill;
                    tf.Visible = Office.MsoTriState.msoTrue;
                    tf.Solid();
                    tf.ForeColor.RGB = rgbGray;     // رمادي ثابت
                    tf.Transparency = editorTrans;  // الكتابة نفسها شفافة

                    shape.Placement = Excel.XlPlacement.xlFreeFloating;
                    shape.ZOrder(Office.MsoZOrderCmd.msoSendBehindText);
                    shape.LockAspectRatio = Office.MsoTriState.msoTrue;
                }
            }
        }

        // ====== تتبع حالة المصنف في الجلسة ======
        private string GetWbKey(Excel.Workbook wb)
        {
            try { var path = (wb.FullName ?? "").Trim(); if (!string.IsNullOrEmpty(path)) return path.ToLowerInvariant(); }
            catch { }
            return "unsaved_" + wb.GetHashCode();
        }
        private void ResetPromptFlag(Excel.Workbook wb) { _promptedThisSession[GetWbKey(wb)] = false; }
        private bool WasPrompted(Excel.Workbook wb) { bool v; return _promptedThisSession.TryGetValue(GetWbKey(wb), out v) && v; }
        private void MarkPrompted(Excel.Workbook wb) { _promptedThisSession[GetWbKey(wb)] = true; }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
