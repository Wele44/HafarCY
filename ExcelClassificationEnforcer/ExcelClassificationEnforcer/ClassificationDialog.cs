using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ExcelClassificationEnforcer.Core;

namespace ExcelClassificationEnforcer
{
    // Dialog بواجهة بسيطة فيها ComboBox لاختيار التصنيف (شكل رسمي بدون تلوين)
    public class ClassificationDialog : Form
    {
        private readonly ComboBox cbLevels = new ComboBox();
        private readonly Button btnOk = new Button { Text = "موافق", Width = 100, Height = 34, DialogResult = DialogResult.OK };
        private readonly Button btnCancel = new Button { Text = "إلغاء", Width = 100, Height = 34, DialogResult = DialogResult.Cancel };

        private readonly List<KeyValuePair<ClassificationLevel, string>> _items =
            new List<KeyValuePair<ClassificationLevel, string>>
            {
                new KeyValuePair<ClassificationLevel,string>(ClassificationLevel.TopSecret,  ClassificationConfig.Labels[ClassificationLevel.TopSecret]),
                new KeyValuePair<ClassificationLevel,string>(ClassificationLevel.Secret,     ClassificationConfig.Labels[ClassificationLevel.Secret]),
                new KeyValuePair<ClassificationLevel,string>(ClassificationLevel.Restricted, ClassificationConfig.Labels[ClassificationLevel.Restricted]),
                new KeyValuePair<ClassificationLevel,string>(ClassificationLevel.Public,     ClassificationConfig.Labels[ClassificationLevel.Public]),
            };

        public ClassificationLevel? Selected { get; private set; }

        public ClassificationDialog(ClassificationLevel? current)
        {
            Text = "تحديد مستوى التصنيف";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 10);
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            ClientSize = new Size(380, 160);

            var lbl = new Label { Text = "اختر مستوى التصنيف:", AutoSize = true, Dock = DockStyle.Top, Padding = new Padding(12, 12, 12, 4) };

            cbLevels.DropDownStyle = ComboBoxStyle.DropDownList;  // رسمي
            cbLevels.DisplayMember = "Value";
            cbLevels.ValueMember = "Key";
            foreach (var it in _items) cbLevels.Items.Add(it);
            cbLevels.Dock = DockStyle.Top;
            cbLevels.Margin = new Padding(12);
            cbLevels.Height = 34;

            if (current.HasValue)
            {
                for (int i = 0; i < _items.Count; i++)
                    if (_items[i].Key == current.Value) { cbLevels.SelectedIndex = i; break; }
            }
            else cbLevels.SelectedIndex = 0;

            var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(12), Height = 60 };
            btnPanel.Controls.Add(btnOk);
            btnPanel.Controls.Add(btnCancel);

            Controls.Add(btnPanel);
            Controls.Add(cbLevels);
            Controls.Add(lbl);

            btnOk.Click += delegate
            {
                if (cbLevels.SelectedIndex < 0)
                {
                    MessageBox.Show("الرجاء اختيار مستوى التصنيف.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }
                var kv = (KeyValuePair<ClassificationLevel, string>)cbLevels.SelectedItem;
                Selected = kv.Key;
            };
        }
    }
}
