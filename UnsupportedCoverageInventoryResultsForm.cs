#if DEBUG
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace D365SolutionComparer
{
    internal sealed class UnsupportedCoverageInventoryResultsForm : Form
    {
        internal UnsupportedCoverageInventoryResultsForm(string report)
        {
            if (report == null) throw new ArgumentNullException(nameof(report));
            Text = "Unsupported Coverage Inventory - Debug Only";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(1250, 780);
            MinimumSize = new Size(700, 400);
            var body = new RichTextBox
            {
                Dock = DockStyle.Fill, ReadOnly = true, WordWrap = false,
                Font = new Font("Consolas", 9F), Text = report
            };
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom, Height = 44, FlowDirection = FlowDirection.RightToLeft
            };
            var close = new Button { Text = "Close", DialogResult = DialogResult.OK, Width = 90 };
            var save = new Button { Text = "Save Report...", Width = 120 };
            save.Click += (sender, args) =>
            {
                using (var dialog = new SaveFileDialog
                {
                    Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*", DefaultExt = "txt", AddExtension = true,
                    FileName = "unsupported-coverage-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmssfff", CultureInfo.InvariantCulture) + ".txt"
                })
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    try { File.WriteAllText(dialog.FileName, report, new UTF8Encoding(true)); }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Failed to save inventory.\n\n" + ex.Message,
                            "Unsupported Coverage Inventory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };
            buttons.Controls.Add(close); buttons.Controls.Add(save);
            Controls.Add(body); Controls.Add(buttons);
            AcceptButton = close; CancelButton = close;
        }
    }
}
#endif
