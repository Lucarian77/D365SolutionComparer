#if DEBUG
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace D365SolutionComparer
{
    internal sealed class Type511EvidenceResultsForm : Form
    {
        internal Type511EvidenceResultsForm(string evidence)
        {
            if (evidence == null) throw new ArgumentNullException(nameof(evidence));
            Text = "Type 511 Team Template Evidence - Debug Only";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(1250, 780);
            MinimumSize = new Size(700, 400);
            var body = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, WordWrap = false,
                Font = new Font("Consolas", 9F), Text = evidence };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, FlowDirection = FlowDirection.RightToLeft };
            var close = new Button { Text = "Close", DialogResult = DialogResult.OK, Width = 90 };
            var save = new Button { Text = "Save Evidence...", Width = 120 };
            save.Click += (sender, args) =>
            {
                using (var dialog = new SaveFileDialog
                {
                    Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*", DefaultExt = "txt", AddExtension = true,
                    FileName = "type511-evidence-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmssfff", CultureInfo.InvariantCulture) + ".txt"
                })
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    try { File.WriteAllText(dialog.FileName, evidence, new UTF8Encoding(true)); }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Failed to save evidence.\n\n" + ex.Message,
                            "Team Template Evidence", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
