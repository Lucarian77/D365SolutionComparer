#if DEBUG
using System;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace D365SolutionComparer
{
    internal sealed class Type92EvidenceResultsForm : Form
    {
        private readonly string evidence;

        internal Type92EvidenceResultsForm(string evidence)
        {
            this.evidence = evidence ?? throw new ArgumentNullException(nameof(evidence));
            Text = "Type 92 Evidence - Debug Only";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(1250, 780);
            var body = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                WordWrap = false,
                Font = new Font("Consolas", 9F),
                Text = evidence
            };
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                FlowDirection = FlowDirection.RightToLeft
            };
            var close = new Button { Text = "Close", DialogResult = DialogResult.OK, Width = 90 };
            var save = new Button { Text = "Save Evidence...", Width = 120 };
            save.Click += (sender, args) => SaveEvidence();
            buttons.Controls.Add(close);
            buttons.Controls.Add(save);
            Controls.Add(body);
            Controls.Add(buttons);
            AcceptButton = close;
            CancelButton = close;
        }

        private void SaveEvidence()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                DefaultExt = "txt",
                AddExtension = true,
                FileName = "type92-evidence-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmssfff",
                    CultureInfo.InvariantCulture) + ".txt",
                Title = "Save Type 92 Evidence"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    System.IO.File.WriteAllText(dialog.FileName, evidence, new UTF8Encoding(true));
                    MessageBox.Show(this, "Type 92 evidence was saved.", "Type 92 Evidence",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Failed to save Type 92 evidence.\n\n" + ex.Message,
                        "Type 92 Evidence", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
#endif
