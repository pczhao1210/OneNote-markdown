using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.Settings;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.OneNote;

namespace OneNoteMarkdown.UI
{
    internal sealed class SettingsDialog : Form
    {
        private TextBox _txtFontFamily;
        private TextBox _txtMonoFont;
        private TextBox _txtMathFont;
        private TextBox _txtParagraphSize;
        private TextBox _txtCodeSize;
        private CheckBox _chkLatex;
        private CheckBox _chkLineNumber;
        private ComboBox _cboPreset;
        private CheckBox _chkPreviewTitle;
        private TextBox _txtPreviewTitle;
        private ComboBox _cboPreviewPosition;
        private TextBox _txtPreviewGap;
        private TextBox _txtPreviewWidth;
        private CheckBox _chkAutoRefresh;
        private TextBox _txtAutoRefreshDelay;
        private CheckBox _chkRemoteImages;
        private ComboBox _cboLanguage;

        public SettingsDialog()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96f, 96f);
            Text = Loc.S("Dialog.Settings.Title");
            ClientSize = new Size(560, 720);
            MinimumSize = new Size(460, 490);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            ShowInTaskbar = true;
            TopMost = true;
            MaximizeBox = false;
            Icon = null;
            ShowIcon = false;

            // Header
            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = Color.FromArgb(104, 33, 122)
            };
            Label headerLabel = new Label
            {
                Text = Loc.S("Dialog.Settings.Header"),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 12f, FontStyle.Bold),
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(12, 0, 0, 0)
            };
            header.Controls.Add(headerLabel);

            // Footer
            Panel footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                BackColor = Color.FromArgb(248, 248, 248)
            };
            footer.Paint += delegate(object s, PaintEventArgs pe)
            {
                using (Pen pen = new Pen(Color.FromArgb(220, 220, 220)))
                    pe.Graphics.DrawLine(pen, 0, 0, footer.Width, 0);
            };

            Button btnSave = new Button
            {
                Text = Loc.S("Settings.Save"),
                Size = new Size(80, 32),
                FlatStyle = FlatStyle.System
            };
            btnSave.Click += delegate { SaveAndClose(); };

            Button btnCancel = new Button
            {
                Text = Loc.S("Settings.Cancel"),
                Size = new Size(80, 32),
                FlatStyle = FlatStyle.System
            };
            btnCancel.Click += delegate { Close(); };

            btnCancel.Location = new Point(footer.Width - btnCancel.Width - 16, 9);
            btnSave.Location = new Point(btnCancel.Location.X - btnSave.Width - 8, 9);
            btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnSave.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            footer.Controls.Add(btnSave);
            footer.Controls.Add(btnCancel);

            // Content
            Panel content = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                AutoScroll = true,
                Padding = new Padding(24, 16, 24, 16)
            };

            int y = 16;
            Font labelFont = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold);
            Font headingFont = new Font("Microsoft YaHei UI", 11f, FontStyle.Bold);
            Font inputFont = new Font("Microsoft YaHei UI", 9.5f);
            Color purple = Color.FromArgb(104, 33, 122);

            // Section: Font
            Label secFont = new Label { Text = Loc.S("Settings.Section.Font"), Location = new Point(24, y), AutoSize = true, Font = headingFont, ForeColor = purple };
            content.Controls.Add(secFont);
            y += 32;

            _txtFontFamily = AddField(content, Loc.S("Settings.FontFamily"), ref y, labelFont, inputFont);
            _txtMonoFont = AddField(content, Loc.S("Settings.MonoFont"), ref y, labelFont, inputFont);
            _txtMathFont = AddField(content, Loc.S("Settings.MathFont"), ref y, labelFont, inputFont);
            _txtParagraphSize = AddField(content, Loc.S("Settings.ParagraphSize"), ref y, labelFont, inputFont);
            _txtCodeSize = AddField(content, Loc.S("Settings.CodeSize"), ref y, labelFont, inputFont);

            Label lblPreset = new Label
            {
                Text = Loc.S("Settings.ThemePreset"),
                Location = new Point(24, y),
                AutoSize = true,
                Font = labelFont,
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            content.Controls.Add(lblPreset);
            y += 22;
            _cboPreset = new ComboBox
            {
                Location = new Point(24, y),
                Size = new Size(240, 26),
                Font = inputFont,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cboPreset.Items.Add(Loc.S("Settings.ThemePreset.Technical"));
            _cboPreset.Items.Add(Loc.S("Settings.ThemePreset.Study"));
            _cboPreset.Items.Add(Loc.S("Settings.ThemePreset.Minimal"));
            _cboPreset.SelectionChangeCommitted += delegate { ApplySelectedPreset(); };
            content.Controls.Add(_cboPreset);
            y += 34;

            y += 8;
            Label secRender = new Label { Text = Loc.S("Settings.Section.Render"), Location = new Point(24, y), AutoSize = true, Font = headingFont, ForeColor = purple };
            content.Controls.Add(secRender);
            y += 32;

            _chkLatex = new CheckBox
            {
                Text = Loc.S("Settings.EnableLatex"),
                Location = new Point(24, y),
                AutoSize = true,
                Font = inputFont
            };
            content.Controls.Add(_chkLatex);
            y += 30;

            _chkLineNumber = new CheckBox
            {
                Text = Loc.S("Settings.EnableLineNumber"),
                Location = new Point(24, y),
                AutoSize = true,
                Font = inputFont
            };
            content.Controls.Add(_chkLineNumber);
            y += 36;

            _chkRemoteImages = new CheckBox
            {
                Text = Loc.S("Settings.AllowRemoteImages"),
                Location = new Point(24, y),
                AutoSize = true,
                Font = inputFont
            };
            content.Controls.Add(_chkRemoteImages);
            y += 36;

            Label secPreview = new Label { Text = Loc.S("Settings.Section.Preview"), Location = new Point(24, y), AutoSize = true, Font = headingFont, ForeColor = purple };
            content.Controls.Add(secPreview);
            y += 32;

            _chkPreviewTitle = new CheckBox
            {
                Text = Loc.S("Settings.PreviewShowTitle"),
                Location = new Point(24, y),
                AutoSize = true,
                Font = inputFont
            };
            content.Controls.Add(_chkPreviewTitle);
            y += 30;
            _txtPreviewTitle = AddField(content, Loc.S("Settings.PreviewTitle"), ref y, labelFont, inputFont);

            Label lblPosition = new Label
            {
                Text = Loc.S("Settings.PreviewPosition"),
                Location = new Point(24, y),
                AutoSize = true,
                Font = labelFont,
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            content.Controls.Add(lblPosition);
            y += 22;
            _cboPreviewPosition = new ComboBox
            {
                Location = new Point(24, y),
                Size = new Size(240, 26),
                Font = inputFont,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cboPreviewPosition.Items.Add(Loc.S("Settings.PreviewPosition.Right"));
            _cboPreviewPosition.Items.Add(Loc.S("Settings.PreviewPosition.Below"));
            content.Controls.Add(_cboPreviewPosition);
            y += 34;

            _txtPreviewGap = AddField(content, Loc.S("Settings.PreviewGap"), ref y, labelFont, inputFont);
            _txtPreviewWidth = AddField(content, Loc.S("Settings.PreviewWidth"), ref y, labelFont, inputFont);

            _chkAutoRefresh = new CheckBox
            {
                Text = Loc.S("Settings.AutoRefresh"),
                Location = new Point(24, y),
                AutoSize = true,
                Font = inputFont
            };
            content.Controls.Add(_chkAutoRefresh);
            y += 30;
            _txtAutoRefreshDelay = AddField(content, Loc.S("Settings.AutoRefreshDelay"), ref y, labelFont, inputFont);

            // Section: Language
            Label secLang = new Label { Text = Loc.S("Settings.Section.Language"), Location = new Point(24, y), AutoSize = true, Font = headingFont, ForeColor = purple };
            content.Controls.Add(secLang);
            y += 32;

            Label lblLang = new Label
            {
                Text = Loc.S("Settings.Language"),
                Location = new Point(24, y),
                AutoSize = true,
                Font = labelFont,
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            content.Controls.Add(lblLang);
            y += 22;

            _cboLanguage = new ComboBox
            {
                Location = new Point(24, y),
                Size = new Size(240, 26),
                Font = inputFont,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cboLanguage.Items.Add(Loc.S("Settings.Language.Auto"));
            _cboLanguage.Items.Add(Loc.S("Settings.Language.Zh"));
            _cboLanguage.Items.Add(Loc.S("Settings.Language.En"));
            content.Controls.Add(_cboLanguage);
            y += 34;

            y += 8;
            Label hint = new Label
            {
                Text = Loc.S("Settings.Hint"),
                Location = new Point(24, y),
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 8.5f),
                ForeColor = Color.Gray
            };
            content.Controls.Add(hint);

            // Assembly
            Controls.Add(content);
            Controls.Add(footer);
            Controls.Add(header);
            CancelButton = btnCancel;
            AcceptButton = btnSave;

            LoadSettings();
        }

        private TextBox AddField(Panel parent, string label, ref int y, Font labelFont, Font inputFont)
        {
            Label lbl = new Label
            {
                Text = label,
                Location = new Point(24, y),
                AutoSize = true,
                Font = labelFont,
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            parent.Controls.Add(lbl);
            y += 22;

            TextBox txt = new TextBox
            {
                Location = new Point(24, y),
                Size = new Size(300, 26),
                Font = inputFont,
                BorderStyle = BorderStyle.FixedSingle
            };
            parent.Controls.Add(txt);
            y += 34;

            return txt;
        }

        private void LoadSettings()
        {
            ThemeSettings s = ThemeSettings.Load();
            _txtFontFamily.Text = s.DefaultFontFamily;
            _txtMonoFont.Text = s.MonospaceFontFamily;
            _txtMathFont.Text = s.MathFontFamily;
            _txtParagraphSize.Text = s.ParagraphFontSize.ToString(CultureInfo.InvariantCulture);
            _txtCodeSize.Text = s.CodeFontSize.ToString(CultureInfo.InvariantCulture);
            _chkLatex.Checked = s.EnableLatexToImage;
            _chkLineNumber.Checked = s.EnableCodeLineNumber;
            _chkRemoteImages.Checked = s.AllowRemoteImages;
            _chkPreviewTitle.Checked = s.PreviewShowTitle;
            _txtPreviewTitle.Text = s.PreviewTitle;
            _cboPreviewPosition.SelectedIndex = s.PreviewPlacement == PreviewPlacement.Below ? 1 : 0;
            _txtPreviewGap.Text = s.PreviewGap.ToString(CultureInfo.InvariantCulture);
            _txtPreviewWidth.Text = s.PreviewWidth.ToString(CultureInfo.InvariantCulture);
            _chkAutoRefresh.Checked = s.AutoRefreshEnabled;
            _txtAutoRefreshDelay.Text = s.AutoRefreshDelayMilliseconds.ToString(CultureInfo.InvariantCulture);
            if (s.ThemePreset == "study") _cboPreset.SelectedIndex = 1;
            else if (s.ThemePreset == "minimal") _cboPreset.SelectedIndex = 2;
            else _cboPreset.SelectedIndex = 0;

            // Language dropdown: 0=auto, 1=zh, 2=en
            string lang = (s.Language ?? "auto").Trim().ToLowerInvariant();
            if (lang == "zh") _cboLanguage.SelectedIndex = 1;
            else if (lang == "en") _cboLanguage.SelectedIndex = 2;
            else _cboLanguage.SelectedIndex = 0;
        }

        private void SaveAndClose()
        {
            try
            {
                string langValue = "auto";
                if (_cboLanguage.SelectedIndex == 1) langValue = "zh";
                else if (_cboLanguage.SelectedIndex == 2) langValue = "en";
                string presetValue = _cboPreset.SelectedIndex == 1 ? "study"
                    : (_cboPreset.SelectedIndex == 2 ? "minimal" : "technical");
                string positionValue = _cboPreviewPosition.SelectedIndex == 1 ? "below" : "right";

                string content =
                    "# OneNote Markdown theme settings\r\n" +
                    "# Edit values and re-render page to apply\r\n" +
                    "font.family=" + _txtFontFamily.Text.Trim() + "\r\n" +
                    "font.monospace=" + _txtMonoFont.Text.Trim() + "\r\n" +
                    "font.math=" + _txtMathFont.Text.Trim() + "\r\n" +
                    "font.size.paragraph=" + _txtParagraphSize.Text.Trim() + "\r\n" +
                    "font.size.code=" + _txtCodeSize.Text.Trim() + "\r\n" +
                    "enable.latex.image=" + (_chkLatex.Checked ? "true" : "false") + "\r\n" +
                    "enable.code.lineNumber=" + (_chkLineNumber.Checked ? "true" : "false") + "\r\n" +
                    "theme.preset=" + presetValue + "\r\n" +
                    "preview.title.show=" + (_chkPreviewTitle.Checked ? "true" : "false") + "\r\n" +
                    "preview.title.text=" + _txtPreviewTitle.Text.Trim() + "\r\n" +
                    "preview.position=" + positionValue + "\r\n" +
                    "preview.gap=" + _txtPreviewGap.Text.Trim() + "\r\n" +
                    "preview.width=" + _txtPreviewWidth.Text.Trim() + "\r\n" +
                    "preview.autoRefresh=" + (_chkAutoRefresh.Checked ? "true" : "false") + "\r\n" +
                    "preview.autoRefresh.delayMs=" + _txtAutoRefreshDelay.Text.Trim() + "\r\n" +
                    "image.allowRemote=" + (_chkRemoteImages.Checked ? "true" : "false") + "\r\n" +
                    "diagram.timeoutMs=3000\r\n" +
                    "language=" + langValue + "\r\n";

                string path = ThemeSettings.EnsureDefaultFile();
                System.IO.File.WriteAllText(path, content, new System.Text.UTF8Encoding(false));

                // Apply language immediately
                LocalizationManager.SetLanguage(langValue);
                PageWriter.InvalidateThemeCache();

                Logger.Info("Settings saved");
                DialogResult = DialogResult.OK;
                Close();

                // Notify user about restart for full effect
                if (langValue != LocalizationManager.OverrideSetting)
                {
                    Msg.Show(Loc.S("Settings.RestartHint"), Loc.S("Common.AppTitle"),
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                Logger.Error("Save settings failed", ex);
                Msg.Show(Loc.S("Settings.SaveFailed", ex.Message), Loc.S("Common.AppTitle"),
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplySelectedPreset()
        {
            if (_cboPreset.SelectedIndex == 1)
            {
                _txtFontFamily.Text = "Calibri";
                _txtParagraphSize.Text = "12";
                _txtCodeSize.Text = "10";
            }
            else if (_cboPreset.SelectedIndex == 2)
            {
                _txtFontFamily.Text = "Segoe UI";
                _txtParagraphSize.Text = "10.5";
                _txtCodeSize.Text = "9.5";
            }
            else
            {
                _txtFontFamily.Text = "Calibri";
                _txtParagraphSize.Text = "11";
                _txtCodeSize.Text = "10";
            }
        }
    }
}
