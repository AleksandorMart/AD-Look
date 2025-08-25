using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ADlook
{
    public partial class ProgressForm : Form
    {
        private readonly ProgressBar _progressBar;
        private readonly Label _statusLabel;
        private readonly Button _cancelButton;
        private bool _cancelled;

        public bool Cancelled => _cancelled;

        public ProgressForm(string title = "Экспорт данных", string initialMessage = "Идет обработка...")
        {
            InitializeComponent();

            // Настройка формы
            this.Text = title;
            this.Size = new Size(400, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;

            // Создание элементов
            _statusLabel = new Label
            {
                Text = initialMessage,
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleCenter,
                Height = 40
            };

            _progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100
            };

            _cancelButton = new Button
            {
                Text = "Отмена",
                Dock = DockStyle.Bottom,
                Height = 30
            };
            _cancelButton.Click += (s, e) =>
            {
                _cancelled = true;
                _cancelButton.Enabled = false;
                _statusLabel.Text = "Отмена операции...";
            };

            // Добавляем элементы на форму
            this.Controls.Add(_progressBar);
            this.Controls.Add(_statusLabel);
            this.Controls.Add(_cancelButton);
        }

        // Обновляем статус и прогресс
        public void UpdateProgress(int current, int total, string message = null)
        {
            if (!string.IsNullOrEmpty(message))
            {
                _statusLabel.Text = message;
            }

            if (total > 0)
            {
                int percent = (int)((double)current / total * 100);
                _progressBar.Value = Math.Min(Math.Max(percent, 0), 100);

                if (string.IsNullOrEmpty(message))
                    _statusLabel.Text = $"Обработано {current} из {total} записей ({percent}%)";
            }

            Application.DoEvents();
        }

        // Обновляем только сообщение
        public void UpdateMessage(string message)
        {
            _statusLabel.Text = message;
            Application.DoEvents();
        }
    }
}
