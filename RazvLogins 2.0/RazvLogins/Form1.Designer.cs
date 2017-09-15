namespace RazvLogins
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.BTest = new System.Windows.Forms.Button();
            this.BEnd = new System.Windows.Forms.Button();
            this.TTestik = new System.Windows.Forms.TextBox();
            this.SmartMultiPage = new System.Windows.Forms.TabControl();
            this.TCBrusakova = new System.Windows.Forms.TabPage();
            this.TCLA = new System.Windows.Forms.TabPage();
            this.TCLK = new System.Windows.Forms.TabPage();
            this.TCRizhkova = new System.Windows.Forms.TabPage();
            this.TCEmbah = new System.Windows.Forms.TabPage();
            this.SmartMultiPage.SuspendLayout();
            this.SuspendLayout();
            // 
            // BTest
            // 
            this.BTest.Font = new System.Drawing.Font("Comic Sans MS", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.BTest.Location = new System.Drawing.Point(708, 517);
            this.BTest.Name = "BTest";
            this.BTest.Size = new System.Drawing.Size(210, 50);
            this.BTest.TabIndex = 0;
            this.BTest.Text = "Логин Пробник";
            this.BTest.UseVisualStyleBackColor = true;
            this.BTest.Visible = false;
            this.BTest.Click += new System.EventHandler(this.BTest_Click);
            // 
            // BEnd
            // 
            this.BEnd.BackColor = System.Drawing.Color.OrangeRed;
            this.BEnd.Font = new System.Drawing.Font("Comic Sans MS", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.BEnd.Location = new System.Drawing.Point(386, 517);
            this.BEnd.Name = "BEnd";
            this.BEnd.Size = new System.Drawing.Size(210, 50);
            this.BEnd.TabIndex = 1;
            this.BEnd.Text = "ЗАКРЫТЬ";
            this.BEnd.UseVisualStyleBackColor = false;
            this.BEnd.Click += new System.EventHandler(this.BEnd_Click);
            // 
            // TTestik
            // 
            this.TTestik.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TTestik.Location = new System.Drawing.Point(21, 504);
            this.TTestik.Name = "TTestik";
            this.TTestik.Size = new System.Drawing.Size(332, 22);
            this.TTestik.TabIndex = 3;
            // 
            // SmartMultiPage
            // 
            this.SmartMultiPage.Controls.Add(this.TCBrusakova);
            this.SmartMultiPage.Controls.Add(this.TCLA);
            this.SmartMultiPage.Controls.Add(this.TCLK);
            this.SmartMultiPage.Controls.Add(this.TCRizhkova);
            this.SmartMultiPage.Controls.Add(this.TCEmbah);
            this.SmartMultiPage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.SmartMultiPage.Location = new System.Drawing.Point(21, 14);
            this.SmartMultiPage.Name = "SmartMultiPage";
            this.SmartMultiPage.SelectedIndex = 0;
            this.SmartMultiPage.Size = new System.Drawing.Size(1267, 471);
            this.SmartMultiPage.TabIndex = 4;
            // 
            // TCBrusakova
            // 
            this.TCBrusakova.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TCBrusakova.Location = new System.Drawing.Point(4, 25);
            this.TCBrusakova.Name = "TCBrusakova";
            this.TCBrusakova.Padding = new System.Windows.Forms.Padding(3);
            this.TCBrusakova.Size = new System.Drawing.Size(1259, 442);
            this.TCBrusakova.TabIndex = 1;
            this.TCBrusakova.Text = "Брусакова Наташа";
            this.TCBrusakova.UseVisualStyleBackColor = true;
            // 
            // TCLA
            // 
            this.TCLA.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TCLA.Location = new System.Drawing.Point(4, 25);
            this.TCLA.Name = "TCLA";
            this.TCLA.Padding = new System.Windows.Forms.Padding(3);
            this.TCLA.Size = new System.Drawing.Size(1259, 442);
            this.TCLA.TabIndex = 0;
            this.TCLA.Text = "Группа Л-А";
            this.TCLA.UseVisualStyleBackColor = true;
            // 
            // TCLK
            // 
            this.TCLK.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TCLK.Location = new System.Drawing.Point(4, 25);
            this.TCLK.Name = "TCLK";
            this.TCLK.Size = new System.Drawing.Size(1259, 442);
            this.TCLK.TabIndex = 4;
            this.TCLK.Text = "Группа Л-К";
            this.TCLK.UseVisualStyleBackColor = true;
            // 
            // TCRizhkova
            // 
            this.TCRizhkova.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TCRizhkova.Location = new System.Drawing.Point(4, 25);
            this.TCRizhkova.Name = "TCRizhkova";
            this.TCRizhkova.Size = new System.Drawing.Size(1259, 442);
            this.TCRizhkova.TabIndex = 3;
            this.TCRizhkova.Text = "Рыжкова Маша";
            this.TCRizhkova.UseVisualStyleBackColor = true;
            // 
            // TCEmbah
            // 
            this.TCEmbah.BackColor = System.Drawing.SystemColors.Control;
            this.TCEmbah.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TCEmbah.Location = new System.Drawing.Point(4, 25);
            this.TCEmbah.Name = "TCEmbah";
            this.TCEmbah.Size = new System.Drawing.Size(1259, 442);
            this.TCEmbah.TabIndex = 2;
            this.TCEmbah.Text = "Эмбах Саша";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(1308, 588);
            this.Controls.Add(this.SmartMultiPage);
            this.Controls.Add(this.TTestik);
            this.Controls.Add(this.BEnd);
            this.Controls.Add(this.BTest);
            this.Name = "Form1";
            this.Text = "Окошечко";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.SmartMultiPage.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BTest;
        private System.Windows.Forms.Button BEnd;
        private System.Windows.Forms.TextBox TTestik;
        private System.Windows.Forms.TabControl SmartMultiPage;
        private System.Windows.Forms.TabPage TCBrusakova;
        private System.Windows.Forms.TabPage TCLA;
        private System.Windows.Forms.TabPage TCLK;
        private System.Windows.Forms.TabPage TCRizhkova;
        private System.Windows.Forms.TabPage TCEmbah;
    }
}

