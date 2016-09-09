namespace MotoBRP
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
            this.btnActualPrice = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnActualPrice
            // 
            this.btnActualPrice.Location = new System.Drawing.Point(12, 12);
            this.btnActualPrice.Name = "btnActualPrice";
            this.btnActualPrice.Size = new System.Drawing.Size(143, 23);
            this.btnActualPrice.TabIndex = 0;
            this.btnActualPrice.Text = "Обработать прайс";
            this.btnActualPrice.UseVisualStyleBackColor = true;
            this.btnActualPrice.Click += new System.EventHandler(this.btnActualPrice_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1015, 434);
            this.Controls.Add(this.btnActualPrice);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnActualPrice;
    }
}

