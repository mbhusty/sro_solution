namespace SRO
{
    partial class Connect
    {
        /// <summary>
        /// Требуется переменная конструктора.
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
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Connect));
            this.butSettings = new System.Windows.Forms.Button();
            this.butConnect = new System.Windows.Forms.Button();
            this.tbPhoneOper = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tbLogin = new System.Windows.Forms.TextBox();
            this.tbPassword = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // butSettings
            // 
            this.butSettings.Location = new System.Drawing.Point(11, 89);
            this.butSettings.Name = "butSettings";
            this.butSettings.Size = new System.Drawing.Size(72, 23);
            this.butSettings.TabIndex = 4;
            this.butSettings.Text = "Настройки";
            this.butSettings.UseVisualStyleBackColor = true;
            // 
            // butConnect
            // 
            this.butConnect.Location = new System.Drawing.Point(100, 89);
            this.butConnect.Name = "butConnect";
            this.butConnect.Size = new System.Drawing.Size(100, 23);
            this.butConnect.TabIndex = 3;
            this.butConnect.Text = "Подключение";
            this.butConnect.UseVisualStyleBackColor = true;
            this.butConnect.Click += new System.EventHandler(this.button2_Click);
            // 
            // tbPhoneOper
            // 
            this.tbPhoneOper.BackColor = System.Drawing.Color.Yellow;
            this.tbPhoneOper.Location = new System.Drawing.Point(100, 61);
            this.tbPhoneOper.Name = "tbPhoneOper";
            this.tbPhoneOper.Size = new System.Drawing.Size(100, 20);
            this.tbPhoneOper.TabIndex = 2;
            this.tbPhoneOper.TextChanged += new System.EventHandler(this.tbPhoneOper_TextChanged);
            this.tbPhoneOper.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tbPhoneOper_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Пользователь";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(45, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Пароль";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Добавочный";
            // 
            // tbLogin
            // 
            this.tbLogin.Location = new System.Drawing.Point(100, 9);
            this.tbLogin.Name = "tbLogin";
            this.tbLogin.Size = new System.Drawing.Size(100, 20);
            this.tbLogin.TabIndex = 0;
            this.tbLogin.Text = "admin";
            this.tbLogin.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tbLogin_KeyDown);
            // 
            // tbPassword
            // 
            this.tbPassword.Location = new System.Drawing.Point(100, 35);
            this.tbPassword.Name = "tbPassword";
            this.tbPassword.Size = new System.Drawing.Size(100, 20);
            this.tbPassword.TabIndex = 1;
            this.tbPassword.Text = "123456";
            this.tbPassword.UseSystemPasswordChar = true;
            this.tbPassword.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tbPassword_KeyDown);
            // 
            // Connect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(202, 115);
            this.Controls.Add(this.tbPassword);
            this.Controls.Add(this.tbLogin);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbPhoneOper);
            this.Controls.Add(this.butConnect);
            this.Controls.Add(this.butSettings);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(218, 154);
            this.MinimumSize = new System.Drawing.Size(218, 154);
            this.Name = "Connect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Подключение";
            this.Load += new System.EventHandler(this.Connect_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button butSettings;
        private System.Windows.Forms.Button butConnect;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbLogin;
        private System.Windows.Forms.TextBox tbPassword;
        public System.Windows.Forms.TextBox tbPhoneOper;
    }
}

