namespace ExcellentAddIn.UI
{
    partial class MainUserForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Очистка всех используемых ресурсов.
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

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // MainUserForm
            // 
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Name = "MainUserForm";
            this.Text = "Main User Form";
            this.Load += new System.EventHandler(this.MainUserForm_Load);
            this.ResumeLayout(false);

        }

        #endregion
    }
}
