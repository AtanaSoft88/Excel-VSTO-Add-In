
namespace excelAddInTest
{
    partial class SidePanel
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.loadBtn = new Guna.UI2.WinForms.Guna2Button();
            this.calculateBtn = new Guna.UI2.WinForms.Guna2Button();
            this.clearBtn = new Guna.UI2.WinForms.Guna2Button();
            this.guna2Separator1 = new Guna.UI2.WinForms.Guna2Separator();
            this.guna2Separator2 = new Guna.UI2.WinForms.Guna2Separator();
            this.guna2Elipse1 = new Guna.UI2.WinForms.Guna2Elipse(this.components);
            this.loadFromDbLabel = new Guna.UI2.WinForms.Guna2HtmlLabel();
            this.calculateCommissionLabel = new Guna.UI2.WinForms.Guna2HtmlLabel();
            this.clearExcelLabel = new Guna.UI2.WinForms.Guna2HtmlLabel();
            this.SuspendLayout();
            // 
            // loadBtn
            // 
            this.loadBtn.BorderRadius = 15;
            this.loadBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.loadBtn.DisabledState.BorderColor = System.Drawing.Color.DarkGray;
            this.loadBtn.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray;
            this.loadBtn.DisabledState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(169)))), ((int)(((byte)(169)))), ((int)(((byte)(169)))));
            this.loadBtn.DisabledState.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(141)))), ((int)(((byte)(141)))), ((int)(((byte)(141)))));
            this.loadBtn.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.loadBtn.FocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.loadBtn.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.loadBtn.ForeColor = System.Drawing.Color.GreenYellow;
            this.loadBtn.Location = new System.Drawing.Point(14, 35);
            this.loadBtn.Name = "loadBtn";
            this.loadBtn.Size = new System.Drawing.Size(121, 48);
            this.loadBtn.TabIndex = 0;
            this.loadBtn.Text = "Load";
            this.loadBtn.Click += new System.EventHandler(this.loadBtn_Click);
            // 
            // calculateBtn
            // 
            this.calculateBtn.BorderRadius = 15;
            this.calculateBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.calculateBtn.DisabledState.BorderColor = System.Drawing.Color.DarkGray;
            this.calculateBtn.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray;
            this.calculateBtn.DisabledState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(169)))), ((int)(((byte)(169)))), ((int)(((byte)(169)))));
            this.calculateBtn.DisabledState.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(141)))), ((int)(((byte)(141)))), ((int)(((byte)(141)))));
            this.calculateBtn.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.calculateBtn.FocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.calculateBtn.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.calculateBtn.ForeColor = System.Drawing.Color.GreenYellow;
            this.calculateBtn.Location = new System.Drawing.Point(14, 147);
            this.calculateBtn.Name = "calculateBtn";
            this.calculateBtn.Size = new System.Drawing.Size(121, 44);
            this.calculateBtn.TabIndex = 1;
            this.calculateBtn.Text = "Calculate";
            this.calculateBtn.Click += new System.EventHandler(this.calculateBtn_Click);
            // 
            // clearBtn
            // 
            this.clearBtn.BorderRadius = 15;
            this.clearBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.clearBtn.DisabledState.BorderColor = System.Drawing.Color.DarkGray;
            this.clearBtn.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray;
            this.clearBtn.DisabledState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(169)))), ((int)(((byte)(169)))), ((int)(((byte)(169)))));
            this.clearBtn.DisabledState.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(141)))), ((int)(((byte)(141)))), ((int)(((byte)(141)))));
            this.clearBtn.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.clearBtn.FocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.clearBtn.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.clearBtn.ForeColor = System.Drawing.Color.GreenYellow;
            this.clearBtn.Location = new System.Drawing.Point(14, 258);
            this.clearBtn.Name = "clearBtn";
            this.clearBtn.Size = new System.Drawing.Size(121, 45);
            this.clearBtn.TabIndex = 2;
            this.clearBtn.Text = "Clear";
            this.clearBtn.Click += new System.EventHandler(this.clearBtn_Click);
            // 
            // guna2Separator1
            // 
            this.guna2Separator1.Location = new System.Drawing.Point(0, 104);
            this.guna2Separator1.Name = "guna2Separator1";
            this.guna2Separator1.Size = new System.Drawing.Size(297, 10);
            this.guna2Separator1.TabIndex = 3;
            // 
            // guna2Separator2
            // 
            this.guna2Separator2.Location = new System.Drawing.Point(3, 213);
            this.guna2Separator2.Name = "guna2Separator2";
            this.guna2Separator2.Size = new System.Drawing.Size(297, 10);
            this.guna2Separator2.TabIndex = 4;
            // 
            // guna2Elipse1
            // 
            this.guna2Elipse1.BorderRadius = 28;
            this.guna2Elipse1.TargetControl = this;
            // 
            // loadFromDbLabel
            // 
            this.loadFromDbLabel.BackColor = System.Drawing.Color.Transparent;
            this.loadFromDbLabel.Location = new System.Drawing.Point(19, 5);
            this.loadFromDbLabel.Name = "loadFromDbLabel";
            this.loadFromDbLabel.Size = new System.Drawing.Size(102, 22);
            this.loadFromDbLabel.TabIndex = 5;
            this.loadFromDbLabel.Text = "Load from DB";
            // 
            // calculateCommissionLabel
            // 
            this.calculateCommissionLabel.BackColor = System.Drawing.Color.Transparent;
            this.calculateCommissionLabel.Location = new System.Drawing.Point(19, 114);
            this.calculateCommissionLabel.Name = "calculateCommissionLabel";
            this.calculateCommissionLabel.Size = new System.Drawing.Size(159, 22);
            this.calculateCommissionLabel.TabIndex = 6;
            this.calculateCommissionLabel.Text = "Calculate Commission";
            // 
            // clearExcelLabel
            // 
            this.clearExcelLabel.BackColor = System.Drawing.Color.Transparent;
            this.clearExcelLabel.Location = new System.Drawing.Point(19, 225);
            this.clearExcelLabel.Name = "clearExcelLabel";
            this.clearExcelLabel.Size = new System.Drawing.Size(120, 22);
            this.clearExcelLabel.TabIndex = 7;
            this.clearExcelLabel.Text = "Clear Excel Cells";
            // 
            // SidePanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.CornflowerBlue;
            this.Controls.Add(this.clearExcelLabel);
            this.Controls.Add(this.calculateCommissionLabel);
            this.Controls.Add(this.loadFromDbLabel);
            this.Controls.Add(this.guna2Separator2);
            this.Controls.Add(this.guna2Separator1);
            this.Controls.Add(this.clearBtn);
            this.Controls.Add(this.calculateBtn);
            this.Controls.Add(this.loadBtn);
            this.Name = "SidePanel";
            this.Size = new System.Drawing.Size(300, 397);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Guna.UI2.WinForms.Guna2Button loadBtn;
        private Guna.UI2.WinForms.Guna2Button calculateBtn;
        private Guna.UI2.WinForms.Guna2Button clearBtn;
        private Guna.UI2.WinForms.Guna2Separator guna2Separator1;
        private Guna.UI2.WinForms.Guna2Separator guna2Separator2;
        private Guna.UI2.WinForms.Guna2Elipse guna2Elipse1;
        private Guna.UI2.WinForms.Guna2HtmlLabel calculateCommissionLabel;
        private Guna.UI2.WinForms.Guna2HtmlLabel loadFromDbLabel;
        private Guna.UI2.WinForms.Guna2HtmlLabel clearExcelLabel;
    }
}
