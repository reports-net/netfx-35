using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Pao.Reports;

namespace Pao.Reports.Sample
{
	/// <summary>
	/// Form1 の概要の説明です。
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.RadioButton opt1;
		private System.Windows.Forms.RadioButton opt3;
		private System.Windows.Forms.RadioButton opt2;
		private System.Windows.Forms.RadioButton opt4;
		private System.Windows.Forms.RadioButton opt5;
		private System.Windows.Forms.Button button3;
        private TextBox txtMessage;
		/// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			//
			// Windows フォーム デザイナ サポートに必要です。
			//
			InitializeComponent();

			//
			// TODO: InitializeComponent 呼び出しの後に、コンストラクタ コードを追加してください。
			//
		}

		/// <summary>
		/// 使用されているリソースに後処理を実行します。
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// デザイナ サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディタで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.opt1 = new System.Windows.Forms.RadioButton();
            this.opt3 = new System.Windows.Forms.RadioButton();
            this.opt2 = new System.Windows.Forms.RadioButton();
            this.opt4 = new System.Windows.Forms.RadioButton();
            this.opt5 = new System.Windows.Forms.RadioButton();
            this.button3 = new System.Windows.Forms.Button();
            this.txtMessage = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("MS UI Gothic", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.Location = new System.Drawing.Point(194, 288);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(422, 144);
            this.button1.TabIndex = 0;
            this.button1.Text = "プレビュー";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("MS UI Gothic", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button2.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.button2.Location = new System.Drawing.Point(739, 288);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(434, 144);
            this.button2.TabIndex = 1;
            this.button2.Text = "印刷";
            this.button2.Click += new System.EventHandler(this.button1_Click);
            // 
            // opt1
            // 
            this.opt1.Checked = true;
            this.opt1.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.opt1.Location = new System.Drawing.Point(458, 64);
            this.opt1.Name = "opt1";
            this.opt1.Size = new System.Drawing.Size(246, 48);
            this.opt1.TabIndex = 2;
            this.opt1.TabStop = true;
            this.opt1.Text = "単純な印刷データ";
            // 
            // opt3
            // 
            this.opt3.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.opt3.Location = new System.Drawing.Point(1126, 64);
            this.opt3.Name = "opt3";
            this.opt3.Size = new System.Drawing.Size(346, 48);
            this.opt3.TabIndex = 3;
            this.opt3.Text = "住所一覧(MySQL 使用)";
            // 
            // opt2
            // 
            this.opt2.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.opt2.Location = new System.Drawing.Point(792, 64);
            this.opt2.Name = "opt2";
            this.opt2.Size = new System.Drawing.Size(299, 48);
            this.opt2.TabIndex = 4;
            this.opt2.Text = "10の倍数出力";
            // 
            // opt4
            // 
            this.opt4.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.opt4.Location = new System.Drawing.Point(440, 160);
            this.opt4.Name = "opt4";
            this.opt4.Size = new System.Drawing.Size(299, 48);
            this.opt4.TabIndex = 5;
            this.opt4.Text = "見積書(MySQL 使用)";
            // 
            // opt5
            // 
            this.opt5.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.opt5.Location = new System.Drawing.Point(792, 160);
            this.opt5.Name = "opt5";
            this.opt5.Size = new System.Drawing.Size(299, 48);
            this.opt5.TabIndex = 6;
            this.opt5.Text = "広告(MySQL 使用)";
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("MS UI Gothic", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button3.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.button3.Location = new System.Drawing.Point(1302, 288);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(434, 144);
            this.button3.TabIndex = 7;
            this.button3.Text = "PDF出力";
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(12, 12);
            this.txtMessage.Multiline = true;
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.Size = new System.Drawing.Size(1894, 518);
            this.txtMessage.TabIndex = 8;
            this.txtMessage.Text = resources.GetString("txtMessage.Text");
            this.txtMessage.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(11, 24);
            this.ClientSize = new System.Drawing.Size(1918, 583);
            this.Controls.Add(this.txtMessage);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.opt5);
            this.Controls.Add(this.opt4);
            this.Controls.Add(this.opt2);
            this.Controls.Add(this.opt3);
            this.Controls.Add(this.opt1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "WEBサーバ(www.pao.ac)の Axis WebService から Reports.jar で作成した印刷データをGETして印刷・プレビューを行うサンプ" +
                "ル";
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// アプリケーションのメイン エントリ ポイントです。
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
            txtMessage.Visible = true;
            return;

			byte[] data = null;

			ac.pao.www.SampleService webTest = new ac.pao.www.SampleService(); //WebService呼び出し

			if(opt1.Checked) data = webTest.getPrintData();		//単純な印刷データを取得
			if(opt2.Checked) data = webTest.getBaisuu();		//10の倍数サンプル 印刷データを取得
			if(opt3.Checked) data = webTest.getAddressList();	//住所一覧サンプル 印刷データを取得
			if(opt4.Checked) data = webTest.getMitsumori();		//見積書サンプル 印刷データを取得
			if(opt5.Checked) data = webTest.getKoukoku();		//広告サンプル 印刷データを取得

			System.Windows.Forms.Button b = (System.Windows.Forms.Button)sender;
			IReport paoRep;
			if(b.Text == "印刷")
			{
				paoRep = ReportCreator.GetReport(); // 印刷オブジェクト生成
			}
			else
			{
				paoRep = ReportCreator.GetPreview(); // プレビュー画面を作成
			}
			paoRep.LoadData(data);	//印刷データを読み込む
			paoRep.Output(); // 印刷又はプレビューを実行

		}

		private void button3_Click(object sender, System.EventArgs e)
		{
            txtMessage.Visible = true;
            return;

			if(opt5.Checked)
			{
				MessageBox.Show("広告サンプルのPDF作成WEBサービスは、制限により出力できません。");
				return;
			}
			byte[] data = null;

			pdf.ac.pao.www.SamplePdfService webTest = new pdf.ac.pao.www.SamplePdfService(); //WebService呼び出し

			if(opt1.Checked) data = webTest.getPrintData();		//単純な印刷データを取得
			if(opt2.Checked) data = webTest.getBaisuu();		//10の倍数サンプル 印刷データを取得
			if(opt3.Checked) data = webTest.getAddressList();	//住所一覧サンプル 印刷データを取得
			if(opt4.Checked) data = webTest.getMitsumori();		//見積書サンプル 印刷データを取得

			//PDF出力
			MessageBox.Show(this,"PDFファイルを「Sample.PDF」という名前でデスクトップに出力します。");
			//デスクトップのパス取得
			string DeskTop = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

			using (FileStream stream = new FileStream(Path.Combine(DeskTop, "Sample.PDF"), FileMode.Create, FileAccess.Write))
				   {
				stream.Write(data, 0, data.Length);
				stream.Close();
			}
			System.Diagnostics.Process.Start(DeskTop + "\\Sample.PDF");

		}

	}
}
