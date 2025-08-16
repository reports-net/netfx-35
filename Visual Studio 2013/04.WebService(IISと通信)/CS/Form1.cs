using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Pao.Reports;

namespace Pao.Reports.Sample1
{
	/// <summary>
	/// Form1 の概要の説明です。
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
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
			this.button1 = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(24, 24);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(240, 136);
			this.button1.TabIndex = 0;
			this.button1.Text = "　　WEBサーバ(iis.pao.ac)のWebServiceから　　印刷データをGETしてプレビューを行う";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 12);
			this.ClientSize = new System.Drawing.Size(292, 189);
			this.Controls.Add(this.button1);
			this.Name = "Form1";
			this.Text = "WebService(IISと通信)";
			this.ResumeLayout(false);

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

            Sample.ac.pao.iis.WebTest webTest = new Sample.ac.pao.iis.WebTest(); //WebService呼び出し
			byte[] data = webTest.get帳票データ(); //印刷データを取得

			IReport paoRep = ReportCreator.GetPreview(); // プレビュー画面を作成
			paoRep.LoadData(data);	//印刷データを読み込む
			paoRep.Output(); // 印刷又はプレビューを実行

		}
	}
}
