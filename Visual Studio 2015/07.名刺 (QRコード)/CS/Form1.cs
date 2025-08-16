using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using Pao.Reports;

namespace Sample
{
	/// <summary>
	/// Form1 の概要の説明です。
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
        private System.Windows.Forms.RadioButton radPrint;
        private System.Windows.Forms.Button btnExe;
        private System.Windows.Forms.RadioButton radPreview;
		private System.Windows.Forms.DataGrid grid;
        /// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		private System.ComponentModel.Container components = null;
        #region コンストラクタ
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
        #endregion
        #region Dispose
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
        #endregion
		#region Windows Form Designer generated code
		/// <summary>
		/// デザイナ サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディタで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
            this.radPreview = new System.Windows.Forms.RadioButton();
            this.radPrint = new System.Windows.Forms.RadioButton();
            this.btnExe = new System.Windows.Forms.Button();
            this.grid = new System.Windows.Forms.DataGrid();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            this.SuspendLayout();
            // 
            // radPreview
            // 
            this.radPreview.Checked = true;
            this.radPreview.Font = new System.Drawing.Font("HG丸ｺﾞｼｯｸM-PRO", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPreview.Location = new System.Drawing.Point(232, 8);
            this.radPreview.Name = "radPreview";
            this.radPreview.Size = new System.Drawing.Size(142, 24);
            this.radPreview.TabIndex = 0;
            this.radPreview.TabStop = true;
            this.radPreview.Text = "プレビュー";
            // 
            // radPrint
            // 
            this.radPrint.Font = new System.Drawing.Font("HG丸ｺﾞｼｯｸM-PRO", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPrint.Location = new System.Drawing.Point(403, 8);
            this.radPrint.Name = "radPrint";
            this.radPrint.Size = new System.Drawing.Size(93, 24);
            this.radPrint.TabIndex = 1;
            this.radPrint.Text = "印刷";
            // 
            // btnExe
            // 
            this.btnExe.Font = new System.Drawing.Font("HGP創英角ﾎﾟｯﾌﾟ体", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnExe.Location = new System.Drawing.Point(16, 312);
            this.btnExe.Name = "btnExe";
            this.btnExe.Size = new System.Drawing.Size(720, 40);
            this.btnExe.TabIndex = 1;
            this.btnExe.Text = "実　　行";
            this.btnExe.Click += new System.EventHandler(this.btnExe_Click);
            // 
            // grid
            // 
            this.grid.DataMember = "";
            this.grid.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.grid.Location = new System.Drawing.Point(16, 48);
            this.grid.Name = "grid";
            this.grid.Size = new System.Drawing.Size(720, 248);
            this.grid.TabIndex = 3;
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 12);
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(746, 360);
            this.Controls.Add(this.grid);
            this.Controls.Add(this.btnExe);
            this.Controls.Add(this.radPrint);
            this.Controls.Add(this.radPreview);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Report.net サンプル (名刺作成)";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion
        #region エントリポイント
		/// <summary>
		/// アプリケーションのメイン エントリ ポイントです。
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}
        #endregion

        // プログラム実行フォルダ
        string appPath = null;
        // Excelデータベースファイル パス
        private string DbXls = "名刺.xls";
        // x64動作時加算パス(フォルダ)
        private string x64dir = "";

        OleDbDataAdapter dataAdapter;
		DataTable table;
		private void Form1_Load(object sender, System.EventArgs e)
		{
            //カレントPath取得
            appPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "/";
            if (!System.IO.File.Exists(appPath +  "../../../" + DbXls))
            {
                x64dir += "../";
                appPath += x64dir;
            }
            DbXls = appPath + "../../../" + DbXls;


            string connectString = null;
            if (IntPtr.Size == 4)
            {
                connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DbXls + ";Extended Properties=Excel 8.0;";
            }
            else if (IntPtr.Size == 8)
            {
                connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DbXls + ";Extended Properties=Excel 12.0;";
            }

            OleDbConnection connection = new OleDbConnection(connectString);

			string SQL = "select * from [名刺$]";
			
			dataAdapter = new OleDbDataAdapter(SQL, connection);
			DataSet ds = new DataSet();
			
            try
            {
                dataAdapter.Fill(ds, "名刺");
            }
            catch
            {
                if (MessageBox.Show("このサンプルプログラムを動作させるためには、データベースへアクセスのため"
                        + Environment.NewLine + "[Microsoft Access データベース エンジン 2010 再頒布可能コンポーネント]"
                        + Environment.NewLine + "をインストールする必要があります。"
                        + Environment.NewLine + "マイクロソフトのインストーラ ダウンロードサイトへジャンプしますか？"
                        , "サンプルが動作しない時", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    == DialogResult.Yes)
                {
                    ExecFile("http://www.microsoft.com/ja-jp/download/details.aspx?id=13255");
                }

                return;
            }
            table = ds.Tables["名刺"];

			grid.DataSource = table;

		}
		
		private void btnExe_Click(object sender, System.EventArgs e)
        {

            
            //選択されている行の取得
            //行数の取得
            int n =
				grid.BindingContext[grid.DataSource,
				grid.DataMember].Count;
			int rowNo;
			for (rowNo = 0; rowNo < n; rowNo++)
			{
				//行が選択されているか調べる
				if (grid.IsSelected(rowNo))
				{
					break;
				}           
			}

			if(rowNo >= n)
			{
				MessageBox.Show("行が選択されていません");
				return;
			}

            IReport paoRep = null;
            
            if (radPreview.Checked){ // プレビューを選択している場合
				//プレビューオブジェクトのインスタンスを獲得
				paoRep = ReportCreator.GetPreview();
			}
			else if(radPrint.Checked)
			{
				//印刷オブジェクトのインスタンスを獲得
				paoRep = ReportCreator.GetReport();
			}

            paoRep.LoadDefFile(appPath + "../../../名刺.prepd");
			DataRow row = table.Rows[rowNo];

			string name1 = row["名前"].ToString();
			string kata = row["肩書き"].ToString();
			string mail = row["メール"].ToString();
			string tel = row["携帯"].ToString();
			string name2 = row["携帯名前"].ToString();
			string kana = row["携帯ｶﾅ"].ToString();

			paoRep.PageStart();

			for(int line=1; line<=5; line++)
			{
                for (int col = 1; col <= 2; col++)
                {

                    //Body Print
                    paoRep.Write("名前", name1, col, line);
                    paoRep.Write("肩書き", kata, col, line);
                    paoRep.Write("メール", mail, col, line);
                    paoRep.Write("携帯", tel, col, line);
                    paoRep.Write("QR", "\"MECARD:N:" + name2 + ";SOUND:" + kana + ";TEL:" + tel + ";EMAIL:" + mail + ";;\"", col, line);
                    paoRep.Write("a", "Pao@Office", col, line);
                    paoRep.Write("b", "有限会社", col, line);
                    paoRep.Write("c", "パオ･アット･オフィス", col, line);
                    paoRep.Write("d", "mail:", col, line);
                    paoRep.Write("e", "携帯", col, line);
                    paoRep.Write("f", "http://www.pao.ac/", col, line);
                    paoRep.Write("g", "本　　　社　〒275-0026　千葉県習志野市谷津3-29-2-401\n　　　　　　TEL:047-452-0057　FAX:047-452-0064", col, line);
                    paoRep.Write("h", "東京事務所　〒105-0004　東京都港区新橋1-8-3 住友新橋ビル7F\n　　　　　　TEL:03-3572-6507　FAX:03-6218-0128", col, line);
                }

            }
            paoRep.PageEnd();

			//印刷/プレビュー
			paoRep.Output();

			dataAdapter.Dispose();
        }

        private void ExecFile(string ExecFilePath)
        {
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo(ExecFilePath);
            startInfo.UseShellExecute = true;
            System.Diagnostics.Process.Start(startInfo);
        }

    }
}
