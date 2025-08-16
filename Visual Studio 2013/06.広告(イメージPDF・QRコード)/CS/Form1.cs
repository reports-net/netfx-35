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
		private System.Windows.Forms.RadioButton radPDF;
		private System.Windows.Forms.SaveFileDialog saveFileDialog;
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
            this.radPDF = new System.Windows.Forms.RadioButton();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // radPreview
            // 
            this.radPreview.Checked = true;
            this.radPreview.Font = new System.Drawing.Font("MS UI Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPreview.Location = new System.Drawing.Point(48, 32);
            this.radPreview.Name = "radPreview";
            this.radPreview.Size = new System.Drawing.Size(104, 24);
            this.radPreview.TabIndex = 0;
            this.radPreview.TabStop = true;
            this.radPreview.Text = "プレビュー";
            // 
            // radPrint
            // 
            this.radPrint.Font = new System.Drawing.Font("MS UI Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPrint.Location = new System.Drawing.Point(168, 32);
            this.radPrint.Name = "radPrint";
            this.radPrint.Size = new System.Drawing.Size(104, 24);
            this.radPrint.TabIndex = 1;
            this.radPrint.Text = "印刷";
            // 
            // btnExe
            // 
            this.btnExe.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnExe.Location = new System.Drawing.Point(34, 75);
            this.btnExe.Name = "btnExe";
            this.btnExe.Size = new System.Drawing.Size(348, 49);
            this.btnExe.TabIndex = 1;
            this.btnExe.Text = "実行";
            this.btnExe.Click += new System.EventHandler(this.btnExe_Click);
            // 
            // radPDF
            // 
            this.radPDF.Font = new System.Drawing.Font("MS UI Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPDF.Location = new System.Drawing.Point(280, 32);
            this.radPDF.Name = "radPDF";
            this.radPDF.Size = new System.Drawing.Size(139, 24);
            this.radPDF.TabIndex = 3;
            this.radPDF.Text = "イメージPDF";
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 12);
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(252)))), ((int)(((byte)(238)))), ((int)(((byte)(235)))));
            this.ClientSize = new System.Drawing.Size(418, 148);
            this.Controls.Add(this.radPDF);
            this.Controls.Add(this.btnExe);
            this.Controls.Add(this.radPrint);
            this.Controls.Add(this.radPreview);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Report.net サンプル (広告)";
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
        private void btnExe_Click(object sender, System.EventArgs e)
        {
            
            //カレントPath取得
            string appPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "/";
            string x64dir = "";
            string DbXls = "広告.xls";

            if (!System.IO.File.Exists(appPath +  "../../../" + DbXls))
            {
                x64dir += "../";
                appPath += x64dir;
            }
            DbXls = appPath + "../../../" + DbXls;



            // データ取得
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

            string SQL = "select * from [広告情報$]";
			
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(SQL, connection);
            DataSet ds = new DataSet();

            try
            {
                dataAdapter.Fill(ds, "広告情報");
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

            DataTable table = ds.Tables["広告情報"];

            IReport paoRep = null;
            
            if (radPreview.Checked){	// プレビューを選択している場合
                paoRep = ReportCreator.GetPreview();
            }else if(radPrint.Checked){ //印刷の場合
                paoRep = ReportCreator.GetReport();
            }else{						//イメージPDF出力の場合
				paoRep = ReportCreator.GetImagePdf();
			}
            
            paoRep.LoadDefFile(appPath + "../../../広告.prepd");
            foreach (DataRow row in table.Rows){

				paoRep.PageStart();

				paoRep.Write("製品名",(string)row["製品名"]);
				paoRep.Write("キャッチフレーズ", (string)row["キャッチフレーズ"]);
				paoRep.Write("商品コード", (string)row["商品コード"]);
				paoRep.Write("JANコード", (string)row["商品コード"]);
				paoRep.Write("売り文句", (string)row["売り文句"]);
				paoRep.Write("説明", (string)row["説明"]);
				paoRep.Write("価格", (string)row["価格"]);
				paoRep.Write("画像1", appPath + "../../../" + (string)row["画像1"]);
				paoRep.Write("画像2", appPath + "../../../" + (string)row["画像2"]);
				paoRep.Write("QR",(string)row["製品名"] + " " + (string)row["キャッチフレーズ"]);

				paoRep.PageEnd();

            }

			if(!radPDF.Checked) //印刷・プレビューが選択されている場合
			{
				//印刷/プレビュー
				paoRep.Output();
			}
			else //PDF出力が選択されている場合
			{
				//ファイルの保存ダイアログの処理
				saveFileDialog.FileName = "広告";
				saveFileDialog.Filter = "PDF形式 (*.pdf)|*.pdf";

				if (saveFileDialog.ShowDialog() == DialogResult.OK)
				{
					//PDF出力
					paoRep.SavePDF(saveFileDialog.FileName);

					if(MessageBox.Show(this,"PDFを表示しますか？", "PDF の表示", MessageBoxButtons.YesNo ) == DialogResult.Yes)
					{
                        ExecFile(saveFileDialog.FileName);
                    }
                }
			}
		}

        private void ExecFile(string ExecFilePath)
        {
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo(ExecFilePath);
            startInfo.UseShellExecute = true;
            System.Diagnostics.Process.Start(startInfo);
        }

    }
}
