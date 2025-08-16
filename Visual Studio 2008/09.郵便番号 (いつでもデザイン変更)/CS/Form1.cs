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
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
		private System.Windows.Forms.RadioButton radPDF;
        private System.Windows.Forms.RadioButton radPrint;
		private System.Windows.Forms.Button btnExe;
        private GroupBox groupBox1;
        private Button btnPreview;
        private RadioButton radD2;
        private RadioButton radD1;
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
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.radPDF = new System.Windows.Forms.RadioButton();
            this.radPrint = new System.Windows.Forms.RadioButton();
            this.btnExe = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radD2 = new System.Windows.Forms.RadioButton();
            this.radD1 = new System.Windows.Forms.RadioButton();
            this.btnPreview = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // radPDF
            // 
            this.radPDF.Checked = true;
            this.radPDF.Font = new System.Drawing.Font("Meiryo UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPDF.Location = new System.Drawing.Point(433, 311);
            this.radPDF.Name = "radPDF";
            this.radPDF.Size = new System.Drawing.Size(96, 32);
            this.radPDF.TabIndex = 10;
            this.radPDF.TabStop = true;
            this.radPDF.Text = "PDF出力";
            // 
            // radPrint
            // 
            this.radPrint.Font = new System.Drawing.Font("Meiryo UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPrint.Location = new System.Drawing.Point(563, 311);
            this.radPrint.Name = "radPrint";
            this.radPrint.Size = new System.Drawing.Size(75, 32);
            this.radPrint.TabIndex = 9;
            this.radPrint.Text = "印刷";
            // 
            // btnExe
            // 
            this.btnExe.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnExe.FlatAppearance.BorderSize = 2;
            this.btnExe.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExe.Font = new System.Drawing.Font("Meiryo UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnExe.ForeColor = System.Drawing.Color.Black;
            this.btnExe.Location = new System.Drawing.Point(425, 344);
            this.btnExe.Name = "btnExe";
            this.btnExe.Size = new System.Drawing.Size(213, 36);
            this.btnExe.TabIndex = 7;
            this.btnExe.Text = "実行";
            this.btnExe.Click += new System.EventHandler(this.btnExe_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radD2);
            this.groupBox1.Controls.Add(this.radD1);
            this.groupBox1.Controls.Add(this.btnPreview);
            this.groupBox1.Font = new System.Drawing.Font("Meiryo UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.groupBox1.Location = new System.Drawing.Point(81, 89);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(492, 174);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "デザイン選択＆プレビュー (帳票データを再読み込みしません)";
            // 
            // radD2
            // 
            this.radD2.AutoSize = true;
            this.radD2.Location = new System.Drawing.Point(278, 50);
            this.radD2.Name = "radD2";
            this.radD2.Size = new System.Drawing.Size(93, 24);
            this.radD2.TabIndex = 2;
            this.radD2.Text = "デザイン２";
            this.radD2.UseVisualStyleBackColor = true;
            // 
            // radD1
            // 
            this.radD1.AutoSize = true;
            this.radD1.Checked = true;
            this.radD1.Location = new System.Drawing.Point(122, 50);
            this.radD1.Name = "radD1";
            this.radD1.Size = new System.Drawing.Size(93, 24);
            this.radD1.TabIndex = 1;
            this.radD1.TabStop = true;
            this.radD1.Text = "デザイン１";
            this.radD1.UseVisualStyleBackColor = true;
            // 
            // btnPreview
            // 
            this.btnPreview.BackColor = System.Drawing.Color.DarkBlue;
            this.btnPreview.FlatAppearance.BorderSize = 5;
            this.btnPreview.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPreview.Font = new System.Drawing.Font("Meiryo UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnPreview.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.btnPreview.Location = new System.Drawing.Point(89, 91);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(311, 57);
            this.btnPreview.TabIndex = 0;
            this.btnPreview.Text = "プレビューしてデザインを確認";
            this.btnPreview.UseVisualStyleBackColor = false;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(7, 16);
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(218)))), ((int)(((byte)(222)))));
            this.ClientSize = new System.Drawing.Size(650, 386);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.radPDF);
            this.Controls.Add(this.radPrint);
            this.Controls.Add(this.btnExe);
            this.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Report.NET使用例－デザイン選択";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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

        bool _loadDesign = false;

        private void btnExe_Click(object sender, System.EventArgs e)
        {
            Output(false);
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            Output(true);
        }


        IReport paoRep = null;
        private void Output(bool flgPreview)
        {
            //カレントPath取得
            string appPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "/";
            string x64dir = "";
            string DbXls = "zip.xls";

            if (!System.IO.File.Exists(appPath +  "../../../" + DbXls))
            {
                x64dir += "../";
                appPath += x64dir;
            }
            DbXls = appPath + "../../../" + DbXls;

            if (flgPreview)
            {
                if (paoRep == null)
                {
                    //プレビューオブジェクトのインスタンスを獲得
                    paoRep = ReportCreator.GetPreview();
                }
            }
            else if (radPrint.Checked)
            {
                //印刷オブジェクトのインスタンスを獲得
                paoRep = ReportCreator.GetReport();
            }
            else if (radPDF.Checked)
            {
                //PDF出力オブジェクトのインスタンスを獲得
                paoRep = ReportCreator.GetPdf();
            }


            int page = 0;
            int line = 999;

            string[] defFile = { appPath + "../../../PaoRep1.prepd", appPath + "../../../PaoRep2.prepd" };
            int defIndex = 0;
            if (radD2.Checked) defIndex = 1;

            string hDate = System.DateTime.Now.ToString();

            if (!_loadDesign)
            {
                paoRep.LoadDefFile(defFile[defIndex]);
            }
            else
            {
                paoRep.ChangeDefFile(defFile[defIndex]);
            }

            if (!_loadDesign)
            {
                //paoRep.ClearData();
            //}

                _loadDesign = true;

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

                string SQL = "select * from [郵便番号テーブル$]";

                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(SQL, connection);
                DataSet ds = new DataSet();
                try
                {
                    dataAdapter.Fill(ds, "PostTable");
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

                DataTable table = new DataTable();
                table = ds.Tables["PostTable"];

                foreach (DataRow row in table.Rows)
                {
                    line++;
                    if (line > 32)
                    { // Head Print
                        if (page != 0) paoRep.PageEnd();

                        page++;

                        paoRep.PageStart();

                        paoRep.Write("日時", hDate);
                        paoRep.Write("ページ", "Page-" + page.ToString());

                        line = 1;

                    }

                    //Body Print
                    paoRep.Write("郵便番号", row["郵便番号"].ToString(), line);
                    paoRep.Write("市区町村", row["市区町村"].ToString(), line);
                    paoRep.Write("住所", row["住所"].ToString(), line);
                    paoRep.Write("横罫線", line);



                }
                paoRep.PageEnd();

                dataAdapter.Dispose();

            }

            if (flgPreview) //プレビューが選択されている場合
            {
                // このサンプルでは1ページのみ出力
                System.Drawing.Printing.PrinterSettings setting = new System.Drawing.Printing.PrinterSettings();
                setting.FromPage = 1;
                setting.ToPage = 1;
                //setting.MinimumPage = 1;
                //setting.MaximumPage = 1;

                // プレビューを実行
                paoRep.Output(setting); 

            }
            else if (radPrint.Checked) //印刷が選択されている場合
            {
                //印刷
                paoRep.Output();
                paoRep = null;
            }
            else if (radPDF.Checked) //PDF出力が選択されている場合
            {

                //ファイルの保存ダイアログの処理
                saveFileDialog.FileName = "郵便番号帳票";
                saveFileDialog.Filter = "PDF形式 (*.pdf)|*.pdf";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //PDF出力
                    paoRep.SavePDF(saveFileDialog.FileName);

                    if (MessageBox.Show(this, "PDFを表示しますか？", "PDF の表示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        ExecFile(saveFileDialog.FileName);
                    }
                }
                paoRep = null;
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
