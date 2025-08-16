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
		private System.Windows.Forms.RadioButton radXPS;
		private System.Windows.Forms.RadioButton radSVG;
		private System.Windows.Forms.RadioButton radPDF;
		private System.Windows.Forms.RadioButton radPrint;
		private System.Windows.Forms.RadioButton radPreview;
		private System.Windows.Forms.Button btnExe;
        private ToolTip toolTip1;
        private RichTextBox txtMessage1;
        private RichTextBox txtMessage2;
        private Button btnExcel;
        private IContainer components;
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
            this.components = new System.ComponentModel.Container();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.radXPS = new System.Windows.Forms.RadioButton();
            this.radSVG = new System.Windows.Forms.RadioButton();
            this.radPDF = new System.Windows.Forms.RadioButton();
            this.radPrint = new System.Windows.Forms.RadioButton();
            this.radPreview = new System.Windows.Forms.RadioButton();
            this.btnExe = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.txtMessage1 = new System.Windows.Forms.RichTextBox();
            this.txtMessage2 = new System.Windows.Forms.RichTextBox();
            this.btnExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // radXPS
            // 
            this.radXPS.Font = new System.Drawing.Font("BIZ UDPゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radXPS.Location = new System.Drawing.Point(504, 24);
            this.radXPS.Name = "radXPS";
            this.radXPS.Size = new System.Drawing.Size(104, 32);
            this.radXPS.TabIndex = 12;
            this.radXPS.Text = "XPS出力";
            this.toolTip1.SetToolTip(this.radXPS, "1. スタート－「設定」－「アプリ」をクリック\r\n2. 「オプション機能の管理」をクリック\r\n3. 「機能の追加」をクリック\r\n4. 「XPS Viewer」をク" +
        "リックし「インストール」をクリック");
            // 
            // radSVG
            // 
            this.radSVG.Font = new System.Drawing.Font("BIZ UDPゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radSVG.Location = new System.Drawing.Point(392, 24);
            this.radSVG.Name = "radSVG";
            this.radSVG.Size = new System.Drawing.Size(95, 32);
            this.radSVG.TabIndex = 11;
            this.radSVG.Text = "SVG出力";
            // 
            // radPDF
            // 
            this.radPDF.Font = new System.Drawing.Font("BIZ UDPゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPDF.Location = new System.Drawing.Point(271, 24);
            this.radPDF.Name = "radPDF";
            this.radPDF.Size = new System.Drawing.Size(98, 32);
            this.radPDF.TabIndex = 10;
            this.radPDF.Text = "PDF出力";
            // 
            // radPrint
            // 
            this.radPrint.Font = new System.Drawing.Font("BIZ UDPゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPrint.Location = new System.Drawing.Point(184, 24);
            this.radPrint.Name = "radPrint";
            this.radPrint.Size = new System.Drawing.Size(96, 32);
            this.radPrint.TabIndex = 9;
            this.radPrint.Text = "印刷";
            // 
            // radPreview
            // 
            this.radPreview.Checked = true;
            this.radPreview.Font = new System.Drawing.Font("BIZ UDPゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPreview.Location = new System.Drawing.Point(72, 24);
            this.radPreview.Name = "radPreview";
            this.radPreview.Size = new System.Drawing.Size(96, 32);
            this.radPreview.TabIndex = 8;
            this.radPreview.TabStop = true;
            this.radPreview.Text = "プレビュー";
            // 
            // btnExe
            // 
            this.btnExe.Font = new System.Drawing.Font("BIZ UDPゴシック", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnExe.ForeColor = System.Drawing.Color.Black;
            this.btnExe.Location = new System.Drawing.Point(21, 80);
            this.btnExe.Name = "btnExe";
            this.btnExe.Size = new System.Drawing.Size(603, 56);
            this.btnExe.TabIndex = 7;
            this.btnExe.Text = "実行";
            this.btnExe.Click += new System.EventHandler(this.btnExe_Click);
            // 
            // toolTip1
            // 
            this.toolTip1.IsBalloon = true;
            this.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.toolTip1.ToolTipTitle = "Windows10/11でXPSビューワーを使う方法";
            // 
            // txtMessage1
            // 
            this.txtMessage1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.txtMessage1.Font = new System.Drawing.Font("BIZ UDPゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMessage1.Location = new System.Drawing.Point(21, 151);
            this.txtMessage1.Name = "txtMessage1";
            this.txtMessage1.ReadOnly = true;
            this.txtMessage1.Size = new System.Drawing.Size(603, 161);
            this.txtMessage1.TabIndex = 13;
            this.txtMessage1.Text = "";
            this.txtMessage1.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.txtMessage_LinkClicked);
            // 
            // txtMessage2
            // 
            this.txtMessage2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.txtMessage2.Font = new System.Drawing.Font("BIZ UDPゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMessage2.Location = new System.Drawing.Point(21, 318);
            this.txtMessage2.Name = "txtMessage2";
            this.txtMessage2.ReadOnly = true;
            this.txtMessage2.Size = new System.Drawing.Size(603, 269);
            this.txtMessage2.TabIndex = 14;
            this.txtMessage2.Text = "";
            this.txtMessage2.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.txtMessage_LinkClicked);
            // 
            // btnExcel
            // 
            this.btnExcel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(222)))));
            this.btnExcel.Font = new System.Drawing.Font("BIZ UDゴシック", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnExcel.ForeColor = System.Drawing.Color.Teal;
            this.btnExcel.Location = new System.Drawing.Point(483, 154);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(137, 48);
            this.btnExcel.TabIndex = 15;
            this.btnExcel.Text = "Excelファイルを開く";
            this.btnExcel.UseVisualStyleBackColor = false;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 12);
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(252)))), ((int)(((byte)(238)))), ((int)(((byte)(235)))));
            this.ClientSize = new System.Drawing.Size(650, 604);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.txtMessage2);
            this.Controls.Add(this.txtMessage1);
            this.Controls.Add(this.radXPS);
            this.Controls.Add(this.radSVG);
            this.Controls.Add(this.radPDF);
            this.Controls.Add(this.radPrint);
            this.Controls.Add(this.radPreview);
            this.Controls.Add(this.btnExe);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Report.newt 使用例 － 郵便番号一覧";
            this.Load += new System.EventHandler(this.Form1_Load);
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
        private string appPath = null;
        // Excelデータベースファイル パス
        private string DbXls = "zip.xls";
        // x64動作時加算パス(フォルダ)
        private string x64dir = "";
        
        private void Form1_Load(object sender, System.EventArgs e)
		{
            // 画面に表示するメッセージの読み込み

            string path = "../../../../";
            if (!System.IO.File.Exists(path + "サンプルプログラムが動作しない時.txt"))
            {
                x64dir += "../";
                path += x64dir;
            }

            txtMessage1.SelectionIndent = 20;
            System.IO.StreamReader sr = new System.IO.StreamReader(
                path + "サンプルプログラムが動作しない時.txt", System.Text.Encoding.GetEncoding("UTF-8"));
            txtMessage1.Text = sr.ReadToEnd();
            sr.Close();

            txtMessage2.SelectionIndent = 20;
            sr = new System.IO.StreamReader(
                path + "Reports.netできること動画集.txt", System.Text.Encoding.GetEncoding("UTF-8"));
            txtMessage2.Text = sr.ReadToEnd();
            sr.Close();

            //カレントPath取得
            appPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "/" + x64dir;
            DbXls = appPath + "../../../" + DbXls;

        }
        
        private void btnExe_Click(object sender, System.EventArgs e)
        {

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
                if(MessageBox.Show("このサンプルプログラムを動作させるためには、データベースへアクセスのため"
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

            DataTable table = ds.Tables["PostTable"];
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
			else if(radPDF.Checked)
			{
				//PDF出力オブジェクトのインスタンスを獲得
				paoRep = ReportCreator.GetPdf();
			}
            else //SVG / XPS 出力が選択されている場合
            {
                //印刷オブジェクトのインスタンスを獲得
                paoRep = ReportCreator.GetReport();
            }


            int page = 0;
            int line = 999;
            string hDate = System.DateTime.Now.ToString();
            
            paoRep.LoadDefFile(appPath + "..\\..\\..\\PaoRep1.prepd");
            foreach (DataRow row in table.Rows){
                line++;
                if (line > 32){ // Head Print
                    if ( page != 0 ) paoRep.PageEnd();
				
					page++;

					if(page == 6)
					{
						paoRep.LoadDefFile(appPath + "..\\..\\..\\PaoRep2.prepd");
					}

					paoRep.PageStart();

                    paoRep.Write("日時", hDate);
                    paoRep.Write("ページ", "Page-" + page.ToString());

					//QRコード描画
					if(page < 6)
					{
						paoRep.Write("QR", row["郵便番号"].ToString() + " " + row["市区町村"].ToString() + row["住所"].ToString());
					}

                    line = 1;		

				}

                //Body Print
                paoRep.Write("郵便番号", row["郵便番号"].ToString(),    line);
                paoRep.Write("市区町村", row["市区町村"].ToString(),    line);
                paoRep.Write("住所",     row["住所"].ToString(),        line);
				paoRep.Write("横罫線", line);


				if(page > 5 && line % 2 == 0)
					paoRep.Write("網掛け", line/2);

            }
            paoRep.PageEnd();

			if(radPreview.Checked || radPrint.Checked) //印刷・プレビューが選択されている場合
			{
				//印刷/プレビュー
				paoRep.Output();
			}
			else if(radPDF.Checked) //PDF出力が選択されている場合
			{

				//ファイルの保存ダイアログの処理
				saveFileDialog.FileName = "郵便番号帳票";
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
            else if(radSVG.Checked) //SVG出力が選択されている場合
            {
                saveFileDialog.FileName = "郵便番号帳票";
                saveFileDialog.Filter = "html形式 (*.html)|*.html";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    paoRep.SaveSVGFile(saveFileDialog.FileName); //SVGデータの保存
                    if (MessageBox.Show(this, "ブラウザで表示しますか？\n表示する場合、SVGプラグインが必要です。", "SVG / SVGZ の表示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        ExecFile(saveFileDialog.FileName);
                    }
                }

            }

            else if (radXPS.Checked) //XPS出力が選択されている場合
            {
                saveFileDialog.FileName = "郵便番号帳票";
                saveFileDialog.Filter = "XPS形式 (*.xps)|*.xps";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    paoRep.SaveXPS(saveFileDialog.FileName); // XPSデータの保存
                    if (MessageBox.Show(this, "XPSを表示しますか？", "XPS の表示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        ExecFile(saveFileDialog.FileName);
                    }
                }

            }

            dataAdapter.Dispose();
        }

        private void txtMessage_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            ExecFile(e.LinkText);
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            ExecFile(DbXls);
        }

        private void ExecFile(string ExecFilePath)
        {
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo(ExecFilePath);
            startInfo.UseShellExecute = true;
            System.Diagnostics.Process.Start(startInfo);
        }
    }
}
