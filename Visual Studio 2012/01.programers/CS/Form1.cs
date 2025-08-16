using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Pao.Reports;

namespace Pao.Reports.Sample
{
	/// <summary>
	/// Form1 の概要の説明です。
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
		private System.Windows.Forms.Button btnExe;
        private RadioButton radGetPrintDocument;
        private PrintPreviewControl printPreviewControl1;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private RadioButton radPreview;
        private RadioButton radPrint;
        private RadioButton radPDF;
        private RadioButton radSVG;
        private RadioButton radXPS;
        private ToolTip toolTip1;
        private IContainer components;

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
            this.components = new System.ComponentModel.Container();
            this.btnExe = new System.Windows.Forms.Button();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.radGetPrintDocument = new System.Windows.Forms.RadioButton();
            this.printPreviewControl1 = new System.Windows.Forms.PrintPreviewControl();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.radPreview = new System.Windows.Forms.RadioButton();
            this.radPrint = new System.Windows.Forms.RadioButton();
            this.radPDF = new System.Windows.Forms.RadioButton();
            this.radSVG = new System.Windows.Forms.RadioButton();
            this.radXPS = new System.Windows.Forms.RadioButton();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // btnExe
            // 
            this.btnExe.Location = new System.Drawing.Point(733, 42);
            this.btnExe.Name = "btnExe";
            this.btnExe.Size = new System.Drawing.Size(104, 56);
            this.btnExe.TabIndex = 0;
            this.btnExe.Text = "実行";
            this.btnExe.Click += new System.EventHandler(this.button1_Click);
            // 
            // radGetPrintDocument
            // 
            this.radGetPrintDocument.Checked = true;
            this.radGetPrintDocument.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radGetPrintDocument.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.radGetPrintDocument.Location = new System.Drawing.Point(69, 54);
            this.radGetPrintDocument.Name = "radGetPrintDocument";
            this.radGetPrintDocument.Size = new System.Drawing.Size(508, 44);
            this.radGetPrintDocument.TabIndex = 9;
            this.radGetPrintDocument.TabStop = true;
            this.radGetPrintDocument.Text = "独自プレビュー  (PrintDocument取得 : ver 6.5.0 新機能)";
            // 
            // printPreviewControl1
            // 
            this.printPreviewControl1.AutoZoom = false;
            this.printPreviewControl1.Location = new System.Drawing.Point(23, 117);
            this.printPreviewControl1.Name = "printPreviewControl1";
            this.printPreviewControl1.Size = new System.Drawing.Size(831, 430);
            this.printPreviewControl1.TabIndex = 8;
            this.printPreviewControl1.Zoom = 1D;
            // 
            // radPreview
            // 
            this.radPreview.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPreview.Location = new System.Drawing.Point(69, 16);
            this.radPreview.Name = "radPreview";
            this.radPreview.Size = new System.Drawing.Size(111, 32);
            this.radPreview.TabIndex = 1;
            this.radPreview.Text = "プレビュー";
            // 
            // radPrint
            // 
            this.radPrint.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPrint.Location = new System.Drawing.Point(197, 16);
            this.radPrint.Name = "radPrint";
            this.radPrint.Size = new System.Drawing.Size(75, 32);
            this.radPrint.TabIndex = 2;
            this.radPrint.Text = "印刷";
            // 
            // radPDF
            // 
            this.radPDF.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radPDF.Location = new System.Drawing.Point(290, 16);
            this.radPDF.Name = "radPDF";
            this.radPDF.Size = new System.Drawing.Size(101, 32);
            this.radPDF.TabIndex = 3;
            this.radPDF.Text = "PDF出力";
            // 
            // radSVG
            // 
            this.radSVG.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radSVG.Location = new System.Drawing.Point(426, 16);
            this.radSVG.Name = "radSVG";
            this.radSVG.Size = new System.Drawing.Size(113, 32);
            this.radSVG.TabIndex = 4;
            this.radSVG.Text = "SVG出力";
            // 
            // radXPS
            // 
            this.radXPS.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.radXPS.Location = new System.Drawing.Point(575, 16);
            this.radXPS.Name = "radXPS";
            this.radXPS.Size = new System.Drawing.Size(109, 32);
            this.radXPS.TabIndex = 5;
            this.radXPS.Text = "XPS出力";
            this.toolTip1.SetToolTip(this.radXPS, "1. スタート－「設定」－「アプリ」をクリック\r\n2. 「オプション機能の管理」をクリック\r\n3. 「機能の追加」をクリック\r\n4. 「XPS Viewer」をク" +
        "リックし「インストール」をクリック\r\n");
            // 
            // toolTip1
            // 
            this.toolTip1.IsBalloon = true;
            this.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.toolTip1.ToolTipTitle = "Windows10/11でXPSビューワーを使う方法";
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 12);
            this.ClientSize = new System.Drawing.Size(879, 572);
            this.Controls.Add(this.radGetPrintDocument);
            this.Controls.Add(this.printPreviewControl1);
            this.Controls.Add(this.radXPS);
            this.Controls.Add(this.radSVG);
            this.Controls.Add(this.radPDF);
            this.Controls.Add(this.radPrint);
            this.Controls.Add(this.radPreview);
            this.Controls.Add(this.btnExe);
            this.Name = "Form1";
            this.Text = "Reports.ne サンプル (10の倍数)";
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
			//IReport インターフェースで宣言(印刷・レポートどちらでも使える入れ物の用意)
			IReport paoRep = null;

            if (radPreview.Checked) //ラジオボタンでプレビューが選択されている場合
			{
				//プレビューオブジェクトのインスタンスを獲得
				paoRep = ReportCreator.GetPreview();
			}
			else if(radPrint.Checked) // 印刷が選択されている場合
			{
				//印刷オブジェクトのインスタンスを獲得
				paoRep = ReportCreator.GetReport();
			}
            else if (radGetPrintDocument.Checked) //ラジオボタンで独自プレビュー(GetPrintDocument取得)が選択されている場合
            {

                //印刷オブジェクトのインスタンスを獲得
                paoRep = ReportCreator.GetReport();

                // ↑ OR ↓(どちらでも可)  

                //プレビューオブジェクトのインスタンスを獲得
                //paoRep = ReportCreator.GetPreview();


            }
            else if (radPDF.Checked) // PDFが選択されている場合
			{
				//PDF出力オブジェクトのインスタンスを獲得
				paoRep = ReportCreator.GetPdf();
			}
            else //SVG / XPS 出力が選択されている場合
            {
				//印刷オブジェクトのインスタンスを獲得
				paoRep = ReportCreator.GetReport();
			}


			//カレントPath取得
			string appPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\";
			//レポート定義ファイルの読み込み
			paoRep.LoadDefFile(appPath + "レポート定義ファイル.prepd");


			int page = 0; //頁数を定義
			int line = 0; //行数を定義

            for (int i = 0; i < 60; i++)
            {
                if (i % 15 == 0) //1頁15行で開始
                {
                    //頁開始を宣言
                    paoRep.PageStart();
                    page++;		//頁数をインクリメント
                    line = 0;	//行数を初期化

                    //＊＊＊ヘッダのセット＊＊＊
                    //文字列のセット
                    paoRep.Write("日付", System.DateTime.Now.ToString());
                    paoRep.Write("頁数", "Page - " + page.ToString());

                    //オブジェクトの属性変更
                    paoRep.z_Objects.SetObject("フォントサイズ");
                    paoRep.z_Objects.z_Text.z_FontAttr.Size = 12;
                    paoRep.Write("フォントサイズ", "フォントサイズ" + Environment.NewLine + " 変更後");

                    if (page == 2)
                        paoRep.Write("Line3", "");　 //２頁目の線をを消す

                }
                line++; //行数をインクリメント

                //＊＊＊明細のセット＊＊＊
                //繰返し文字列のセット
                paoRep.Write("行番号", (i + 1).ToString(), line);
                paoRep.Write("10倍数", ((i + 1) * 10).ToString(), line);
                //繰返し図形(横線)のセット
                paoRep.Write("横線", line);

                if (((i + 1) % 15) == 0) paoRep.PageEnd(); //1頁15行で終了宣言
            }

			if(radPreview.Checked || radPrint.Checked) //印刷・プレビューが選択されている場合
			{
				//オマケのコメントです。m(_ _;)m 印刷の設定を色々試してみてください。m(_ _)m
				//System.Drawing.Printing.PrinterSettings setting = new System.Drawing.Printing.PrinterSettings();
				//setting.PrinterName = "Acrobat Distiller";
				//setting.FromPage    = 1;
				//setting.ToPage      = 5;
				//setting.MinimumPage = 2;
				//setting.MaximumPage = 3;
				//		
				paoRep.DisplayDialog = true;
				//
				//paoRep.Output(setting); // 印刷又はプレビューを実行

                // ドキュメント名
                paoRep.DocumentName = "10の倍数の印刷ドキュメント";

                // プレビューウィンドウタイトル
                paoRep.z_PreviewWindow.z_TitleText = "10の倍数の印刷プレビュー";

                // プレビューウィンドウアイコン
                paoRep.z_PreviewWindow.z_Icon = new Icon(appPath + "PreView.ico");

                // (初期)プレビュー表示倍率
                paoRep.ZoomPreview = 77;

                // バージョンウィンドウの情報変更
                paoRep.z_PreviewWindow.z_VersionWindow.ProductName = "御社製品名";
                paoRep.z_PreviewWindow.z_VersionWindow.ProductName_ForeColor = Color.Blue;

                MessageBox.Show("ページ数 : " + paoRep.AllPages.ToString());

                paoRep.Output(); // 印刷又はプレビューを実行
			}
            else if (radGetPrintDocument.Checked) // 独自プレビュー(PrrintDocument取得)が選択されている場合
            {
                // PrintDocument 取得
                printDocument1 = paoRep.GetPrintDocument();

                // このフォームのプレビューコントロールへ プレビュー実行
                printPreviewControl1.Document = printDocument1;
                printPreviewControl1.InvalidatePreview();

                // ここでは、抜けることにします。(印刷データの保存・読み込み・プレビューはしない)
                return;

            }
            else if (radPDF.Checked) //PDF出力が選択されている場合
			{

				//PDF出力
				saveFileDialog.FileName = "印刷データ";
				saveFileDialog.Filter = "PDF形式 (*.pdf)|*.pdf";

				if (saveFileDialog.ShowDialog() == DialogResult.OK)
				{
					paoRep.SavePDF(saveFileDialog.FileName); //印刷データの保存

					if(MessageBox.Show(this,"PDFを表示しますか？", "PDF の表示", MessageBoxButtons.YesNo ) == DialogResult.Yes)
					{
                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo(saveFileDialog.FileName);
                        startInfo.UseShellExecute = true;
                        System.Diagnostics.Process.Start(startInfo);
                    }
                }

			}
            else if (radSVG.Checked) //SVG出力が選択されている場合
            {
                saveFileDialog.FileName = "印刷データ";
                saveFileDialog.Filter = "html形式 (*.html)|*.html";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    paoRep.SaveSVGFile(saveFileDialog.FileName); //SVGデータの保存
                    if (MessageBox.Show(this, "ブラウザで表示しますか？\n表示する場合、SVGプラグインが必要です。", "SVG の表示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo(saveFileDialog.FileName);
                        startInfo.UseShellExecute = true;
                        System.Diagnostics.Process.Start(startInfo);
                    }
                }

            }

            else if (radXPS.Checked) //XPS出力が選択されている場合
            {
                saveFileDialog.FileName = "印刷データ";
                saveFileDialog.Filter = "XPS形式 (*.xps)|*.xps";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    paoRep.SaveXPS(saveFileDialog.FileName); // XPSデータの保存
                    if (MessageBox.Show(this, "XPSを表示しますか？", "XPS の表示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo(saveFileDialog.FileName);
                        startInfo.UseShellExecute = true;
                        System.Diagnostics.Process.Start(startInfo);
                    }
                }

            }

            //マニュアル・ヘルプにはありませんが付け加えました。
            if (MessageBox.Show(this, "続いて、印刷データXMLファイルを保存して再度読み込んでプレビューを行います。", "Save And Reload Print Data", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
            {
                return;
            }

			paoRep.SaveXMLFile("印刷データ.prepe"); //印刷データの保存

			//プレビューオブジェクトのインスタンスを獲得しなおし(一旦初期化)
			paoRep = ReportCreator.GetPreview(); 

			paoRep.LoadXMLFile("印刷データ.prepe"); //印刷データの読み込み

			paoRep.Output(); // プレビューを実行

		}

	}
}
