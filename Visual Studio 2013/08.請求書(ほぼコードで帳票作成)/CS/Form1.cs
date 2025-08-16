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
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
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
            this.btnExcel.Location = new System.Drawing.Point(484, 154);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(136, 48);
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
            this.Text = "Reports.net サンプル (請求書) - ほぼコードで帳票作成";
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
        string appPath = null;
        // Excelデータベースファイル パス
        private string DbXls = "請求書.xls";
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
            //インスタンスの生成
            IReport paoRep = null;

            if (radPreview.Checked)
            { // プレビューを選択している場合
                paoRep = ReportCreator.GetPreview();
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
            else //SVG / XPS 出力が選択されている場合
            {
                //印刷オブジェクトのインスタンスを獲得
                paoRep = ReportCreator.GetReport();
            }

 
            //サンプルの「請求書.xls」への接続 Jet4.0を使用
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
            OleDbDataAdapter oda;

            //データセットの作成
            DataSet ds = new DataSet();

            //データセットへテーブルをセットする。ヘッダと明細の2テーブル
            string SQL = "";
            SQL = "select * from [請求ヘッダ$] ORDER BY 請求番号";
            oda = new OleDbDataAdapter(SQL, connection);

            try
            {
                oda.Fill(ds, "請求ヘッダ");
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

            SQL = "select * from [請求明細$] ORDER BY 請求番号, 行番号";
            oda = new OleDbDataAdapter(SQL, connection);
            oda.Fill(ds, "請求明細");

            //請求書の生成
            paoRep.LoadDefFile(appPath + "../../../請求書.prepd");

            // 各列幅調整の配列
            float[] arr_w = {  -5, 44, -20, -10, -9 };

            DataTable ht = ds.Tables["請求ヘッダ"];
            foreach (DataRow hdr in ht.Rows)
            {

                paoRep.PageStart();

                paoRep.Write("txtNo", (string)hdr["請求番号"]);
                paoRep.Write("txtCustomer", (string)hdr["お客様名"]);
                paoRep.Write("txtDate", DateTime.Now.ToString("yyyy年M月d日"));

                // デザイン時の行数・列数取得
                paoRep.z_Objects.SetObject("hLine");
                int maxHLine = paoRep.z_Objects.z_Line.Repeat - 1;
                paoRep.z_Objects.SetObject("vLine");
                int maxVLine = paoRep.z_Objects.z_Line.Repeat - 1;

                //空の表を作成
                for (int i = 0; i < maxHLine; i++)
                {
                    // 「横罫線」描画
                    paoRep.Write("hLine", i + 1);

                    // 外枠の上を太く
                    if (i == 0)
                    {
                        paoRep.z_Objects.SetObject("hLine", i + 1);
                        paoRep.z_Objects.z_Line.z_LineAttr.Width = 0.5f;
                    }

                    // 行ヘッダの下を二重線
                    if (i == 1)
                    {
                        paoRep.z_Objects.SetObject("hLine", i + 1);
                        paoRep.z_Objects.z_Line.z_LineAttr.Type = PmLineType.Double;
                    }

                    // 「行の背景」描画
                    paoRep.Write("LineRect", i + 1);
                    paoRep.z_Objects.SetObject("LineRect", i + 1);

                    if (i == 0)
                    // 行ヘッダはデザイン通り
                    {
                    }
                    else if (i < maxHLine - 3)
                    // 明細行
                    {
                        // 白・青の順番で背景色をセット
                        if (i % 2 == 1)
                        {
                            paoRep.z_Objects.z_Square.PaintColor = System.Drawing.Color.White;
                        }
                        else
                        {
                            paoRep.z_Objects.z_Square.PaintColor = System.Drawing.Color.LightSkyBlue;
                        }
                    }
                    else
                    // 集計行
                    {
                        paoRep.z_Objects.z_Square.PaintColor = Color.FromArgb(255, 255, 180);
                    }


                    // 次回のXの位置
                    float svX = -1;

                    for (int j = 0; j < maxVLine; j++)
                    {

                        // 文字列項目の属性(幅/Font/Align/)変更
                        paoRep.z_Objects.SetObject("field" + (j + 1).ToString(), i + 1);

                        // 幅(TextBox)
                        paoRep.z_Objects.z_Text.Width = paoRep.z_Objects.z_Text.Width + arr_w[j];

                        // 位置(TextBox)
                        if (j > 0)
                        {
                            paoRep.z_Objects.z_Text.X = svX;
                        }
                        svX = paoRep.z_Objects.z_Text.X + paoRep.z_Objects.z_Text.Width;

                        // 行ヘッダの場合
                        if (i == 0)
                        {
                            paoRep.z_Objects.z_Text.z_FontAttr.Bold = true;
                        }
                        // 明細行の場合
                        else
                        {
                            paoRep.z_Objects.z_Text.z_FontAttr.Bold = false;
                            paoRep.z_Objects.z_Text.z_FontAttr.Size = 12;

                            // 文字位置(Text Align)
                            switch (j + 1)
                            {
                                case 1:
                                    paoRep.z_Objects.z_Text.TextAlign = Pao.Reports.PmAlignType.Center;
                                    break;
                                case 2:
                                    paoRep.z_Objects.z_Text.TextAlign = Pao.Reports.PmAlignType.Left;
                                    break;
                                case 3:
                                case 4:
                                case 5:
                                    paoRep.z_Objects.z_Text.TextAlign = Pao.Reports.PmAlignType.Right;
                                    break;
                            }

                        }
                    }
                    //集計行の文字設定
                    for (int j = maxHLine; j > maxHLine - 3; j--)
                    {
                        paoRep.z_Objects.SetObject("field4", j);
                        paoRep.z_Objects.z_Text.z_FontAttr.Size = 16;
                        paoRep.z_Objects.z_Text.TextAlign = Pao.Reports.PmAlignType.Center;
                        paoRep.z_Objects.z_Text.z_FontAttr.Bold = true;
                    }


                }

                // 縦罫線描画
                paoRep.z_Objects.SetObject("vLine");
                float baseX = paoRep.z_Objects.z_Line.X;
                for (int j = 0; j <= maxVLine; j++)
                {
                    paoRep.Write("vLine", j + 1);

                    paoRep.z_Objects.SetObject("vLine", j + 1);

                    //// 幅調整
                    for (int jj = 1; jj <= j && j < maxVLine; jj++)
                    {
                        float baseIntervalX = paoRep.z_Objects.z_Line.IntervalX;
                        paoRep.z_Objects.z_Line.IntervalX = baseIntervalX + arr_w[j - jj];
                    }

                    // 外枠を太線にする
                    if (j == 0 || j == maxVLine)
                    {
                        paoRep.z_Objects.z_Line.z_LineAttr.Width = 0.5f;
                    }

                }


                // 見出し文字入れ
                paoRep.Write("field1", "品番", 1);
                paoRep.Write("field2", "品名", 1);
                paoRep.Write("field3", "数量", 1);
                paoRep.Write("field4", "単価", 1);
                paoRep.Write("field5", "金額", 1);

                //明細の作成
                DataView dv = new DataView(ds.Tables["請求明細"]);
                dv.RowFilter = "請求番号 = '" + (string)hdr["請求番号"] + "'";
                long totalAmount = 0;
                int ii = 0;
                for (; ii < dv.Count; ii++)
                {
                    paoRep.Write("field1", (string)dv[ii]["品番"], ii + 2);
                    paoRep.Write("field2", (string)dv[ii]["品名"], ii + 2);
                    paoRep.Write("field3", dv[ii]["数量"].ToString(), ii + 2);
                    paoRep.Write("field4", string.Format("{0:N0}", dv[ii]["単価"]), ii + 2);
                    long amount = Convert.ToInt64(dv[ii]["数量"]) * Convert.ToInt64(dv[ii]["単価"]);
                    paoRep.Write("field5", string.Format("{0:N0}", amount), ii + 2);
                    totalAmount += amount;

                }

                double tax = 0.05;

                paoRep.Write("field4", "小計", maxHLine - 2);
                paoRep.Write("field5", string.Format("{0:N0}", totalAmount), maxHLine - 2);
                ii++;
                paoRep.Write("field4", "消費税", maxHLine - 1);
                paoRep.Write("field5", string.Format("{0:N0}", totalAmount * tax), maxHLine - 1);
                ii++;
                paoRep.Write("field4", "合計", maxHLine);
                paoRep.Write("field5", string.Format("{0:N0}", totalAmount + (totalAmount * tax)), maxHLine);

                paoRep.Write("txtTotal", string.Format("{0:N0}", totalAmount + (totalAmount * tax)));


                // 小計の上を二重線
                paoRep.z_Objects.SetObject("hLine", maxHLine - 2);
                paoRep.z_Objects.z_Line.z_LineAttr.Type = PmLineType.Double;

                // 最終行を太く
                paoRep.Write("hLine", maxHLine + 1);
                paoRep.z_Objects.SetObject("hLine", maxHLine + 1);
                paoRep.z_Objects.z_Line.z_LineAttr.Width = 0.5f;


                paoRep.PageEnd();

            }


            if (radPreview.Checked || radPrint.Checked) //印刷・プレビューが選択されている場合
            {
                //印刷/プレビュー
                paoRep.Output();
            }
            else if (radPDF.Checked) //PDF出力が選択されている場合
            {

                //ファイルの保存ダイアログの処理
                saveFileDialog.FileName = "請求書";
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
            }
            else if (radSVG.Checked) //SVG出力が選択されている場合
            {
                saveFileDialog.FileName = "請求書";
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
                saveFileDialog.FileName = "請求書";
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
