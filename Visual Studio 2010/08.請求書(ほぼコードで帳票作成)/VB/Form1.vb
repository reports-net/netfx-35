Imports System.Data.OleDb
Imports Pao.Reports

Public Class Form1
    Inherits System.Windows.Forms.Form


#Region " Windows フォーム デザイナで生成されたコード "

    Public Sub New()
        MyBase.New()

        ' この呼び出しは Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後に初期化を追加します。

    End Sub

    ' Form は dispose をオーバーライドしてコンポーネント一覧を消去します。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    ' メモ : 以下のプロシージャは、Windows フォーム デザイナで必要です。
    ' Windows フォーム デザイナを使って変更してください。  
    ' コード エディタは使用しないでください。
    Friend WithEvents saveFileDialog As System.Windows.Forms.SaveFileDialog
    Private WithEvents radXPS As RadioButton
    Private WithEvents radSVG As RadioButton
    Private WithEvents radPDF As RadioButton
    Private WithEvents radPrint As RadioButton
    Private WithEvents radPreview As RadioButton
    Private WithEvents btnExe As Button
    Private WithEvents toolTip1 As ToolTip
    Private WithEvents btnExcel As Button
    Private WithEvents txtMessage1 As RichTextBox
    Private WithEvents txtMessage2 As RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.saveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.radXPS = New System.Windows.Forms.RadioButton()
        Me.radSVG = New System.Windows.Forms.RadioButton()
        Me.radPDF = New System.Windows.Forms.RadioButton()
        Me.radPrint = New System.Windows.Forms.RadioButton()
        Me.radPreview = New System.Windows.Forms.RadioButton()
        Me.btnExe = New System.Windows.Forms.Button()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.txtMessage2 = New System.Windows.Forms.RichTextBox()
        Me.txtMessage1 = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'toolTip1
        '
        Me.toolTip1.IsBalloon = True
        Me.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.toolTip1.ToolTipTitle = "Windows10/11でXPSビューワーを使う方法"
        '
        'radXPS
        '
        Me.radXPS.Font = New System.Drawing.Font("BIZ UDPゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radXPS.Location = New System.Drawing.Point(509, 16)
        Me.radXPS.Name = "radXPS"
        Me.radXPS.Size = New System.Drawing.Size(104, 32)
        Me.radXPS.TabIndex = 21
        Me.radXPS.Text = "XPS出力"
        Me.toolTip1.SetToolTip(Me.radXPS, "1. スタート－「設定」－「アプリ」をクリック" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "2. 「オプション機能の管理」をクリック" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "3. 「機能の追加」をクリック" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "4. 「XPS Viewer」をク" &
        "リックし「インストール」をクリック")
        '
        'radSVG
        '
        Me.radSVG.Font = New System.Drawing.Font("BIZ UDPゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radSVG.Location = New System.Drawing.Point(397, 16)
        Me.radSVG.Name = "radSVG"
        Me.radSVG.Size = New System.Drawing.Size(95, 32)
        Me.radSVG.TabIndex = 20
        Me.radSVG.Text = "SVG出力"
        '
        'radPDF
        '
        Me.radPDF.Font = New System.Drawing.Font("BIZ UDPゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPDF.Location = New System.Drawing.Point(276, 16)
        Me.radPDF.Name = "radPDF"
        Me.radPDF.Size = New System.Drawing.Size(98, 32)
        Me.radPDF.TabIndex = 19
        Me.radPDF.Text = "PDF出力"
        '
        'radPrint
        '
        Me.radPrint.Font = New System.Drawing.Font("BIZ UDPゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPrint.Location = New System.Drawing.Point(189, 16)
        Me.radPrint.Name = "radPrint"
        Me.radPrint.Size = New System.Drawing.Size(96, 32)
        Me.radPrint.TabIndex = 18
        Me.radPrint.Text = "印刷"
        '
        'radPreview
        '
        Me.radPreview.Checked = True
        Me.radPreview.Font = New System.Drawing.Font("BIZ UDPゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPreview.Location = New System.Drawing.Point(77, 16)
        Me.radPreview.Name = "radPreview"
        Me.radPreview.Size = New System.Drawing.Size(96, 32)
        Me.radPreview.TabIndex = 17
        Me.radPreview.TabStop = True
        Me.radPreview.Text = "プレビュー"
        '
        'btnExe
        '
        Me.btnExe.Font = New System.Drawing.Font("BIZ UDPゴシック", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnExe.Location = New System.Drawing.Point(26, 72)
        Me.btnExe.Name = "btnExe"
        Me.btnExe.Size = New System.Drawing.Size(599, 56)
        Me.btnExe.TabIndex = 16
        Me.btnExe.Text = "実行"
        '
        'btnExcel
        '
        Me.btnExcel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(222, Byte), Integer))
        Me.btnExcel.Font = New System.Drawing.Font("BIZ UDゴシック", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnExcel.ForeColor = System.Drawing.Color.Teal
        Me.btnExcel.Location = New System.Drawing.Point(488, 146)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(137, 48)
        Me.btnExcel.TabIndex = 24
        Me.btnExcel.Text = "Excelファイルを開く"
        Me.btnExcel.UseVisualStyleBackColor = False
        '
        'txtMessage2
        '
        Me.txtMessage2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtMessage2.Font = New System.Drawing.Font("BIZ UDPゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtMessage2.Location = New System.Drawing.Point(26, 310)
        Me.txtMessage2.Name = "txtMessage2"
        Me.txtMessage2.ReadOnly = True
        Me.txtMessage2.Size = New System.Drawing.Size(603, 269)
        Me.txtMessage2.TabIndex = 23
        Me.txtMessage2.Text = ""
        '
        'txtMessage1
        '
        Me.txtMessage1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtMessage1.Font = New System.Drawing.Font("BIZ UDPゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtMessage1.Location = New System.Drawing.Point(26, 143)
        Me.txtMessage1.Name = "txtMessage1"
        Me.txtMessage1.ReadOnly = True
        Me.txtMessage1.Size = New System.Drawing.Size(603, 161)
        Me.txtMessage1.TabIndex = 22
        Me.txtMessage1.Text = ""
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(252, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(235, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(654, 594)
        Me.Controls.Add(Me.btnExcel)
        Me.Controls.Add(Me.txtMessage1)
        Me.Controls.Add(Me.txtMessage2)
        Me.Controls.Add(Me.radXPS)
        Me.Controls.Add(Me.radSVG)
        Me.Controls.Add(Me.radPDF)
        Me.Controls.Add(Me.radPrint)
        Me.Controls.Add(Me.radPreview)
        Me.Controls.Add(Me.btnExe)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reports.ne サンプル (請求書)"
        Me.ResumeLayout(False)

    End Sub

#End Region

    ' プログラム実行フォルダ
    Private appPath As String = Nothing
    ' Excelデータベースファイル パス
    Private DbXls As String = "請求書.xls"
    ' x64動作時加算パス(フォルダ)
    Private x64dir As String = ""
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim path As String = "../../../"

        If System.IO.File.Exists(path & "サンプルプログラムが動作しない時.txt") = False Then
            x64dir += "../../"
            path += x64dir
        End If

        txtMessage1.SelectionIndent = 20
        Dim sr As New System.IO.StreamReader(
            path & "サンプルプログラムが動作しない時.txt", System.Text.Encoding.GetEncoding("UTF-8"))
        txtMessage1.Text = sr.ReadToEnd()
        sr.Close()

        txtMessage2.SelectionIndent = 20
        sr = New System.IO.StreamReader(
            path & "Reports.netできること動画集.txt", System.Text.Encoding.GetEncoding("UTF-8"))
        txtMessage2.Text = sr.ReadToEnd()
        sr.Close()

        'カレントPath取得
        appPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) & "/" & x64dir
        DbXls = appPath & "../../" & DbXls

    End Sub

    Private Sub btnExe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExe.Click


        'サンプルの「請求書.xls」への接続 Jet4.0を使用
        Dim connectString As String
        If IntPtr.Size = 4 Then
            connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbXls & ";Extended Properties=Excel 8.0;"
        Else
            connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DbXls & ";Extended Properties=Excel 12.0;"
        End If
        Dim connection As OleDbConnection = New OleDbConnection(connectString)

        Dim oda As OleDbDataAdapter

        'データセットの作成
        Dim ds As DataSet = New DataSet

        'データセットへテーブルをセットする。ヘッダと明細の2テーブル
        Dim SQL As String = ""
        SQL = "select * from [請求ヘッダ$] ORDER BY 請求番号"
        oda = New OleDbDataAdapter(SQL, connection)

        Try
            oda.Fill(ds, "請求ヘッダ")
        Catch
            If MessageBox.Show("このサンプルプログラムを動作させるためには、データベースへアクセスのため" _
                    & Environment.NewLine + "[Microsoft Access データベース エンジン 2010 再頒布可能コンポーネント]" _
                    & Environment.NewLine + "をインストールする必要があります。" _
                    & Environment.NewLine + "マイクロソフトのインストーラ ダウンロードサイトへジャンプしますか？" _
                    , "サンプルが動作しない時", MessageBoxButtons.YesNo, MessageBoxIcon.Information) _
                = DialogResult.Yes Then
                ExecFile("http://www.microsoft.com/ja-jp/download/details.aspx?id=13255")
            End If
            Return
        End Try
        SQL = "select * from [請求明細$] ORDER BY 請求番号, 行番号"
        oda = New OleDbDataAdapter(SQL, connection)
        oda.Fill(ds, "請求明細")

        'インスタンスの生成
        Dim paoRep As IReport = Nothing

        If (radPreview.Checked) Then
            'プレビューを選択している場合
            paoRep = ReportCreator.GetPreview()
        ElseIf radPrint.Checked Then
            '印刷の場合
            paoRep = ReportCreator.GetReport()
        ElseIf radPdf.Checked = True Then
            'イメージPDFを選択されている場合
            paoRep = ReportCreator.GetPdf()
        Else
            'SVG / XPSを選択されている場合
            paoRep = ReportCreator.GetReport()
        End If



        ' 各列幅調整の配列
        Dim arr_w As Single() = { -5, 44, -20, -10, -9}

        Dim ht As DataTable = ds.Tables("請求ヘッダ")

        Dim hdr As DataRow
        Dim i As Integer = 0
        Dim j As Integer = 0

        '請求書の生成
        paoRep.LoadDefFile(appPath & "../../請求書.prepd")

        For Each hdr In ht.Rows

            paoRep.PageStart()

            paoRep.Write("txtNo", CStr(hdr("請求番号")))
            paoRep.Write("txtCustomer", CStr(hdr("お客様名")))
            paoRep.Write("txtDate", DateTime.Now.ToString("yyyy年M月d日"))


            ' デザイン時の行数・列数取得
            paoRep.z_Objects.SetObject("hLine")
            Dim maxHLine As Integer = paoRep.z_Objects.z_Line.Repeat - 1
            paoRep.z_Objects.SetObject("vLine")
            Dim maxVLine As Integer = paoRep.z_Objects.z_Line.Repeat - 1

            '空の表を作成
            For i = 0 To maxHLine - 1

                ' 「横罫線」描画
                paoRep.Write("hLine", i + 1)

                ' 外枠の上を太く
                If i = 0 Then
                    paoRep.z_Objects.SetObject("hLine", i + 1)
                    paoRep.z_Objects.z_Line.z_LineAttr.Width = 0.5F
                End If

                ' 行ヘッダの下を二重線
                If i = 1 Then
                    paoRep.z_Objects.SetObject("hLine", i + 1)
                    paoRep.z_Objects.z_Line.z_LineAttr.Type = PmLineType.Double
                End If

                ' 「行の背景」描画
                paoRep.Write("LineRect", i + 1)
                paoRep.z_Objects.SetObject("LineRect", i + 1)

                If i = 0 Then
                    ' 行ヘッダはデザイン通り

                ElseIf i < maxHLine - 3 Then
                    ' 明細行
                    ' 白・青の順番で背景色をセット
                    If i Mod 2 = 1 Then
                        paoRep.z_Objects.z_Square.PaintColor = System.Drawing.Color.White
                    Else
                        paoRep.z_Objects.z_Square.PaintColor = System.Drawing.Color.LightSkyBlue
                    End If
                Else
                    ' 集計行
                    paoRep.z_Objects.z_Square.PaintColor = Color.FromArgb(255, 255, 180)

                End If

                ' 次回のXの位置
                Dim svX As Single = -1

                For j = 0 To maxVLine - 1

                    ' 文字列項目の属性(幅/Font/Align/)変更
                    paoRep.z_Objects.SetObject("field" + (j + 1).ToString(), i + 1)

                    ' 幅(TextBox)
                    paoRep.z_Objects.z_Text.Width = paoRep.z_Objects.z_Text.Width + arr_w(j)

                    ' 位置(TextBox)
                    If j > 0 Then
                        paoRep.z_Objects.z_Text.X = svX
                    End If
                    svX = paoRep.z_Objects.z_Text.X + paoRep.z_Objects.z_Text.Width

                    ' 行ヘッダの場合
                    If i = 0 Then
                        paoRep.z_Objects.z_Text.z_FontAttr.Bold = True
                    Else
                        ' 明細行の場合
                        paoRep.z_Objects.z_Text.z_FontAttr.Bold = False
                        paoRep.z_Objects.z_Text.z_FontAttr.Size = 12

                        ' 文字位置(Text Align)
                        Select Case (j + 1)
                            Case 1
                                paoRep.z_Objects.z_Text.TextAlign = Pao.Reports.PmAlignType.Center
                            Case 2
                                paoRep.z_Objects.z_Text.TextAlign = Pao.Reports.PmAlignType.Left
                            Case 3, 4, 5
                                paoRep.z_Objects.z_Text.TextAlign = Pao.Reports.PmAlignType.Right
                        End Select

                    End If

                Next

                '集計行の文字設定
                Dim k As Integer = maxHLine
                For k = maxHLine To maxHLine - 3 + 1 Step -1
                    paoRep.z_Objects.SetObject("field4", k)
                    paoRep.z_Objects.z_Text.z_FontAttr.Size = 16
                    paoRep.z_Objects.z_Text.TextAlign = Pao.Reports.PmAlignType.Center
                    paoRep.z_Objects.z_Text.z_FontAttr.Bold = True
                Next


            Next

            ' 縦罫線描画
            paoRep.z_Objects.SetObject("vLine")
            Dim baseX As Single = paoRep.z_Objects.z_Line.X
            For j = 0 To maxVLine
                paoRep.Write("vLine", j + 1)

                paoRep.z_Objects.SetObject("vLine", j + 1)

                '' 幅調整
                Dim jj As Integer = 1
                For jj = 1 To j
                    If j = maxVLine Then Exit For
                    Dim baseIntervalX As Single = paoRep.z_Objects.z_Line.IntervalX
                    paoRep.z_Objects.z_Line.IntervalX = baseIntervalX + arr_w(j - jj)
                Next

                ' 外枠を太線にする
                If j = 0 Or j = maxVLine Then
                    paoRep.z_Objects.z_Line.z_LineAttr.Width = 0.5F
                End If

            Next


            ' 見出し文字入れ
            paoRep.Write("field1", "品番", 1)
            paoRep.Write("field2", "品名", 1)
            paoRep.Write("field3", "数量", 1)
            paoRep.Write("field4", "単価", 1)
            paoRep.Write("field5", "金額", 1)

            '明細の作成
            Dim dv As DataView
            dv = New DataView(ds.Tables("請求明細"))
            dv.RowFilter = "請求番号 = '" & hdr("請求番号").ToString() & "'"
            Dim totalAmount As Long = 0
            Dim ii As Integer = 0
            For ii = 0 To dv.Count - 1
                paoRep.Write("field1", dv(ii).Item("品番").ToString(), ii + 2)
                paoRep.Write("field2", dv(ii).Item("品名").ToString(), ii + 2)
                paoRep.Write("field3", dv(ii).Item("数量").ToString(), ii + 2)
                paoRep.Write("field4", String.Format("{0:N0}", dv(ii).Item("単価")), ii + 2)
                Dim amount As Long = Convert.ToInt64(dv(ii).Item("数量")) * Convert.ToInt64(dv(ii).Item("単価"))
                paoRep.Write("field5", String.Format("{0:N0}", amount), ii + 2)
                totalAmount += amount
            Next

            Dim tax As Double = 0.05

            paoRep.Write("field4", "小計", maxHLine - 2)
            paoRep.Write("field5", String.Format("{0:N0}", totalAmount), maxHLine - 2)
            ii = ii + 1
            paoRep.Write("field4", "消費税", maxHLine - 1)
            paoRep.Write("field5", String.Format("{0:N0}", totalAmount * tax), maxHLine - 1)
            ii = ii + 1
            paoRep.Write("field4", "合計", maxHLine)
            paoRep.Write("field5", String.Format("{0:N0}", totalAmount + (totalAmount * tax)), maxHLine)

            paoRep.Write("txtTotal", String.Format("{0:N0}", totalAmount + (totalAmount * tax)))


            ' 小計の上を二重線
            paoRep.z_Objects.SetObject("hLine", maxHLine - 2)
            paoRep.z_Objects.z_Line.z_LineAttr.Type = PmLineType.Double

            ' 最終行を太く
            paoRep.Write("hLine", maxHLine + 1)
            paoRep.z_Objects.SetObject("hLine", maxHLine + 1)
            paoRep.z_Objects.z_Line.z_LineAttr.Width = 0.5F

            paoRep.PageEnd()

        Next


        If radPreview.Checked = True Or radPrint.Checked = True Then '印刷・プレビューが選択されている場合

            '印刷/プレビュー
            paoRep.Output()

        ElseIf radPdf.Checked = True Then 'PDF出力が選択されている場合

            'PDF出力
            saveFileDialog.FileName = "請求書"
            saveFileDialog.Filter = "PDF形式 (*.pdf)|*.pdf"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then

                'PDF出力
                paoRep.SavePDF(saveFileDialog.FileName)

                If (MessageBox.Show(Me, "PDFを表示しますか？", "PDF の表示", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                    ExecFile(saveFileDialog.FileName)
                End If

            End If

        ElseIf radSvg.Checked = True Then  'SVGが選択されている場合

            'SVG / SVGZ出力
            saveFileDialog.FileName = "請求書"
            saveFileDialog.Filter = "html形式 (*.html)|*.html"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then

                'SVGデータの保存
                paoRep.SaveSVGFile(saveFileDialog.FileName)

                If (MessageBox.Show(Me, "ブラウザで表示しますか？" & vbCrLf & "表示する場合、SVGプラグインが必要です。", "SVG / SVGZ の表示", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                    ExecFile(saveFileDialog.FileName)
                End If

            End If

        ElseIf radXPS.Checked = True Then 'XPS出力が選択されている場合

            'XPS出力
            saveFileDialog.FileName = "請求書"
            saveFileDialog.Filter = "XPS形式 (*.xps)|*.xps"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then

                'XPSデータの保存
                paoRep.SaveXPS(saveFileDialog.FileName)

                If (MessageBox.Show(Me, "XPSを表示しますか？", "XPS の表示", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                    ExecFile(saveFileDialog.FileName)
                End If

            End If
        End If

    End Sub

    Private Sub txtMessage_LinkClicked(sender As Object, e As LinkClickedEventArgs) Handles txtMessage1.LinkClicked, txtMessage2.LinkClicked
        ExecFile(e.LinkText)
    End Sub
    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        ExecFile(DbXls)
    End Sub

    Private Sub ExecFile(ExecFilePath As String)
        Dim startInfo As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(ExecFilePath)
        startInfo.UseShellExecute = True
        System.Diagnostics.Process.Start(startInfo)
    End Sub

End Class
