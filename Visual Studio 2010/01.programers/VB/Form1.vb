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
    Friend WithEvents radXPS As System.Windows.Forms.RadioButton
    Friend WithEvents radPdf As System.Windows.Forms.RadioButton
    Friend WithEvents radSvg As System.Windows.Forms.RadioButton
    Friend WithEvents saveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents radPrint As System.Windows.Forms.RadioButton
    Friend WithEvents radPreview As System.Windows.Forms.RadioButton
    Private WithEvents radGetPrintDocument As System.Windows.Forms.RadioButton
    Private WithEvents printPreviewControl1 As System.Windows.Forms.PrintPreviewControl
    Private WithEvents printDocument1 As System.Drawing.Printing.PrintDocument
    Private WithEvents toolTip1 As ToolTip
    Friend WithEvents btnExe As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.radXPS = New System.Windows.Forms.RadioButton()
        Me.radPdf = New System.Windows.Forms.RadioButton()
        Me.radSvg = New System.Windows.Forms.RadioButton()
        Me.saveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.radPrint = New System.Windows.Forms.RadioButton()
        Me.radPreview = New System.Windows.Forms.RadioButton()
        Me.btnExe = New System.Windows.Forms.Button()
        Me.radGetPrintDocument = New System.Windows.Forms.RadioButton()
        Me.printPreviewControl1 = New System.Windows.Forms.PrintPreviewControl()
        Me.printDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'radXPS
        '
        Me.radXPS.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radXPS.Location = New System.Drawing.Point(578, 25)
        Me.radXPS.Name = "radXPS"
        Me.radXPS.Size = New System.Drawing.Size(109, 32)
        Me.radXPS.TabIndex = 11
        Me.radXPS.Text = "XPS出力"
        Me.toolTip1.SetToolTip(Me.radXPS, "1. スタート－「設定」－「アプリ」をクリック" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "2. 「オプション機能の管理」をクリック" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "3. 「機能の追加」をクリック" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "4. 「XPS Viewer」をク" &
        "リックし「インストール」をクリック")
        '
        'radPdf
        '
        Me.radPdf.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPdf.Location = New System.Drawing.Point(342, 25)
        Me.radPdf.Name = "radPdf"
        Me.radPdf.Size = New System.Drawing.Size(110, 32)
        Me.radPdf.TabIndex = 9
        Me.radPdf.Text = "PDF出力"
        '
        'radSvg
        '
        Me.radSvg.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radSvg.Location = New System.Drawing.Point(458, 25)
        Me.radSvg.Name = "radSvg"
        Me.radSvg.Size = New System.Drawing.Size(95, 32)
        Me.radSvg.TabIndex = 10
        Me.radSvg.Text = "SVZ出力"
        '
        'radPrint
        '
        Me.radPrint.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPrint.Location = New System.Drawing.Point(246, 25)
        Me.radPrint.Name = "radPrint"
        Me.radPrint.Size = New System.Drawing.Size(84, 32)
        Me.radPrint.TabIndex = 8
        Me.radPrint.Text = "印刷"
        '
        'radPreview
        '
        Me.radPreview.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPreview.Location = New System.Drawing.Point(117, 25)
        Me.radPreview.Name = "radPreview"
        Me.radPreview.Size = New System.Drawing.Size(118, 32)
        Me.radPreview.TabIndex = 7
        Me.radPreview.Text = "プレビュー"
        '
        'btnExe
        '
        Me.btnExe.Location = New System.Drawing.Point(751, 33)
        Me.btnExe.Name = "btnExe"
        Me.btnExe.Size = New System.Drawing.Size(104, 56)
        Me.btnExe.TabIndex = 6
        Me.btnExe.Text = "実行"
        '
        'radGetPrintDocument
        '
        Me.radGetPrintDocument.Checked = True
        Me.radGetPrintDocument.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radGetPrintDocument.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.radGetPrintDocument.Location = New System.Drawing.Point(116, 61)
        Me.radGetPrintDocument.Name = "radGetPrintDocument"
        Me.radGetPrintDocument.Size = New System.Drawing.Size(508, 44)
        Me.radGetPrintDocument.TabIndex = 13
        Me.radGetPrintDocument.TabStop = True
        Me.radGetPrintDocument.Text = "独自プレビュー  (PrintDocument取得 : ver 6.5.0 新機能)"
        '
        'printPreviewControl1
        '
        Me.printPreviewControl1.AutoZoom = False
        Me.printPreviewControl1.Location = New System.Drawing.Point(24, 123)
        Me.printPreviewControl1.Name = "printPreviewControl1"
        Me.printPreviewControl1.Size = New System.Drawing.Size(831, 430)
        Me.printPreviewControl1.TabIndex = 12
        Me.printPreviewControl1.Zoom = 1.0R
        '
        'toolTip1
        '
        Me.toolTip1.IsBalloon = True
        Me.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.toolTip1.ToolTipTitle = "Windows10/11でXPSビューワーを使う方法"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(876, 566)
        Me.Controls.Add(Me.radGetPrintDocument)
        Me.Controls.Add(Me.printPreviewControl1)
        Me.Controls.Add(Me.radPdf)
        Me.Controls.Add(Me.radSvg)
        Me.Controls.Add(Me.radPrint)
        Me.Controls.Add(Me.radPreview)
        Me.Controls.Add(Me.btnExe)
        Me.Controls.Add(Me.radXPS)
        Me.Name = "Form1"
        Me.Text = "Reports.ne サンプル (10の倍数)"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnExe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExe.Click

        'IReport インターフェースで宣言(印刷・プレビュー・PDF出力どちらでも使える入れ物の用意)
        Dim paoRep As IReport = Nothing

        If radPreview.Checked = True Then   'ラジオボタンでプレビューが選択されている場合
            'プレビューオブジェクトのインスタンスを獲得
            paoRep = ReportCreator.GetPreview()
        ElseIf radPrint.Checked = True Then  'ラジオボタンで印刷が選択されている場合
            '印刷オブジェクトのインスタンスを獲得
            paoRep = ReportCreator.GetReport()
        ElseIf radGetPrintDocument.Checked = True Then  'ラジオボタンで独自プレビュー(GetPrintDocument取得)が選択されている場合

            '印刷オブジェクトのインスタンスを獲得
            paoRep = ReportCreator.GetReport()

            ' ↑ OR ↓(どちらでも可) 

            'プレビューオブジェクトのインスタンスを獲得
            'paoRep = ReportCreator.GetPreview()


        ElseIf radPdf.Checked = True Then    'ラジオボタンでPDF出力が選択されている場合
            'PDF出力オブジェクトのインスタンスを獲得
            paoRep = ReportCreator.GetPdf()
        Else 'ラジオボタンでSVG / XPSが選択されている場合
            '印刷オブジェクトのインスタンスを獲得
            paoRep = ReportCreator.GetReport()
        End If

        'カレントPath取得
        Dim appPath As String = System.IO.Path.GetDirectoryName(Application.ExecutablePath) & "\"

        'レポート定義ファイルの読み込み
        paoRep.LoadDefFile(appPath & "レポート定義ファイル.prepd")

        Dim page As Integer = 0 '頁数を定義
        Dim line As Integer = 0 '行数を定義
        Dim i As Integer
        For i = 1 To 60
            If ((i - 1) Mod 15 = 0) Then '1頁15行で開始
                '頁開始を宣言
                paoRep.PageStart()
                page = page + 1  '頁数をインクリメント
                line = 0 '行数を初期化

                '＊＊＊ヘッダのセット＊＊＊
                '文字列のセット
                paoRep.Write("日付", System.DateTime.Now.ToString())
                paoRep.Write("頁数", "Page - " + page.ToString())

                'オブジェクトの属性変更
                paoRep.z_Objects.SetObject("フォントサイズ")
                paoRep.z_Objects.z_Text.z_FontAttr.Size = 12
                paoRep.Write("フォントサイズ", "フォントサイズ変更後")

                '２頁目の線をを消す
                If page = 2 Then paoRep.Write("Line3", "")

            End If
            line = line + 1 '行数をインクリメント

            '＊＊＊明細のセット＊＊＊
            '繰返し文字列のセット
            paoRep.Write("行番号", i.ToString(), line)
            paoRep.Write("10倍数", (i * 10).ToString(), line)
            '繰返し図形(横線)のセット
            paoRep.Write("横線", line)

            If ((i Mod 15) = 0) Then paoRep.PageEnd() '1頁15行で終了
        Next i

        If radPreview.Checked = True _
        Or radPrint.Checked = True Then '印刷・プレビューが選択されている場合
            ' オマケのコメントです。m(_ _;)m 印刷の設定を色々試してみてください。m(_ _)m
            'Dim setting = New System.Drawing.Printing.PrinterSettings()
            'setting.PrinterName = "Acrobat Distiller"
            'setting.FromPage = 1
            'setting.ToPage = 5
            'setting.MinimumPage = 2
            'setting.MaximumPage = 3
            'paoRep.DisplayDialog = False
            'paoRep.Output(setting) ' 印刷又はプレビューを実行

            'ドキュメント名
            paoRep.DocumentName = "10の倍数の印刷ドキュメント"

            'プレビューウィンドウタイトル
            paoRep.z_PreviewWindow.z_TitleText = "10の倍数の印刷プレビュー"

            'プレビューウィンドウアイコン
            paoRep.z_PreviewWindow.z_Icon = New Icon(appPath + "PreView.ico")

            'バージョンウィンドウの情報変更
            paoRep.z_PreviewWindow.z_VersionWindow.ProductName = "御社製品名"
            paoRep.z_PreviewWindow.z_VersionWindow.ProductName_ForeColor = Color.Blue

            '(初期)プレビュー表示倍率
            paoRep.ZoomPreview = 77


            paoRep.Output() '印刷/プレビューを実行

        ElseIf radGetPrintDocument.Checked = True Then  '独自プレビュー(PrrintDocument取得)が選択されている場合

            ' PrintDocument 取得
            printDocument1 = paoRep.GetPrintDocument()

            ' このフォームのプレビューコントロールへ プレビュー実行
            printPreviewControl1.Document = printDocument1
            printPreviewControl1.InvalidatePreview()

            'ここでは、抜けることにします。(印刷データの保存・読み込み・プレビューはしない)
            Return


        ElseIf radPdf.Checked = True Then  'PDF出力が選択されている場合

            'PDF出力
            saveFileDialog.FileName = "印刷データ"
            saveFileDialog.Filter = "PDF形式 (*.pdf)|*.pdf"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then

                paoRep.SavePDF(saveFileDialog.FileName) '印刷データの保存

                If (MessageBox.Show(Me, "PDFを表示しますか？", "PDF の表示", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                    Dim startInfo As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(saveFileDialog.FileName)
                    startInfo.UseShellExecute = True
                    System.Diagnostics.Process.Start(startInfo)
                End If

            End If

        ElseIf radSvg.Checked = True Then  'SVGが選択されている場合

            'SVG出力
            saveFileDialog.FileName = "印刷データ"
            saveFileDialog.Filter = "html形式 (*.html)|*.html"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then

                'SVGデータの保存
                paoRep.SaveSVGFile(saveFileDialog.FileName)

                If (MessageBox.Show(Me, "ブラウザで表示しますか？" & vbCrLf & "表示する場合、SVGプラグインが必要です。", "SVG / SVGZ の表示", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                    Dim startInfo As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(saveFileDialog.FileName)
                    startInfo.UseShellExecute = True
                    System.Diagnostics.Process.Start(startInfo)
                End If

            End If

        ElseIf radXPS.Checked = True Then 'XPS出力が選択されている場合

            'XPS出力
            saveFileDialog.FileName = "印刷データ"
            saveFileDialog.Filter = "XPS形式 (*.xps)|*.xps"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then

                'XPSデータの保存
                paoRep.SaveXPS(saveFileDialog.FileName)

                If (MessageBox.Show(Me, "XPSを表示しますか？", "XPS の表示", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                    Dim startInfo As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(saveFileDialog.FileName)
                    startInfo.UseShellExecute = True
                    System.Diagnostics.Process.Start(startInfo)
                End If

            End If


        End If


        'マニュアル・ヘルプにはありませんが付け加えました。
        If MessageBox.Show(Me, "続いて、印刷データXMLファイルを保存して再度読み込んでプレビューを行います。", "Save And Reload Print Data", MessageBoxButtons.OKCancel) = DialogResult.Cancel Then
            Exit Sub
        End If

        paoRep.SaveXMLFile("印刷データファイル.prepe") '印刷データの保存

        'プレビューオブジェクトのインスタンスを獲得しなおし(一旦初期化)
        paoRep = ReportCreator.GetPreview()

        paoRep.LoadXMLFile("印刷データファイル.prepe") '印刷データの読み込み

        paoRep.Output() ' プレビューを実行

    End Sub

End Class
