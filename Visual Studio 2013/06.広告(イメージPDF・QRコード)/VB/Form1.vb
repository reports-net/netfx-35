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
    Friend WithEvents btnExe As System.Windows.Forms.Button
    Friend WithEvents radPrint As System.Windows.Forms.RadioButton
    Friend WithEvents radPreview As System.Windows.Forms.RadioButton
    Friend WithEvents radImagePDF As System.Windows.Forms.RadioButton
    Friend WithEvents saveFileDialog As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnExe = New System.Windows.Forms.Button()
        Me.radPrint = New System.Windows.Forms.RadioButton()
        Me.radPreview = New System.Windows.Forms.RadioButton()
        Me.radImagePDF = New System.Windows.Forms.RadioButton()
        Me.saveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.SuspendLayout()
        '
        'btnExe
        '
        Me.btnExe.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnExe.Location = New System.Drawing.Point(24, 84)
        Me.btnExe.Name = "btnExe"
        Me.btnExe.Size = New System.Drawing.Size(329, 53)
        Me.btnExe.TabIndex = 4
        Me.btnExe.Text = "実行"
        '
        'radPrint
        '
        Me.radPrint.Font = New System.Drawing.Font("MS UI Gothic", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPrint.Location = New System.Drawing.Point(155, 32)
        Me.radPrint.Name = "radPrint"
        Me.radPrint.Size = New System.Drawing.Size(85, 24)
        Me.radPrint.TabIndex = 3
        Me.radPrint.Text = "印刷"
        '
        'radPreview
        '
        Me.radPreview.Checked = True
        Me.radPreview.Font = New System.Drawing.Font("MS UI Gothic", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPreview.Location = New System.Drawing.Point(42, 32)
        Me.radPreview.Name = "radPreview"
        Me.radPreview.Size = New System.Drawing.Size(107, 24)
        Me.radPreview.TabIndex = 2
        Me.radPreview.TabStop = True
        Me.radPreview.Text = "プレビュー"
        '
        'radImagePDF
        '
        Me.radImagePDF.Font = New System.Drawing.Font("MS UI Gothic", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radImagePDF.Location = New System.Drawing.Point(246, 32)
        Me.radImagePDF.Name = "radImagePDF"
        Me.radImagePDF.Size = New System.Drawing.Size(126, 24)
        Me.radImagePDF.TabIndex = 5
        Me.radImagePDF.Text = "イメージPDF"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(252, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(235, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(385, 163)
        Me.Controls.Add(Me.radImagePDF)
        Me.Controls.Add(Me.btnExe)
        Me.Controls.Add(Me.radPrint)
        Me.Controls.Add(Me.radPreview)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Report.NET サンプル (広告)"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnExe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExe.Click


        'カレントPath取得
        ' プログラム実行フォルダ
        Dim appPath As String = System.IO.Path.GetDirectoryName(Application.ExecutablePath) & "/"
        ' Excelデータベースファイル パス
        Dim DbXls As String = "広告.xls"
        ' x64動作時加算パス(フォルダ)
        Dim x64dir As String = ""

        If System.IO.File.Exists(appPath & "../../" + DbXls) = False Then
            x64dir += "../../"
            appPath += x64dir
        End If
        DbXls = appPath & "../../" + DbXls

        'サンプルの「広告.xls」への接続 Jetエンジンを使用
        Dim connectString As String
        If IntPtr.Size = 4 Then
            connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbXls & ";Extended Properties=Excel 8.0;"
        Else
            connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DbXls & ";Extended Properties=Excel 12.0;"
        End If
        Dim connection As OleDbConnection = New OleDbConnection(connectString)

        'データセットへテーブルをセットする。
        Dim SQL As String = ""
        SQL = "select * from [広告情報$]"

        Dim dataAdapter As OleDbDataAdapter = New OleDbDataAdapter(SQL, connection)
        Dim ds As DataSet = New DataSet()
        Try
            dataAdapter.Fill(ds, "広告情報")
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

        Dim table As DataTable = ds.Tables("広告情報")

        'インスタンスの生成
        Dim paoRep As IReport = Nothing

        If radPreview.Checked Then
            'プレビューを選択している場合
            paoRep = ReportCreator.GetPreview()
        ElseIf radPrint.Checked Then
            '印刷の場合
            paoRep = ReportCreator.GetReport()
        Else
            'イメージPDF出力の場合
            paoRep = ReportCreator.GetImagePdf()
        End If

        paoRep.LoadDefFile(appPath & "../../広告.prepd")

        Dim row As DataRow
        For Each row In table.Rows

            paoRep.PageStart()

            paoRep.Write("製品名", CStr(row("製品名")))
            paoRep.Write("キャッチフレーズ", CStr(row("キャッチフレーズ")))
            paoRep.Write("商品コード", CStr(row("商品コード")))
            paoRep.Write("JANコード", CStr(row("商品コード")))
            paoRep.Write("売り文句", CStr(row("売り文句")))
            paoRep.Write("説明", CStr(row("説明")))
            paoRep.Write("価格", CStr(row("価格")))
            paoRep.Write("画像1", appPath & "../../" & CStr(row("画像1")))
            paoRep.Write("画像2", appPath & "../../" & CStr(row("画像2")))
            paoRep.Write("QR", CStr(row("製品名")) & CStr(row("キャッチフレーズ")))

            paoRep.PageEnd()
        Next

        If radImagePDF.Checked = False Then '印刷・プレビューが選択されている場合

            '印刷/プレビュー
            paoRep.Output()

        Else 'PDF出力が選択されている場合

            'PDF出力
            SaveFileDialog.FileName = "広告"
            SaveFileDialog.Filter = "PDF形式 (*.pdf)|*.pdf"

            If SaveFileDialog.ShowDialog() = DialogResult.OK Then

                'PDF出力
                paoRep.SavePDF(SaveFileDialog.FileName)

                If (MessageBox.Show(Me, "PDFを表示しますか？", "PDF の表示", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                    ExecFile(saveFileDialog.FileName)
                End If

            End If

        End If
    End Sub
    Private Sub ExecFile(ExecFilePath As String)
        Dim startInfo As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(ExecFilePath)
        startInfo.UseShellExecute = True
        System.Diagnostics.Process.Start(startInfo)
    End Sub
End Class
