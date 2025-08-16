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
    Private WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents radD2 As System.Windows.Forms.RadioButton
    Private WithEvents radD1 As System.Windows.Forms.RadioButton
    Private WithEvents btnPreview As System.Windows.Forms.Button
    Private WithEvents radPDF As System.Windows.Forms.RadioButton
    Private WithEvents radPrint As System.Windows.Forms.RadioButton
    Private WithEvents btnExe As System.Windows.Forms.Button

    ' メモ : 以下のプロシージャは、Windows フォーム デザイナで必要です。
    ' Windows フォーム デザイナを使って変更してください。  
    ' コード エディタは使用しないでください。
    Friend WithEvents saveFileDialog As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.saveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.radD2 = New System.Windows.Forms.RadioButton()
        Me.radD1 = New System.Windows.Forms.RadioButton()
        Me.btnPreview = New System.Windows.Forms.Button()
        Me.radPDF = New System.Windows.Forms.RadioButton()
        Me.radPrint = New System.Windows.Forms.RadioButton()
        Me.btnExe = New System.Windows.Forms.Button()
        Me.groupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.radD2)
        Me.groupBox1.Controls.Add(Me.radD1)
        Me.groupBox1.Controls.Add(Me.btnPreview)
        Me.groupBox1.Font = New System.Drawing.Font("Meiryo UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.groupBox1.Location = New System.Drawing.Point(70, 76)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(492, 174)
        Me.groupBox1.TabIndex = 17
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "デザイン選択＆プレビュー (帳票データを再読み込みしません)"
        '
        'radD2
        '
        Me.radD2.AutoSize = True
        Me.radD2.Location = New System.Drawing.Point(278, 50)
        Me.radD2.Name = "radD2"
        Me.radD2.Size = New System.Drawing.Size(93, 24)
        Me.radD2.TabIndex = 2
        Me.radD2.Text = "デザイン２"
        Me.radD2.UseVisualStyleBackColor = True
        '
        'radD1
        '
        Me.radD1.AutoSize = True
        Me.radD1.Checked = True
        Me.radD1.Location = New System.Drawing.Point(122, 50)
        Me.radD1.Name = "radD1"
        Me.radD1.Size = New System.Drawing.Size(93, 24)
        Me.radD1.TabIndex = 1
        Me.radD1.TabStop = True
        Me.radD1.Text = "デザイン１"
        Me.radD1.UseVisualStyleBackColor = True
        '
        'btnPreview
        '
        Me.btnPreview.BackColor = System.Drawing.Color.DarkBlue
        Me.btnPreview.FlatAppearance.BorderSize = 5
        Me.btnPreview.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPreview.Font = New System.Drawing.Font("Meiryo UI", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnPreview.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.btnPreview.Location = New System.Drawing.Point(89, 91)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(311, 57)
        Me.btnPreview.TabIndex = 0
        Me.btnPreview.Text = "プレビューしてデザインを確認"
        Me.btnPreview.UseVisualStyleBackColor = False
        '
        'radPDF
        '
        Me.radPDF.Checked = True
        Me.radPDF.Font = New System.Drawing.Font("Meiryo UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPDF.Location = New System.Drawing.Point(412, 298)
        Me.radPDF.Name = "radPDF"
        Me.radPDF.Size = New System.Drawing.Size(96, 32)
        Me.radPDF.TabIndex = 16
        Me.radPDF.TabStop = True
        Me.radPDF.Text = "PDF出力"
        '
        'radPrint
        '
        Me.radPrint.Font = New System.Drawing.Font("Meiryo UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPrint.Location = New System.Drawing.Point(542, 298)
        Me.radPrint.Name = "radPrint"
        Me.radPrint.Size = New System.Drawing.Size(75, 32)
        Me.radPrint.TabIndex = 15
        Me.radPrint.Text = "印刷"
        '
        'btnExe
        '
        Me.btnExe.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.btnExe.FlatAppearance.BorderSize = 2
        Me.btnExe.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExe.Font = New System.Drawing.Font("Meiryo UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnExe.ForeColor = System.Drawing.Color.Black
        Me.btnExe.Location = New System.Drawing.Point(404, 331)
        Me.btnExe.Name = "btnExe"
        Me.btnExe.Size = New System.Drawing.Size(213, 44)
        Me.btnExe.TabIndex = 14
        Me.btnExe.Text = "実行"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(9, 21)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(251, Byte), Integer), CType(CType(218, Byte), Integer), CType(CType(222, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(624, 387)
        Me.Controls.Add(Me.groupBox1)
        Me.Controls.Add(Me.radPDF)
        Me.Controls.Add(Me.radPrint)
        Me.Controls.Add(Me.btnExe)
        Me.Font = New System.Drawing.Font("Meiryo UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reports.ne サンプル (デザイン選択)"
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        Output(True)
    End Sub

    Private Sub btnExe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExe.Click
        Output(False)
    End Sub

    Dim paoRep As IReport = Nothing
    Dim _loadDesign As Boolean = False

    Private Sub Output(ByVal flgPreview As Boolean)


        'カレントPath取得
        Dim appPath As String = System.IO.Path.GetDirectoryName(Application.ExecutablePath) & "/"
        ' Excelデータベースファイル パス
        Dim DbXls As String = "zip.xls"
        ' x64動作時加算パス(フォルダ)
        Dim x64dir As String = ""

        If System.IO.File.Exists(appPath & "../../" + DbXls) = False Then
            x64dir += "../../"
            appPath += x64dir
        End If
        DbXls = appPath & "../../" + DbXls


        If flgPreview Then

            If paoRep Is Nothing Then

                'プレビューオブジェクトのインスタンスを獲得
                paoRep = ReportCreator.GetPreview()

            End If

        ElseIf radPrint.Checked Then

            '印刷オブジェクトのインスタンスを獲得
            paoRep = ReportCreator.GetReport()

        ElseIf radPDF.Checked Then

            'PDF出力オブジェクトのインスタンスを獲得
            paoRep = ReportCreator.GetPdf()

        End If


        Dim page As Integer = 0
        Dim line As Integer = 999


        Dim defFile As String() = {appPath + "../../PaoRep1.prepd", appPath + "../../PaoRep2.prepd"}
        Dim defIndex As Integer = 0

        If radD2.Checked Then defIndex = 1

        Dim hDate As String = System.DateTime.Now.ToString()

        If _loadDesign = False Then
            paoRep.LoadDefFile(defFile(defIndex))
        Else
            paoRep.ChangeDefFile(defFile(defIndex))
        End If

        If _loadDesign = False Then

            _loadDesign = True

            ' データ取得
            Dim connectString As String
            If IntPtr.Size = 4 Then
                connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbXls & ";Extended Properties=Excel 8.0;"
            Else
                connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DbXls & ";Extended Properties=Excel 12.0;"
            End If

            Dim connection As OleDbConnection = New OleDbConnection(connectString)

            'データセットへテーブルをセットする。
            Dim SQL As String = ""
            SQL = "select * from [郵便番号テーブル$]"

            Dim dataAdapter As OleDbDataAdapter = New OleDbDataAdapter(SQL, connection)
            Dim ds As DataSet = New DataSet
            Try
                dataAdapter.Fill(ds, "PostTable")
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

            Dim table As DataTable = ds.Tables("PostTable")

            Dim row As DataRow
            For Each row In table.Rows

                line = line + 1
                If line > 32 Then ' Head Print
                    If page <> 0 Then paoRep.PageEnd()

                    page = page + 1

                    paoRep.PageStart()

                    paoRep.Write("日時", hDate)
                    paoRep.Write("ページ", "Page-" + page.ToString())

                    line = 1
                End If

                'Body Print
                paoRep.Write("郵便番号", row("郵便番号").ToString(), line)
                paoRep.Write("市区町村", row("市区町村").ToString(), line)
                paoRep.Write("住所", row("住所").ToString(), line)
                paoRep.Write("横罫線", line)

            Next
            paoRep.PageEnd()

            dataAdapter.Dispose()

        End If

        If flgPreview Then 'プレビューが選択されている場合

            ' このサンプルでは1ページのみ出力
            Dim setting As System.Drawing.Printing.PrinterSettings = New System.Drawing.Printing.PrinterSettings()
            setting.FromPage = 1
            setting.ToPage = 1

            ' プレビューを実行
            paoRep.Output(setting)

        ElseIf radPrint.Checked = True Then '印刷選択されている場合

            '印刷
            paoRep.Output()
            paoRep = Nothing

        ElseIf radPDF.Checked = True Then 'PDF出力が選択されている場合

            'PDF出力
            saveFileDialog.FileName = "郵便番号帳票"
            saveFileDialog.Filter = "PDF形式 (*.pdf)|*.pdf"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then

                'PDF出力
                paoRep.SavePDF(saveFileDialog.FileName)

                If (MessageBox.Show(Me, "PDFを表示しますか？", "PDF の表示", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                    ExecFile(saveFileDialog.FileName)
                End If

            End If

            paoRep = Nothing

        End If

    End Sub
    Private Sub ExecFile(ExecFilePath As String)
        Dim startInfo As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(ExecFilePath)
        startInfo.UseShellExecute = True
        System.Diagnostics.Process.Start(startInfo)
    End Sub

End Class
