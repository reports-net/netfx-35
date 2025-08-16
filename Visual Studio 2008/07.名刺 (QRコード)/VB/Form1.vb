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
    Friend WithEvents grid As System.Windows.Forms.DataGrid
    Friend WithEvents btnExe As System.Windows.Forms.Button
    Friend WithEvents radPrint As System.Windows.Forms.RadioButton
    Friend WithEvents radPreview As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grid = New System.Windows.Forms.DataGrid()
        Me.btnExe = New System.Windows.Forms.Button()
        Me.radPrint = New System.Windows.Forms.RadioButton()
        Me.radPreview = New System.Windows.Forms.RadioButton()
        CType(Me.grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grid
        '
        Me.grid.DataMember = ""
        Me.grid.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.grid.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grid.Location = New System.Drawing.Point(20, 53)
        Me.grid.Name = "grid"
        Me.grid.Size = New System.Drawing.Size(720, 248)
        Me.grid.TabIndex = 7
        '
        'btnExe
        '
        Me.btnExe.Font = New System.Drawing.Font("HGP創英角ﾎﾟｯﾌﾟ体", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnExe.Location = New System.Drawing.Point(20, 317)
        Me.btnExe.Name = "btnExe"
        Me.btnExe.Size = New System.Drawing.Size(720, 40)
        Me.btnExe.TabIndex = 6
        Me.btnExe.Text = "実　　行"
        '
        'radPrint
        '
        Me.radPrint.Font = New System.Drawing.Font("HG丸ｺﾞｼｯｸM-PRO", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPrint.Location = New System.Drawing.Point(396, 13)
        Me.radPrint.Name = "radPrint"
        Me.radPrint.Size = New System.Drawing.Size(93, 24)
        Me.radPrint.TabIndex = 5
        Me.radPrint.Text = "印刷"
        '
        'radPreview
        '
        Me.radPreview.Checked = True
        Me.radPreview.Font = New System.Drawing.Font("HG丸ｺﾞｼｯｸM-PRO", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.radPreview.Location = New System.Drawing.Point(236, 13)
        Me.radPreview.Name = "radPreview"
        Me.radPreview.Size = New System.Drawing.Size(136, 24)
        Me.radPreview.TabIndex = 4
        Me.radPreview.TabStop = True
        Me.radPreview.Text = "プレビュー"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(760, 374)
        Me.Controls.Add(Me.grid)
        Me.Controls.Add(Me.btnExe)
        Me.Controls.Add(Me.radPrint)
        Me.Controls.Add(Me.radPreview)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Report.NET使用例－名刺作成"
        CType(Me.grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim dataAdapter As OleDbDataAdapter
    Dim table As DataTable

    ' プログラム実行フォルダ
    Private appPath As String = Nothing
    ' Excelデータベースファイル パス
    Private DbXls As String = "名刺.xls"
    ' x64動作時加算パス(フォルダ)
    Private x64dir As String = ""

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'カレントPath取得
        appPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) & "/"
        If System.IO.File.Exists(appPath & "../../" + DbXls) = False Then
            x64dir += "../../"
            appPath += x64dir
        End If
        DbXls = appPath & "../../" + DbXls

        'サンプルの「名刺.xls」への接続 Jetエンジンを使用
        Dim connectString As String

        If IntPtr.Size = 4 Then
            connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbXls & ";Extended Properties=Excel 8.0;"
        Else
            connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DbXls & ";Extended Properties=Excel 12.0;"
        End If

        Dim connection As OleDbConnection = New OleDbConnection(connectString)

        'データセットへテーブルをセットする。
        Dim SQL As String = ""
        SQL = "select * from [名刺$]"

        dataAdapter = New OleDbDataAdapter(SQL, connection)
        Dim ds As DataSet = New DataSet
        Try
            dataAdapter.Fill(ds, "名刺")
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
        table = ds.Tables("名刺")

        grid.DataSource = table

    End Sub

    Private Sub btnExe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExe.Click


        '選択されている行の取得
        '行数の取得
        Dim n As Integer
        n = grid.BindingContext(grid.DataSource, grid.DataMember).Count
        Dim rowNo As Integer

        For rowNo = 0 To n - 1
            '行が選択されているか調べる
            If (grid.IsSelected(rowNo)) Then Exit For
        Next
        If rowNo >= n Then
            MessageBox.Show("行が選択されていません")
            Exit Sub
        End If


        'インスタンスの生成
        Dim paoRep As IReport = Nothing

        If radPreview.Checked Then
            'プレビューを選択している場合
            paoRep = ReportCreator.GetPreview()
        ElseIf radPrint.Checked Then
            '印刷の場合
            paoRep = ReportCreator.GetReport()
        End If

        paoRep.LoadDefFile(appPath & "../../名刺.prepd")

        Dim row As DataRow = table.Rows(rowNo)

        Dim name1 As String = row("名前").ToString()
        Dim kata As String = row("肩書き").ToString()
        Dim mail As String = row("メール").ToString()
        Dim tel As String = row("携帯").ToString()
        Dim name2 As String = row("携帯名前").ToString()
        Dim kana As String = row("携帯ｶﾅ").ToString()

        paoRep.PageStart()

        For line As Integer = 1 To 5

            For col As Integer = 1 To 2

                paoRep.Write("名前", name1, col, line)
                paoRep.Write("肩書き", kata, col, line)
                paoRep.Write("メール", mail, col, line)
                paoRep.Write("携帯", tel, col, line)
                paoRep.Write("QR", """MECARD:N:" & name2 & ";SOUND:" & kana & ";TEL:" & tel & ";EMAIL:" & mail & ";;""", col, line)
                paoRep.Write("a", "Pao@Office", col, line)
                paoRep.Write("b", "有限会社", col, line)
                paoRep.Write("c", "パオ･アット･オフィス", col, line)
                paoRep.Write("d", "mail:", col, line)
                paoRep.Write("e", "携帯", col, line)
                paoRep.Write("f", "http://www.pao.ac/", col, line)
                paoRep.Write("g", "本　　　社　〒275-0026　千葉県習志野市谷津3-29-2-401" & vbCrLf & "　　　　　　TEL:047-452-0057　FAX:047-452-0064", col, line)
                paoRep.Write("h", "東京事務所　〒105-0004　東京都港区新橋1-8-3 住友新橋ビル7F" & vbCrLf & "　　　　　　TEL:03-3572-6507　FAX:03-6218-0128", col, line)
            Next

        Next

        paoRep.PageEnd()

        '印刷/プレビュー
        paoRep.Output()

        dataAdapter.Dispose()

    End Sub
    Private Sub ExecFile(ExecFilePath As String)
        Dim startInfo As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(ExecFilePath)
        startInfo.UseShellExecute = True
        System.Diagnostics.Process.Start(startInfo)
    End Sub

End Class
