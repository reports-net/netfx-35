Imports System.IO
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
    Friend WithEvents opt5 As System.Windows.Forms.RadioButton
    Friend WithEvents opt4 As System.Windows.Forms.RadioButton
    Friend WithEvents opt2 As System.Windows.Forms.RadioButton
    Friend WithEvents opt3 As System.Windows.Forms.RadioButton
    Friend WithEvents opt1 As System.Windows.Forms.RadioButton
    Friend WithEvents button2 As System.Windows.Forms.Button
    Friend WithEvents button1 As System.Windows.Forms.Button
    Private WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.opt5 = New System.Windows.Forms.RadioButton
        Me.opt4 = New System.Windows.Forms.RadioButton
        Me.opt2 = New System.Windows.Forms.RadioButton
        Me.opt3 = New System.Windows.Forms.RadioButton
        Me.opt1 = New System.Windows.Forms.RadioButton
        Me.button2 = New System.Windows.Forms.Button
        Me.button1 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.txtMessage = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'opt5
        '
        Me.opt5.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.opt5.Location = New System.Drawing.Point(792, 224)
        Me.opt5.Name = "opt5"
        Me.opt5.Size = New System.Drawing.Size(299, 48)
        Me.opt5.TabIndex = 13
        Me.opt5.Text = "広告(MySQL 使用)"
        '
        'opt4
        '
        Me.opt4.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.opt4.Location = New System.Drawing.Point(455, 224)
        Me.opt4.Name = "opt4"
        Me.opt4.Size = New System.Drawing.Size(300, 48)
        Me.opt4.TabIndex = 12
        Me.opt4.Text = "見積書(MySQL 使用)"
        '
        'opt2
        '
        Me.opt2.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.opt2.Location = New System.Drawing.Point(792, 128)
        Me.opt2.Name = "opt2"
        Me.opt2.Size = New System.Drawing.Size(299, 48)
        Me.opt2.TabIndex = 11
        Me.opt2.Text = "10の倍数出力"
        '
        'opt3
        '
        Me.opt3.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.opt3.Location = New System.Drawing.Point(1126, 128)
        Me.opt3.Name = "opt3"
        Me.opt3.Size = New System.Drawing.Size(346, 48)
        Me.opt3.TabIndex = 10
        Me.opt3.Text = "住所一覧(MySQL 使用)"
        '
        'opt1
        '
        Me.opt1.Checked = True
        Me.opt1.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.opt1.Location = New System.Drawing.Point(458, 128)
        Me.opt1.Name = "opt1"
        Me.opt1.Size = New System.Drawing.Size(246, 48)
        Me.opt1.TabIndex = 9
        Me.opt1.TabStop = True
        Me.opt1.Text = "単純な印刷データ"
        '
        'button2
        '
        Me.button2.Font = New System.Drawing.Font("MS UI Gothic", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.button2.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.button2.Location = New System.Drawing.Point(792, 400)
        Me.button2.Name = "button2"
        Me.button2.Size = New System.Drawing.Size(440, 144)
        Me.button2.TabIndex = 8
        Me.button2.Text = "印刷"
        '
        'button1
        '
        Me.button1.Font = New System.Drawing.Font("MS UI Gothic", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.button1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.button1.Location = New System.Drawing.Point(141, 400)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(440, 144)
        Me.button1.TabIndex = 7
        Me.button1.Text = "プレビュー"
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("MS UI Gothic", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Button3.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Button3.Location = New System.Drawing.Point(1373, 400)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(440, 144)
        Me.Button3.TabIndex = 14
        Me.Button3.Text = "PDF出力"
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(0, 59)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(1894, 500)
        Me.txtMessage.TabIndex = 15
        Me.txtMessage.Text = resources.GetString("txtMessage.Text")
        Me.txtMessage.Visible = False
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(11, 24)
        Me.ClientSize = New System.Drawing.Size(1895, 637)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.opt5)
        Me.Controls.Add(Me.opt4)
        Me.Controls.Add(Me.opt2)
        Me.Controls.Add(Me.opt3)
        Me.Controls.Add(Me.opt1)
        Me.Controls.Add(Me.button2)
        Me.Controls.Add(Me.button1)
        Me.Name = "Form1"
        Me.Text = "WEBサーバ(www.pao.ac)の Axis WebService から Reports.jar で作成した印刷データをGETして印刷・プレビューを行うサンプ" & _
            "ル"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    'プレビューボタン
    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        txtMessage.Visible = True
        Exit Sub


        Dim data As Byte() = Nothing

        Dim webTest As ac.pao.www.SampleService = New ac.pao.www.SampleService   'WebService呼び出し

        If (opt1.Checked) Then data = webTest.getPrintData() '単純な印刷データを取得
        If (opt2.Checked) Then data = webTest.getBaisuu() '10の倍数サンプル 印刷データを取得
        If (opt3.Checked) Then data = webTest.getAddressList() '住所一覧サンプル 印刷データを取得
        If (opt4.Checked) Then data = webTest.getMitsumori() '見積書サンプル 印刷データを取得
        If (opt5.Checked) Then data = webTest.getKoukoku() '広告サンプル 印刷データを取得

        Dim paoRep As IReport = ReportCreator.GetPreview() ' プレビューオブジェクトを作成
        paoRep.LoadData(data) '印刷データを読み込む
        paoRep.Output() ' プレビューを実行

    End Sub

    '印刷ボタン
    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        txtMessage.Visible = True
        Exit Sub

        Dim data As Byte() = Nothing

        Dim webTest As ac.pao.www.SampleService = New ac.pao.www.SampleService   'WebService呼び出し

        If (opt1.Checked) Then data = webTest.getPrintData() '単純な印刷データを取得
        If (opt2.Checked) Then data = webTest.getBaisuu() '10の倍数サンプル 印刷データを取得
        If (opt3.Checked) Then data = webTest.getAddressList() '住所一覧サンプル 印刷データを取得
        If (opt4.Checked) Then data = webTest.getMitsumori() '見積書サンプル 印刷データを取得
        If (opt5.Checked) Then data = webTest.getKoukoku() '広告サンプル 印刷データを取得

        Dim paoRep As IReport = ReportCreator.GetReport() ' 印刷オブジェクトを作成
        paoRep.LoadData(data) '印刷データを読み込む
        paoRep.Output() ' 印刷を実行

    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        txtMessage.Visible = True
        Exit Sub

        If (opt5.Checked) Then
            MessageBox.Show("広告サンプルのPDF作成WEBサービスは、制限により出力できません。")
            Return
        End If

        Dim data As Byte() = Nothing

        Dim webTest As pdf.ac.pao.www.SamplePdfService = New pdf.ac.pao.www.SamplePdfService   'WebService呼び出し

        If (opt1.Checked) Then data = webTest.getPrintData() '単純な印刷データを取得
        If (opt2.Checked) Then data = webTest.getBaisuu() '10の倍数サンプル 印刷データを取得
        If (opt3.Checked) Then data = webTest.getAddressList() '住所一覧サンプル 印刷データを取得
        If (opt4.Checked) Then data = webTest.getMitsumori() '見積書サンプル 印刷データを取得


        'PDF出力
        MessageBox.Show(Me, "PDFファイルを「Sample.PDF」という名前でデスクトップに出力します。")
        'デスクトップのパス取得
        Dim DeskTop As String = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)

        Dim stream As FileStream = New FileStream(Path.Combine(DeskTop, "Sample.PDF"), FileMode.Create, FileAccess.Write)
        stream.Write(data, 0, data.Length)
        stream.Close()

        System.Diagnostics.Process.Start(DeskTop + "\Sample.PDF")

    End Sub
End Class
