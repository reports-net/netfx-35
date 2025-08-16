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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 50)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(240, 136)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "　　WEBサーバ(iis.pao.ac)のWebServiceから　　印刷データをGETしてプレビューを行う"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(264, 237)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "WebService(IISと通信)"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim webTest As ac.pao.iis.WebTest = New ac.pao.iis.WebTest   'WebService呼び出し
        Dim data As Byte() = webTest.get帳票データ() '印刷データを取得

        Dim paoRep As IReport = ReportCreator.GetPreview() ' プレビュー画面を作成
        paoRep.LoadData(data) '印刷データを読み込む
        paoRep.Output() ' 印刷又はプレビューを実行

    End Sub
End Class
