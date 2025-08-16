Imports System.Web.Services
Imports Pao.Reports


<System.Web.Services.WebService(Namespace:="http://tempuri.org/Pao.Reports/WebTest")> _
Public Class WebTest
    Inherits System.Web.Services.WebService

#Region " Web サービス デザイナで生成されたコード "

    Public Sub New()
        MyBase.New()

        'この呼び出しは Web サービス デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後に独自の初期化コードを追加してください。

    End Sub

    'Web サービス デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ : 以下のプロシージャは、Web サービス デザイナで必要です。
    'Web サービス デザイナを使って変更することができます。  
    'コード エディタによる変更は行わないでください。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: このプロシージャは Web サービス デザイナで必要です。
        'コード エディタによる変更は行わないでください。
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

    <WebMethod()> _
    Public Function get帳票データ() As Byte()

        'インスタンスの生成
        Dim paoRep As IReport = ReportCreator.GetReport()

        '帳票定義体の読み込み
        paoRep.LoadDefFile("C:\PrintDefine\PaoRep1.prepd")

        '帳票編集
        paoRep.PageStart()
        paoRep.Write("Square1")
        paoRep.Write("Square2")
        paoRep.Write("Circle1")
        paoRep.Write("Circle2")
        paoRep.Write("Line1")
        paoRep.Write("Line2")
        paoRep.Write("Barcode1", "123456789012")
        paoRep.Write("Barcode2", "123456789012")
        paoRep.Write("Barcode3", "123456789012")
        paoRep.Write("Barcode4", "123456789012")
        paoRep.Write("Barcode5", "123456789012")
        paoRep.Write("Barcode6", "123456789012")
        paoRep.Write("Barcode7", "123456789012")
        paoRep.Write("Text1", "文字列")
        paoRep.Write("Text2", "これはWEBサーバー\n(iis.pao.ac)で作った\n印刷データですよ～ん♪")
        paoRep.PageEnd()

        Return paoRep.SaveData() ' 印刷データを保存 & 復帰

    End Function

End Class
