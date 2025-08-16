using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Web;
using System.Web.Services;

namespace Pao.Reports
{
	/// <summary>
	/// Service1 の概要の説明です。
	/// </summary>
	public class WebTest : System.Web.Services.WebService
	{
		public WebTest()
		{
			//CODEGEN: この呼び出しは、ASP.NET Web サービス デザイナで必要です。
			InitializeComponent();
		}

		#region コンポーネント デザイナで生成されたコード 
		
		//Web サービス デザイナで必要です。
		private IContainer components = null;
				
		/// <summary>
		/// デザイナ サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディタで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
		}

		/// <summary>
		/// 使用されているリソースに後処理を実行します。
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if(disposing && components != null)
			{
				components.Dispose();
			}
			base.Dispose(disposing);		
		}
		
		#endregion

		[WebMethod]
		public byte[] get帳票データ()
		{
			//インスタンスの生成
			IReport paoRep = ReportCreator.GetReport();
            
			//帳票定義体の読み込み
			paoRep.LoadDefFile(@"C:\PrintDefine\PaoRep1.prepd");

			//帳票編集
			paoRep.PageStart();
			paoRep.Write("Square1");
			paoRep.Write("Square2");
			paoRep.Write("Circle1");
			paoRep.Write("Circle2");
			paoRep.Write("Line1");
			paoRep.Write("Line2");
			paoRep.Write("Barcode1", "123456789012");
			paoRep.Write("Barcode2", "123456789012");
			paoRep.Write("Barcode3", "123456789012");
			paoRep.Write("Barcode4", "123456789012");
			paoRep.Write("Barcode5", "123456789012");
			paoRep.Write("Barcode6", "123456789012");
			paoRep.Write("Barcode7", "123456789012");
			paoRep.Write("Text1", "文字列");
			paoRep.Write("Text2", "これはWEBサーバー\n(iis.pao.ac)で作った\n印刷データですよ～ん♪");
			paoRep.PageEnd();

			return paoRep.SaveData(); // 印刷データを保存 & 復帰
		}
	}
}
