using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using SelectPdf;

namespace HtmlToPDF.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
	{
		//获取项目路径使用
		IHostingEnvironment hostingEnv;
		public ValuesController(IHostingEnvironment hostingEnv)
		{
			this.hostingEnv = hostingEnv;
		}
		
		/// <summary>
		/// Html导出PDF
		/// </summary>
		/// <returns></returns>
		[HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
			//测试Html素材
            string htmlString = @"<!DOCTYPE html>
						<html>
						<head>
						    <meta charset='utf-8' />
						    <title></title>
						     <style>
						       body {
						         margin: 0;
						         padding: 0;
						         min-width: 2000px;
						      }
						      .m_table {
						         border-collapse: separate;
						         margin: 100px auto 0;
						         min-width: 1900px;
						         text-align: center;
						         font: 500 17px '微软雅黑';
						         border-spacing: 0;
							     border: 1px solid #EBEEF5;
						      }
						
						      .m_table th {
						         background-color: #F7F3F7;
						      }
						
						      .m_table th,
						      .m_table td {
						         border-right: 1px solid #EBEEF5;
								 border-bottom: 1px solid #EBEEF5; 
						         padding: 5px;
						         height: 60px;
						         width: 60px;
						      }
						   </style>
						</head>
						
						<body>
						    
						    <table class='m_table'> 
						        <tr>
						            <th colspan='9' style='text-align:center;font-size:28px;'>期末学生综合评价表1</th>
						        </tr>
								<tr>
						            <th colspan='9' style='text-align:center;font-size:23px;'>班级：一年级一班  姓名：测试  班主任：测试</th> 
						        </tr>
						        <tr>
						            <th  rowspan='2' style='text-align:center;font-size:24px;'>学科</th>
						            <th  colspan='2'  rowspan='2' style='text-align:center;font-size:24px;'>过程性评价</th>
						            <th colspan='2' rowspan='2' style='text-align:center;font-size:24px;'>表现性评价</th>
						            <th colspan='2'  style='text-align:center;font-size:20px;'>考试性评价</th>
						            <th colspan='2' style='text-align:center;font-size:20px;'>综合性评价</th>
						        </tr>
						        <tr> 
						            <th   style='text-align:center;font-size:18px;'>卷面分</th>
						            <th  style='text-align:center;font-size:18px;'>权重分</th>
						            <th   style='text-align:center;font-size:18px;'>总分</th>
						            <th  style='text-align:center;font-size:18px;'>等级</th>
						        </tr>
						        <tr>
								<td> 语文 </td>
								<td> 作业评价/4 </td>
								<td> 课堂表现/5 </td>
								<td> 学科必选/9 </td>
								<td> 学科自选/9.6 </td>
								<td> 99 </td>
								<td> 69.3 </td>
								<td> 96.9 </td>
								<td> A </td>
								</tr>
								
								<tr><td> 数学 </td><td> 作业评价/4 </td><td> 课堂表现/4 </td><td> 学科必选/10 </td><td> 学科自选/7 </td><td> 98 </td><td> 68.6 </td><td> 93.6 </td><td> A </td></tr><tr><td> 英语 </td><td> 作业评价/4.5 </td><td> 课堂表现/5 </td><td> 学科必选/1.7 </td><td> 学科自选/1.7 </td><td> 59 </td><td> 41.3 </td><td> 54.2 </td><td> D </td></tr><tr><td> 英语 </td><td> 作业评价/4.5 </td><td> 课堂表现/5 </td><td> 学科必选/1.7 </td><td> 学科自选/1.7 </td><td> 59 </td><td> 41.3 </td><td> 54.2 </td><td> D </td></tr><tr><td> 英语 </td><td> 作业评价/4.5 </td><td> 课堂表现/5 </td><td> 学科必选/1.7 </td><td> 学科自选/1.7 </td><td> 59 </td><td> 41.3 </td><td> 54.2 </td><td> D </td></tr><tr><td> 英语 </td><td> 作业评价/4.5 </td><td> 课堂表现/5 </td><td> 学科必选/1.7 </td><td> 学科自选/1.7 </td><td> 59 </td><td> 41.3 </td><td> 54.2 </td><td> D </td></tr>
						    </table>  
						</body>
						</html>";
			HtmlToPdf Renderer = new HtmlToPdf();
			//设置Pdf参数
			Renderer.Options.PdfPageOrientation = PdfPageOrientation.Landscape;//设置页面方式-横向  PdfPageOrientation.Portrait  竖向
			Renderer.Options.PdfPageSize = PdfPageSize.A4;//设置页面大小，30种页面大小可以选择
			Renderer.Options.MarginTop = 10;   //上下左右边距设置  
			Renderer.Options.MarginBottom = 10;
			Renderer.Options.MarginLeft = 10;
			Renderer.Options.MarginRight = 10;
			 
			//设置更多额参数可以去HtmlToPdfOptions里面选择设置
			var docHtml = Renderer.ConvertHtmlString(htmlString);//根据html内容导出PDF
			var docUrl = Renderer.ConvertUrl("https://fanyi.baidu.com/#en/zh/");//根据url路径导出PDF
			string webRootPath = hostingEnv.ContentRootPath;  //获取项目运行绝对路径
			var path = $"/ExportPDF/{DateTime.Now.ToString("yyyyMMdd")}/";//文件相对路径
			var savepathHtml = $"{webRootPath}{path}{Guid.NewGuid().ToString()}-Html.pdf";//保存绝对路径
			if (!Directory.Exists(Path.GetDirectoryName(webRootPath + path)))
			{
				Directory.CreateDirectory(Path.GetDirectoryName(webRootPath + path));
			}
			docHtml.Save(savepathHtml);
			var savepathUrl = $"{webRootPath}{path}{Guid.NewGuid().ToString()}-Url.pdf";//保存绝对路径
			docUrl.Save(savepathUrl);


			return new string[] { savepathHtml, savepathUrl };
        }

		/// <summary>
		/// Html导出PDF一个文件多页
		/// </summary>
		/// <param name="PageSize"></param>
		/// <returns></returns>
		[HttpGet("HtmlToPdfList")]
		public ActionResult<string> HtmlToPdfList(int PageSize = 1)
		{
			//测试Html素材
			string htmlString = @"<!DOCTYPE html>
						<html>
						<head>
						    <meta charset='utf-8' />
						    <title></title>
						     <style>
						       body {
						         margin: 0;
						         padding: 0;
						         min-width: 2000px;
						      }
						      .m_table {
						         border-collapse: separate;
						         margin: 100px auto 0;
						         min-width: 1900px;
						         text-align: center;
						         font: 500 17px '微软雅黑';
						         border-spacing: 0;
							     border: 1px solid #EBEEF5;
						      }
						
						      .m_table th {
						         background-color: #F7F3F7;
						      }
						
						      .m_table th,
						      .m_table td {
						         border-right: 1px solid #EBEEF5;
								 border-bottom: 1px solid #EBEEF5; 
						         padding: 5px;
						         height: 60px;
						         width: 60px;
						      }
						   </style>
						</head>
						
						<body>
						    
						    <table class='m_table'> 
						        <tr>
						            <th colspan='9' style='text-align:center;font-size:28px;'>期末学生综合评价表1</th>
						        </tr>
								<tr>
						            <th colspan='9' style='text-align:center;font-size:23px;'>班级：一年级一班  姓名：测试  班主任：测试</th> 
						        </tr>
						        <tr>
						            <th  rowspan='2' style='text-align:center;font-size:24px;'>学科</th>
						            <th  colspan='2'  rowspan='2' style='text-align:center;font-size:24px;'>过程性评价</th>
						            <th colspan='2' rowspan='2' style='text-align:center;font-size:24px;'>表现性评价</th>
						            <th colspan='2'  style='text-align:center;font-size:20px;'>考试性评价</th>
						            <th colspan='2' style='text-align:center;font-size:20px;'>综合性评价</th>
						        </tr>
						        <tr> 
						            <th   style='text-align:center;font-size:18px;'>卷面分</th>
						            <th  style='text-align:center;font-size:18px;'>权重分</th>
						            <th   style='text-align:center;font-size:18px;'>总分</th>
						            <th  style='text-align:center;font-size:18px;'>等级</th>
						        </tr>
						        <tr>
								<td> 语文 </td>
								<td> 作业评价/4 </td>
								<td> 课堂表现/5 </td>
								<td> 学科必选/9 </td>
								<td> 学科自选/9.6 </td>
								<td> 99 </td>
								<td> 69.3 </td>
								<td> 96.9 </td>
								<td> A </td>
								</tr>
								
								<tr><td> 数学 </td><td> 作业评价/4 </td><td> 课堂表现/4 </td><td> 学科必选/10 </td><td> 学科自选/7 </td><td> 98 </td><td> 68.6 </td><td> 93.6 </td><td> A </td></tr><tr><td> 英语 </td><td> 作业评价/4.5 </td><td> 课堂表现/5 </td><td> 学科必选/1.7 </td><td> 学科自选/1.7 </td><td> 59 </td><td> 41.3 </td><td> 54.2 </td><td> D </td></tr><tr><td> 英语 </td><td> 作业评价/4.5 </td><td> 课堂表现/5 </td><td> 学科必选/1.7 </td><td> 学科自选/1.7 </td><td> 59 </td><td> 41.3 </td><td> 54.2 </td><td> D </td></tr><tr><td> 英语 </td><td> 作业评价/4.5 </td><td> 课堂表现/5 </td><td> 学科必选/1.7 </td><td> 学科自选/1.7 </td><td> 59 </td><td> 41.3 </td><td> 54.2 </td><td> D </td></tr><tr><td> 英语 </td><td> 作业评价/4.5 </td><td> 课堂表现/5 </td><td> 学科必选/1.7 </td><td> 学科自选/1.7 </td><td> 59 </td><td> 41.3 </td><td> 54.2 </td><td> D </td></tr>
						    </table>  
						</body>
						</html>";
			PdfDocument docHtml = null;
			for (int j = 0; j < PageSize; j++)
			{
				HtmlToPdf Renderer = new HtmlToPdf();
				//设置Pdf参数
				Renderer.Options.PdfPageOrientation = PdfPageOrientation.Landscape;//设置页面方式-横向  PdfPageOrientation.Portrait  竖向
				Renderer.Options.PdfPageSize = PdfPageSize.A4;//设置页面大小，30种页面大小可以选择
				Renderer.Options.MarginTop = 10;   //上下左右边距设置  
				Renderer.Options.MarginBottom = 10;
				Renderer.Options.MarginLeft = 10;
				Renderer.Options.MarginRight = 10;
				//设置更多额参数可以去HtmlToPdfOptions里面选择设置

				if (docHtml == null)
					docHtml = Renderer.ConvertHtmlString(htmlString);//根据html内容导出PDF 
				else
					//在上一个pdf元素页面下面追加Pdf页面，官方文档对于一个pdf文件打印多页的处理提供了分页符，在你想打印一页的元素外面加上   <div style="font-size: 28px; page-break-after: always">元素
					//也就是分页符，但是试用感觉效果并不理想，下面这个Append追加一个pdf页面效果会更好点，但是可能会损耗一些性能
					docHtml.Append(Renderer.ConvertHtmlString(htmlString));
			}
			string webRootPath = hostingEnv.ContentRootPath;  //获取项目运行绝对路径
			var path = $"/ExportPDF/{DateTime.Now.ToString("yyyyMMdd")}/";//文件相对路径
			var savepathHtml = $"{webRootPath}{path}{Guid.NewGuid().ToString()}-Html.pdf";//保存绝对路径
			if (!Directory.Exists(Path.GetDirectoryName(webRootPath + path)))
			{
				Directory.CreateDirectory(Path.GetDirectoryName(webRootPath + path));
			}
			docHtml.Save(savepathHtml);
			return savepathHtml;
		}
    }
}
