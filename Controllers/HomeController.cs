using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using WordToImage.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO.Compression;

namespace WordToImage.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult ConvertAllPages()
        {
            using(FileStream docStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using(WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    using(DocIORenderer renderer = new DocIORenderer())
                    {
                        Stream[] imageStreams = wordDocument.RenderAsImages();
                        using(MemoryStream ms = new MemoryStream())
                        {
                            using(var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                            {
                                int i = 1;
                                foreach(Stream imageStream in imageStreams)
                                {
                                    imageStream.Position = 0;
                                    var image = zip.CreateEntry("WordToImage_" + i + ".jpeg");
                                    using(var entryStream = image.Open())
                                    {
                                        imageStream.CopyTo(entryStream);
                                    }
                                    i++;
                                }
                            }
                            return File(ms.ToArray(), "application/zip", "WordToImage.zip");
                        }
                    }
                }
            }
            return View();
        }

        public IActionResult RangeOfPages()
        {
            using (FileStream docStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        Stream[] imageStreams = wordDocument.RenderAsImages(1,2);
                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                            {
                                int i = 1;
                                foreach (Stream imageStream in imageStreams)
                                {
                                    imageStream.Position = 0;
                                    var image = zip.CreateEntry("WordToImage_" + i + ".jpeg");
                                    using (var entryStream = image.Open())
                                    {
                                        imageStream.CopyTo(entryStream);
                                    }
                                    i++;
                                }
                            }
                            return File(ms.ToArray(), "application/zip", "WordToImage.zip");
                        }
                    }
                }
            }
            return View();
        }

        public IActionResult ThumbnailImage()
        {
            using (FileStream docStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
                        return File(imageStream, "image/jpeg", "WordToImage_1.jpeg");
                    }
                }
            }
            return View();
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}