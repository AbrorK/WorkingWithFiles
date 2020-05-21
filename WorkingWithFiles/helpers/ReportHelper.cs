using System;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Hosting;
using QRCoder;
using WorkingWithFiles.models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Path = System.IO.Path;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace WorkingWithFiles.helpers
{
    public static class ReportHelper
    {
        public static void DownloadDeal(this LogisTicketItem item, string templatePath, string savePath, string tokenUrl, IHostingEnvironment environment)
        {
            File.Copy(templatePath, savePath, true);

            var fields = typeof(LogisTicketItem).GetProperties().ToDictionary(a => a.Name.Replace("_", ""), a =>
            {
                var obj = a.GetValue(item, null);
                var result = "";

                switch (Type.GetTypeCode(a.PropertyType))
                {
                    case TypeCode.DateTime:
                        result = Convert.ToDateTime(obj).ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                        break;
                    case TypeCode.Decimal:
                    case TypeCode.Double:
                    case TypeCode.Single:
                        result = AsFormat((decimal?)obj);
                        break;
                    default:
                        result = EscapeXml($"{obj}");
                        break;
                }
                return result;
            });

            using (var wordDoc = WordprocessingDocument.Open(savePath, true))
            {
                var mainPart = wordDoc.MainDocumentPart;

                //replacer
                var docText = "";
                using (var sr = new StreamReader(mainPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }
                foreach (var field in fields.OrderByDescending(a => a.Key))
                {
                    docText = Regex.Replace(docText, field.Key, field.Value, RegexOptions.IgnoreCase);
                }

                //save
                using (var sw = new StreamWriter(mainPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }

                //image
                var pictureElem = mainPart.Document.Body.Descendants<Text>().Where(x => x.Text == "qrcode").FirstOrDefault();
                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var stream = new MemoryStream())
                {
                    //qr generator
                    var qrGenerator = new QRCodeGenerator();
                    var qrCodeData = qrGenerator.CreateQrCode(tokenUrl, QRCodeGenerator.ECCLevel.Q);
                    var qrCode = new QRCode(qrCodeData);
                    var logoPath = Path.Combine(environment.WebRootPath, "logo.png");
                    var qrCodeImage = qrCode.GetGraphic(20, System.Drawing.Color.Black, System.Drawing.Color.White, (System.Drawing.Bitmap)System.Drawing.Image.FromFile(logoPath), 25);

                    //save to stream
                    qrCodeImage.Save(stream, ImageFormat.Png);

                    //back to start position
                    stream.Position = 0;

                    //pass to doc
                    imagePart.FeedData(stream);
                }
                AddImage(pictureElem, mainPart.GetIdOfPart(imagePart));
                pictureElem.Remove();

                wordDoc.Save();
            }

        }

        private static void AddImage(Text cell, string relationshipId)
        {
            var element =
              new Drawing(
                new DW.Inline(
                  new DW.Extent() { Cx = 990000L, Cy = 792000L },
                  new DW.EffectExtent()
                  {
                      LeftEdge = 0L,
                      TopEdge = 0L,
                      RightEdge = 0L,
                      BottomEdge = 0L
                  },
                  new DW.DocProperties()
                  {
                      Id = 1U,
                      Name = "Picture 1"
                  },
                  new DW.NonVisualGraphicFrameDrawingProperties(
                      new A.GraphicFrameLocks() { NoChangeAspect = true }),
                  new A.Graphic(
                    new A.GraphicData(
                      new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                          new PIC.NonVisualDrawingProperties()
                          {
                              Id = 0U,
                              Name = "New Bitmap Image.jpg"
                          },
                          new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                          new A.Blip(
                            new A.BlipExtensionList(
                              new A.BlipExtension()
                              {
                                  Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                              })
                           )
                          {
                              Embed = relationshipId,
                              CompressionState =
                              A.BlipCompressionValues.Print
                          },
                          new A.Stretch(
                            new A.FillRectangle())),
                          new PIC.ShapeProperties(
                            new A.Transform2D(
                              new A.Offset() { X = 0L, Y = 0L },
                              new A.Extents() { Cx = 990000L, Cy = 792000L }),
                            new A.PresetGeometry(
                              new A.AdjustValueList()
                            )
                            { Preset = A.ShapeTypeValues.Rectangle }))
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                });

            cell.InsertAfterSelf(new Paragraph(new Run(element)));
        }

        private static string AsFormat(this decimal? d, string currency = "")
        {
            if (d == null)
                return "";

            var nfi = new CultureInfo("en-US", false).NumberFormat;

            nfi.CurrencyDecimalSeparator = ",";
            nfi.CurrencyGroupSeparator = " ";
            nfi.CurrencySymbol = currency;

            return d.Value.ToString("C2", nfi);
        }

        private static string EscapeXml(string s)
        {
            return s.Replace("&", "").Replace("<", "").Replace(">", "").Replace("\"", "").Replace("'", "");
        }

    }
}
