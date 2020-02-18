using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using estagioprobatorioback.Models;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using System.Xml.XPath;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace estagioprobatorioback.Reports
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportAvaliacaoFinal
    {
        Avaliacao_Final Model { get;set;}
        PdfDocument document { get; set; }
        TextFrame addressFrame { get; set; }
        MigraDoc.DocumentObjectModel.Tables.Table table { get; set; }

        public ReportAvaliacaoFinal(Avaliacao_Final model)
        {
            this.Model = model;
        }

        public FileStreamResult CreateDocument()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            MemoryStream stream = new MemoryStream();

            // Create a new MigraDoc document
            this.document = new PdfDocument();
            this.document.Info.Title = "Relatório Final";
            this.document.Info.Subject = "Relatório Final";
            this.document.Info.Author = "Wanialdo Lima";          
            document.Info.Keywords = "PDFsharp, XGraphics";
 
            //Gerar o Relatório de Avaliação
            ReportPage1(document);
            //SamplePage2(document);

            document.Save(stream, false);
           
            return new FileStreamResult(stream, "Application/pdf");
        }

        //Se houverem dúvidas, há duas páginas de exemplo oficial do PDFSHarp após esse método.
        protected void ReportPage1(PdfDocument document)
        {
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            gfx.MUH = PdfFontEncoding.Unicode;

            // You always need a MigraDoc document for rendering.
            Document doc = new Document();

            Section sec = doc.AddSection();

            Paragraph image = sec.AddParagraph();
            image.Format.Alignment = ParagraphAlignment.Right;
            Image imagem = image.AddImage("logo.png");
            imagem.Height = "1.5cm";
            imagem.LockAspectRatio = true;

            Paragraph para1 = sec.AddParagraph();
            para1.Format.Alignment = ParagraphAlignment.Center;
            para1.Format.Font.Name = "Arial";
            para1.Format.Font.Size = 14;
            para1.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            para1.AddFormattedText("RELATÓRIO FINAL DA AVALIAÇÃO DE EXEMPLO DE USO DO MIGRADOC", TextFormat.Bold);
            para1.Format.Borders.Distance = "5pt";

            // Tabela da Fase I
            Table table1 = sec.AddTable();
            table1.Style = "Table";
            table1.Format.Font.Size = 9;
            table1.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(100, 100, 100);
            table1.Borders.Width = 0.5;
            table1.Borders.Left.Width = 0.5;
            table1.Borders.Right.Width = 0.5;
            table1.Rows.LeftIndent = 0;
            table1.BottomPadding = Unit.FromCentimeter(0.1);
            table1.TopPadding = Unit.FromCentimeter(0.1);
            table1.LeftPadding = Unit.FromCentimeter(0.2);
            table1.RightPadding = Unit.FromCentimeter(0.2);
            table1.Rows.Alignment = RowAlignment.Left;
            table1.Rows.VerticalAlignment = VerticalAlignment.Top;

            // Before you can add a row, you must define the columns
            Column column = table1.AddColumn("4.25cm");
            column = table1.AddColumn("4.25cm");
            column = table1.AddColumn("4.25cm");
            column = table1.AddColumn("4.25cm");

            // Create the header of the table
            Row row = table1.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row.Cells[0].AddParagraph("I - IDENTIFICAÇÃO DO FUNCIONÁRIO");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].MergeRight = 3;

            row = table1.AddRow();
            row.Cells[0].AddParagraph("Nome");
            row.Cells[0].AddParagraph(this.Model.Nome);
            row.Cells[0].MergeRight = 1;
            row.Cells[2].AddParagraph("Cargo");
            row.Cells[2].AddParagraph(this.Model.Cargo);
            row.Cells[3].AddParagraph("Matrícula");
            row.Cells[3].AddParagraph(this.Model.Matricula);

            row = table1.AddRow();
            row.Cells[0].AddParagraph("Filial");
            row.Cells[0].AddParagraph(this.Model.Filial);
            row.Cells[0].MergeRight = 1;
            row.Cells[2].AddParagraph("Setor");
            row.Cells[2].AddParagraph(this.Model.Setor);
            row.Cells[3].AddParagraph("Data Admissão");
            row.Cells[3].AddParagraph(this.Model.DataAdmissao);

            row = table1.AddRow();
            row.Cells[0].AddParagraph(" ");
            row.Cells[0].MergeRight = 3;
            row.Borders.Visible = false;

            row = table1.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row.Cells[0].AddParagraph("II - RESULTADO FINAL");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].MergeRight = 3;

            row = table1.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row.Cells[0].AddParagraph("ETAPA");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].MergeDown = 1;
            row.Cells[1].AddParagraph("PERÍODO");
            row.Cells[1].Format.Font.Bold = false;
            row.Cells[1].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[1].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[1].MergeRight = 1;
            row.Cells[3].AddParagraph("RESULTADO DA AVALIAÇÃO (RA)");
            row.Cells[3].Format.Font.Bold = false;
            row.Cells[3].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[3].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[3].MergeDown = 1;

            row = table1.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row.Cells[1].AddParagraph("DATA INÍCIO");
            row.Cells[1].Format.Font.Bold = false;
            row.Cells[1].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[1].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[2].AddParagraph("DATA FIM");
            row.Cells[2].Format.Font.Bold = false;
            row.Cells[2].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].VerticalAlignment = VerticalAlignment.Center;

            foreach (var item in this.Model.ResultadosAvaliacao) 
            {
                row = table1.AddRow();
                row.Cells[0].AddParagraph(item.NomeEtapa);
                row.Cells[1].AddParagraph(item.DataInicioEtapa);
                row.Cells[1].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[2].AddParagraph(item.DataFimEtapa);
                row.Cells[2].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[3].AddParagraph(item.ResultadoEtapa);
                row.Cells[3].Format.Alignment = ParagraphAlignment.Right;
            }

            row = table1.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row.Cells[0].AddParagraph("RESULTADO FINAL = [(RAE1) + (RAE2) + (RAE3) + (RAE4)]/4");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].MergeRight = 2;
            row.Cells[3].AddParagraph(this.Model.ResultadoFinal);
            row.Cells[3].Format.Font.Bold = false;
            row.Cells[3].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[3].VerticalAlignment = VerticalAlignment.Center;

            row = table1.AddRow();
            row.Cells[0].AddParagraph(" ");
            row.Cells[0].MergeRight = 3;
            row.Borders.Visible = false;

            row = table1.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row.Cells[0].AddParagraph("III - CONCLUSÃO");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].MergeRight = 3;

            row = table1.AddRow();
            row.Format.Alignment = ParagraphAlignment.Justify;
            row.Cells[0].AddParagraph("ipit iurero dolum zzriliquisis nit wis dolore vel et nonsequipit, velendigna "+
                "auguercilit lor se dipisl duismod tatem zzrit at laore magna feummod oloborting ea con vel "+
                "essit augiati onsequat luptat nos diatum vel ullum illummy nonsent nit ipis et nonsequis "+
                "essent augait el ing eumsan hendre feugait prat augiatem amconul laoreet.");
            row.Cells[0].AddParagraph(" ");
            row.Cells[0].AddParagraph((this.Model.Conclusao ? "APTO" : "INAPTO") + " a ser efetivado no serviço público.");
            row.Cells[0].MergeRight = 3;

            row = table1.AddRow();
            row.Cells[0].AddParagraph(" ");
            row.Cells[0].MergeRight = 3;
            row.Borders.Visible = false;

            row = table1.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row.Cells[0].AddParagraph("IV - MEMBROS AVALIADORES");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].MergeRight = 3;

            row = table1.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row.Cells[0].AddParagraph("NOME");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].MergeRight = 1;
            row.Cells[2].AddParagraph("MATRÍCULA");
            row.Cells[2].Format.Font.Bold = false;
            row.Cells[2].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[3].AddParagraph("ASSINATURA");
            row.Cells[3].Format.Font.Bold = false;
            row.Cells[3].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[3].VerticalAlignment = VerticalAlignment.Center;

            foreach(var item in this.Model.MembrosComissao) {
                row = table1.AddRow();
                row.Cells[0].AddParagraph(item.NomeMembroComissao);
                row.Cells[0].MergeRight = 1;
                row.Cells[2].AddParagraph(item.MatriculaMembroComissao);
                row.Cells[2].Format.Alignment = ParagraphAlignment.Center;
                row.Cells[3].AddParagraph(" ");
            }

            row = table1.AddRow();
            row.Cells[0].AddParagraph(" ");
            row.Cells[0].MergeRight = 3;
            row.Borders.Visible = false;

            row = table1.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row.Cells[0].AddParagraph("V - DE ACORDO GERENTE");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].MergeRight = 3;

            row = table1.AddRow();
            row.Cells[0].AddParagraph("DE ACORDO.");
            row.Cells[0].AddParagraph("AO RH, PARA PROVIDÊNCIAS.");
            row.Cells[0].AddParagraph(" ");
            row.Cells[0].Format.Borders.Bottom.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(255, 255, 255);
            row.Cells[0].MergeRight = 3;

            row = table1.AddRow();
            row.Cells[0].AddParagraph("_______________________________");
            row.Cells[0].AddParagraph("Data");
            row.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].Format.Borders.Top.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(255, 255, 255);
            row.Cells[0].Format.Borders.Right.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(255, 255, 255);
            row.Cells[0].MergeRight = 1;
            row.Cells[2].AddParagraph("_______________________________");
            row.Cells[2].AddParagraph("Assinatura");
            row.Cells[2].Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].Format.Borders.Top.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(255, 255, 255);
            row.Cells[0].Format.Borders.Left.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(255, 255, 255);
            row.Cells[2].MergeRight = 1;


            table1.SetEdge(0, 0, 3, 3, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200));
            
            // Create footer
            Paragraph paragraph = sec.Footers.Primary.AddParagraph();
            paragraph.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            paragraph.AddText("_____________________________ \n");
            paragraph.AddText("Av. XYZ, 23 - CEP 60.000-000 \n");
            paragraph.AddText("Fortaleza, Ceará, Brasil \n");
            paragraph.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.DarkGray;
            paragraph.AddText("(85) 0101.0101 - FAX: (85) 3232.3232");
            paragraph.Format.Font.Size = 8;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            // Render the paragraph. You can render tables or shapes the same way.
            MigraDoc.Rendering.DocumentRenderer docRenderer = new MigraDoc.Rendering.DocumentRenderer(doc);
            docRenderer.PrepareDocument();
            
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(2), XUnit.FromCentimeter(1.5), "17cm", image);
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(2), XUnit.FromCentimeter(3), "17cm", para1);
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(2), XUnit.FromCentimeter(5), "17cm", table1);
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(2), XUnit.FromCentimeter(27.5), "17cm", paragraph);
        }

        // ====================================================================================================
        // USEFULL MIGRADOC EXAMPLES ==========================================================================
        // ====================================================================================================

        static void SamplePage1(PdfDocument document)
        {
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            // HACK²
            gfx.MUH = PdfFontEncoding.Unicode;
            //gfx.MFEH = PdfFontEmbedding.Default;
            
            //XFont font = new XFont("Verdana", 13, XFontStyle.Bold);            
            //gfx.DrawString("The following paragraph was rendered using MigraDoc:", font, XBrushes.Black,
            //    new XRect(100, 100, page.Width - 200, 300), XStringFormats.Center);
            
            // You always need a MigraDoc document for rendering.
            Document doc = new Document();
            Section sec = doc.AddSection();
            // Add a single paragraph with some text and format information.
            Paragraph para = sec.AddParagraph();
            para.Format.Alignment = ParagraphAlignment.Justify;
            para.Format.Font.Name = "Times New Roman";
            para.Format.Font.Size = 12;
            para.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.DarkGray;
            para.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.DarkGray;
            para.AddText("Duisism odigna acipsum delesenisl ");
            para.AddFormattedText("ullum in velenit", TextFormat.Bold);
            para.AddText(" ipit iurero dolum zzriliquisis nit wis dolore vel et nonsequipit, velendigna "+
                "auguercilit lor se dipisl duismod tatem zzrit at laore magna feummod oloborting ea con vel "+
                "essit augiati onsequat luptat nos diatum vel ullum illummy nonsent nit ipis et nonsequis "+
                "niation utpat. Odolobor augait et non etueril landre min ut ulla feugiam commodo lortie ex "+
                "essent augait el ing eumsan hendre feugait prat augiatem amconul laoreet. ≤≥≈≠");
            para.Format.Borders.Distance = "5pt";
            para.Format.Borders.Color = Colors.Gold;
            
            // Create a renderer and prepare (=layout) the document
            MigraDoc.Rendering.DocumentRenderer docRenderer = new MigraDoc.Rendering.DocumentRenderer(doc);
            docRenderer.PrepareDocument();
            
            // Render the paragraph. You can render tables or shapes the same way.
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(5), XUnit.FromCentimeter(10), "12cm", para);
        }

        static XRect GetRect(int index)
        {
            XRect rect = new XRect(0, 0, A4Width / 3 * 0.9, A4Height / 3 * 0.9);
            rect.X = (index % 3) * A4Width / 3 + A4Width * 0.05 / 3;
            rect.Y = (index / 3) * A4Height / 3 + A4Height * 0.05 / 3;
            return rect;
        }
        static double A4Width = XUnit.FromCentimeter(21).Point;
        static double A4Height = XUnit.FromCentimeter(29.7).Point;

        public static void DefineStyles(Document document)
        {
            // Get the predefined style Normal.
            Style style = document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Times New Roman";
            
            // Heading1 to Heading9 are predefined styles with an outline level. An outline level
            // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks) 
            // in PDF.
            
            style = document.Styles["Heading1"];
            style.Font.Name = "Tahoma";
            style.Font.Size = 14;
            style.Font.Bold = true;
            style.Font.Color = Colors.DarkBlue;
            style.ParagraphFormat.PageBreakBefore = true;
            style.ParagraphFormat.SpaceAfter = 6;
            
            style = document.Styles["Heading2"];
            style.Font.Size = 12;
            style.Font.Bold = true;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 6;
            
            style = document.Styles["Heading3"];
            style.Font.Size = 10;
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 3;
            
            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);
            
            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);
            
            // Create a new style called TextBox based on style Normal
            style = document.Styles.AddStyle("TextBox", "Normal");
            style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            style.ParagraphFormat.Borders.Width = 2.5;
            style.ParagraphFormat.Borders.Distance = "3pt";
            style.ParagraphFormat.Shading.Color = Colors.SkyBlue;
            
            // Create a new style called TOC based on style Normal
            style = document.Styles.AddStyle("TOC", "Normal");
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right, TabLeader.Dots);
            style.ParagraphFormat.Font.Color = Colors.Blue;
        }

    }
}
