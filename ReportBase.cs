using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;
using static EmissorNFe.Printing.PrintUtility;

namespace EmissorNFe.Printing
{
   public abstract class ReportBase : IPrint
   {
      // instancia a classe de funções utilitárias
      //internal readonly static PrintUtility util = new();

      // controle de impressão e visualização
      private PreviewPrintController ppController = null;
      private PrintDocument doc = null;

      private bool pageAdded = false;
      protected int record = 0;
      protected int page = 0;

      // fontes a serem usadas
      protected Font hdrAppTitle = new Font("Arial", 12, FontStyle.Bold | FontStyle.Italic);
      protected Font hdrRptName = new Font("Arial", 10, FontStyle.Bold | FontStyle.Italic);
      protected Font hdrRptTitle = new Font("Arial", 9, FontStyle.Bold);

      protected Font hdrRegular8 = new Font("Arial", 8, FontStyle.Regular);
      protected Font hdrRegularBold8 = new Font("Arial", 8, FontStyle.Bold);
      protected Font hdrRegular9 = new Font("Arial", 9, FontStyle.Regular);
      protected Font hdrRegularBold9 = new Font("Arial", 9, FontStyle.Bold);

      // formatação de strings
      protected StringFormat ltFormat = new StringFormat() { LineAlignment = StringAlignment.Near };
      protected StringFormat lcFormat = new StringFormat() { LineAlignment = StringAlignment.Center };
      protected StringFormat lbFormat = new StringFormat() { LineAlignment = StringAlignment.Far };
      protected StringFormat rtFormat = new StringFormat() { Alignment = StringAlignment.Far };
      protected StringFormat rcFormat = new StringFormat() { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };
      protected StringFormat rbFormat = new StringFormat() { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Far };
      protected StringFormat ctFormat = new StringFormat() { Alignment = StringAlignment.Center };
      protected StringFormat ccFormat = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
      protected StringFormat cbFormat = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Far };

      // variáveis locais
      protected float lineNormal;      // altura normal da linha
      protected float lineFifth;       // altrua da linha em 1/5
      protected float lineQuarter;     // altura da linha em 1/4
      protected float lineHalf;        // altura da linha em 1/2
      protected float lineSz125;       // altura da linha em 1.25
      protected float lineSz15;        // altura da linha em 1.5
      protected float lineDouble;      // altura da linha em 2.0
      protected float pgSizeW;         // largura da página
      protected float pgSizeH;         // altura da página

      protected float hdrYPosition;    // posição atual do cabeçalho

      public ReportBase()
      {
         Copies = 1;
         Settings = new ReportSettings();
      }

      public IPrintSettings Settings { get; set; } = null;

      protected bool IsPreview
      {
         get { return doc.PrintController.IsPreview; }
      }

      protected bool Landscape { get; set; }

      protected abstract void PrintData(PrintPageEventArgs e);

      #region IPrint interface implementation

      public virtual short Copies { get; set; }

      public virtual object[] DataSource { get; set; } = null;

      public virtual void Preview()
      {
         ppController = new PreviewPrintController();

         doc = new PrintDocument
         {
            PrintController = ppController
         };

         doc.DefaultPageSettings.Margins = new Margins(20, 20, 20, 20);
         //doc.DefaultPageSettings.Landscape = Landscape;
         doc.BeginPrint += new PrintEventHandler(OnBeginPrint);
         doc.PrintPage += new PrintPageEventHandler(OnPrint);

         using (var dlg = new PrintPreviewDialog())
         {
            var printButton = dlg.Controls.OfType<ToolStrip>().ElementAtOrDefault(0);
            if (printButton != null) printButton.Items[0].Click += (sender, e) => Print();
            
            dlg.Document = doc;
            dlg.WindowState = FormWindowState.Maximized;
            dlg.ShowDialog((Form)Settings.Window);
         }
      }

      public virtual void Print()
      {
         doc = new PrintDocument();
         doc.DefaultPageSettings.Margins = new Margins(20, 20, 20, 20);
         //doc.DefaultPageSettings.Landscape = Landscape;
         doc.BeginPrint += new PrintEventHandler(OnBeginPrint);
         doc.PrintPage += new PrintPageEventHandler(OnPrint);
         doc.Print();
      }

      protected virtual void OnBeginPrint(object sender, PrintEventArgs e)
      {
         doc.DefaultPageSettings.Landscape = Landscape;
         record = 0;
         page = 1;
      }

      protected virtual void OnPrint(object sender, PrintPageEventArgs e)
      {
         // não há dados para impressão
         if (DataSource == null) return;

         if (e.Graphics == null)
            throw new ArgumentNullException(nameof(e));

         // variáveis de uso geral
         var g = e.Graphics;

         // define o valor das variáveis de uso para impressão
         lineNormal = g.MeasureString("W", hdrRegular8).Height;
         lineFifth = lineNormal * 0.2f;
         lineQuarter = lineNormal * 0.25f;
         lineHalf = lineNormal * 0.5f;
         lineSz125 = lineNormal * 1.25f;
         lineSz15 = lineNormal * 1.5f;
         lineDouble = lineNormal * 2f;

         if (Landscape)
         {
            pgSizeH = ToInch(200);
            pgSizeW = ToInch(289);
         }
         else
         {
            pgSizeH = ToInch(289);
            pgSizeW = ToInch(205);
         }

         // define nenhuma nova página adicionada
         pageAdded = false;

         if (IsPreview)
            g.TranslateTransform(e.PageSettings.HardMarginX, e.PageSettings.HardMarginY);

         // *******************************
         // imprime o cabeçalho
         // *******************************
         // desenha o logo
         if (Settings.ImageLogo != null)
            g.DrawImage(Settings.ImageLogo, ToInch(10), ToInch(5));

         float hY = ToInch(5);

         // imprime o título
         g.DrawString(string.Format("{0} v{1}", Settings.AppTitle, Settings.AppVersion), hdrAppTitle, Brushes.Black, ToInch(25), hY);
         g.DrawString(Settings.ReportName, hdrRptName, Brushes.Black, ToInch(25), hY + lineSz125);
         g.DrawString(Settings.ReportTitle, hdrRptTitle, Brushes.Black, ToInch(25), hY + 2 * lineSz125);
         //g.DrawLine(new Pen(Color.Black, 1.5F), util.ToInch(0), util.ToInch(13) + lineNormal, util.ToInch(202), util.ToInch(13) + lineNormal);

         var rcDh = new RectangleF(ToInch(Landscape ? 245 : 160), hY, ToInch(45), lineSz15);
         var rcPag = new RectangleF(ToInch(Landscape ? 245 : 160), hY + lineNormal, ToInch(45), lineSz15);
         g.DrawString(string.Format("{0:dd/MM/yyyy HH:mm}", DateTime.Now), hdrRegular8, Brushes.Black, rcDh, rtFormat);
         g.DrawString(string.Format("Pág.: {0}", page), hdrRegular8, Brushes.Black, rcPag, rtFormat);

         hdrYPosition = hY + 2 * lineSz125;
         PrintData(e);

         // fim da impressão
         if (!pageAdded)
            e.HasMorePages = false;
      }

      #endregion

      protected void AddPage(PrintPageEventArgs e)
      {
         e.HasMorePages = true;
         pageAdded = true;
         page++;
      }

      protected virtual void PrintFooter(Graphics gr, float lineSize)
      {
         gr.DrawLine(new Pen(Color.Black, lineSize), ToInch(5), pgSizeH, pgSizeW, pgSizeH);
      }

      protected static int GetTotalPageCount(ReportBase report)
      {
         int count = 0;
         using (var printDocument = new PrintDocument())
         {
            printDocument.PrintController = new PreviewPrintController();
            printDocument.BeginPrint += report.OnBeginPrint;
            printDocument.PrintPage += (o, e) => { report.OnPrint(o, e); count++; };
            printDocument.Print();
            return count;
         }
      }
   }
}
