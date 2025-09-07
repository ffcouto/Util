using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;
using EmissorNFe.Models;
using static EmissorNFe.Printing.PrintUtility;

namespace EmissorNFe.Printing
{
   public class DanfeReport : ReportBase
   {
      // origem dos dados
      private Fatura ds = null;
      private NFeEvento nfe = null;
      private Empresa empr = null;
      private Cliente cli = null;
      private Transportadora transp = null;
      //private Expedicao exp = null;
      private Image logo = null;

      private int pages = 0;
      private int curLote = 0;
      private bool itemPrinted = false;
      private bool lotePrinted = false;
      private bool fciPrinted = false;
      private bool fecpPrinted = false;
      private bool ctrlPrinted = false;

      // fontes do DANFE
      private readonly Font hdrRegular4 = new Font("Arial", 4, FontStyle.Regular);
      private readonly Font hdrRegular5 = new Font("Arial", 5, FontStyle.Regular);
      private readonly Font hdrRegular6 = new Font("Arial", 6, FontStyle.Regular);
      private readonly Font hdrRegular7 = new Font("Arial", 7, FontStyle.Regular);
      private readonly Font hdrRegularBold7 = new Font("Arial", 7, FontStyle.Bold);

      private readonly Font hdrCourierBold4 = new Font("Courier New", 4, FontStyle.Bold);
      private readonly Font hdrCourierBold6 = new Font("Courier New", 6, FontStyle.Bold);
      private readonly Font hdrCourierBold7 = new Font("Courier New", 7, FontStyle.Bold);
      private readonly Font hdrCourierBold8 = new Font("Courier New", 8, FontStyle.Bold);
      private readonly Font hdrCourierBold9 = new Font("Courier New", 9, FontStyle.Bold);

      private readonly Pen pen2 = new Pen(Color.Black, 2f);

      private readonly Brush brBlack = Brushes.Black;

      // define os limites da página
      private readonly static float px1 = ToInch(5);
      private readonly static float px2 = ToInch(204);

      // define as variáveis de posicionamento e tamanho
      private readonly static float rowHeight = ToInch(6.5);
      private readonly static float rowHeight2 = rowHeight * 2f;
      private readonly static float rowHeight3 = rowHeight * 3f;
      private readonly static float rowHeight4 = rowHeight * 4f;

      private readonly static float spcToStartRow = ToInch(1);
      private readonly static float spcToStartRow2 = ToInch(3);
      private readonly static float sectionHeight = ToInch(3);

      // define as variáveis de posicionamento fixo das seções e campos da nf
      private readonly static float pyDeliveryHdr = ToInch(6);                                  // início do canhoto
      private readonly static float dashedLineStartYPos = pyDeliveryHdr + rowHeight2 + spcToStartRow * 2;    // linha tracejada
      private readonly static float pyInvoiceHdr = dashedLineStartYPos + spcToStartRow * 2;     // logo + info nfe
      private readonly static float pyInvoiceHdr2 = pyInvoiceHdr + rowHeight4 + ToInch(3.5);    // operação
      private readonly static float secDest = pyInvoiceHdr2 + rowHeight2 + sectionHeight;       // seção destinatário
      private readonly static float pyDest = secDest + spcToStartRow2;                          // início destinatário
      private readonly static float secDupl = pyDest + rowHeight3 + sectionHeight;              // seção faturas/duplicatas
      private readonly static float pyDupl = secDupl + spcToStartRow2;                          // início faturas
      private readonly static float pyStartDupl = pyDupl + ToInch(1.5);                         // início duplicatas
      private static readonly float secIssqn = ToInch(243);                                     // seção cálculo ISSQN
      private static readonly float pyIssqnHdr = secIssqn + spcToStartRow2;                     // início seção cálculo ISSQN
      private static readonly float secInfoAdic = pyIssqnHdr + rowHeight + sectionHeight;       // seção info adicional
      private static readonly float pyInfoAdic = secInfoAdic + spcToStartRow2;                  // início seção info adicional
      private static readonly float pyEndPage = ToInch(290);

      // define as variáveis de posicionamento variável das seções e campos da nf
      private static int countLineDupls;           // n° de linhas para a seção faturas
      private static float pyEndDupl;              // fim duplicatas
      private static float secTax;                 // seção impostos
      private static float pyTaxCalc;              // início dos tributos
      private static float secTransport;           // seção transporte
      private static float pyTransport;            // início seção transporte
      private static float secProduct;             // seção produtos/serviços
      private static float pyProduct;              // início seção produtos/serviços
      private static float pyItems;                // início dos items

      private static float szProductSection;       // altura da área dos itens da nota fiscal

      public DanfeReport()
         : base()
      { }

      #region ReportBase abstract class implementation

      public override object[] DataSource
      {
         get { return new object[] { ds }; }
         set { ds = value.Cast<Fatura>().FirstOrDefault(); }
      }

      protected override void OnBeginPrint(object sender, PrintEventArgs e)
      {
         base.OnBeginPrint(sender, e);

         if (Settings.UserData.TryGetValue("NFE", out object value1))
            nfe = (NFeEvento)value1;

         if (Settings.UserData.TryGetValue("EMPRESA", out object value2))
            empr = (Empresa)value2;

         if (Settings.UserData.TryGetValue("CLIENTE", out object value3))
            cli = (Cliente)value3;

         if (Settings.UserData.TryGetValue("TRANSPORTE", out object value4))
            transp = (Transportadora)value4;

         //if (Settings.UserData.TryGetValue("EXPEDICAO", out object value5))
         //   exp = (Expedicao)value5;

         if (Settings.UserData.TryGetValue("LOGO", out object value6))
            logo = (Image)value6;

         // calcula o números de linhas para a seção fatura/duplicata
         countLineDupls = (ds.Parcelas - 1) / 3 + 1;
         
         // calcula a posição após a última linha da seção fatura/duplicata
         float pyLastRowDupl = pyStartDupl + countLineDupls * spcToStartRow2;
         
         // atribui o posicionamento das seções variáveis
         pyEndDupl = (pyLastRowDupl > pyStartDupl + rowHeight) ? pyLastRowDupl : pyStartDupl + rowHeight;
         
         secTax = pyEndDupl + spcToStartRow2;
         pyTaxCalc = secTax + spcToStartRow2;
         secTransport = pyTaxCalc + rowHeight2 + spcToStartRow2;
         pyTransport = secTransport + spcToStartRow2;
         secProduct = pyTransport + rowHeight3 + spcToStartRow2;
         pyProduct = secProduct + spcToStartRow2;
         pyItems = pyProduct + ToInch(5) + spcToStartRow;

         szProductSection = secIssqn - (sectionHeight * 2) - pyItems;

         itemPrinted = false;
         lotePrinted = false;
         fciPrinted = false;
         fecpPrinted = false;
         ctrlPrinted = false;

         record = 0;
      }

      protected override void OnPrint(object sender, PrintPageEventArgs e)
      {
         // define o valor das variáveis de uso para impressão
         lineNormal = e.Graphics.MeasureString("W", hdrRegular8).Height;
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

         if (IsPreview)
            e.Graphics.TranslateTransform(e.PageSettings.HardMarginX, e.PageSettings.HardMarginY);

         hdrYPosition = 0;

         float szTotalItems = GetNeededSizeItems(e.Graphics);
         pages = (int)Math.Ceiling(szTotalItems / szProductSection);

         PrintData(e);
      }

      protected override void PrintData(PrintPageEventArgs e)
      {
         // maniputador gráfico
         var g = e.Graphics;

         var cnpjMask = new MaskedTextProvider("00,000,000/0000-00");
         var ncmMask = new MaskedTextProvider("0000,00,00");

         // canhoto recebimento
         g.DrawLine(pen2, px1, pyDeliveryHdr, px2, pyDeliveryHdr);

         g.DrawString($"RECEBEMOS DE {empr.Nome.ToUpper()}, OS PRODUTOS E/OU SERVIÇOS " +
            "CONSTANTES DA NOTA FISCAL ELETRÔNICA INDICADO AO LADO", hdrRegular5, brBlack, px1, pyDeliveryHdr + spcToStartRow);

         g.DrawLine(pen2, px1, pyDeliveryHdr + rowHeight, ToInch(171.5), pyDeliveryHdr + rowHeight);

         g.DrawString("DATA DE RECEBIMENTO", hdrRegular5, brBlack, px1, pyDeliveryHdr + rowHeight + spcToStartRow);
         g.DrawString("IDENTIFICAÇÃO E ASSINATURA DO RECEBEDOR", hdrRegular5, brBlack, ToInch(41), pyDeliveryHdr + rowHeight + spcToStartRow);

         g.DrawLine(pen2, px1, pyDeliveryHdr + rowHeight2, px2, pyDeliveryHdr + rowHeight2);

         g.DrawLine(pen2, ToInch(40), pyDeliveryHdr + rowHeight, ToInch(40), pyDeliveryHdr + rowHeight2);
         g.DrawLine(pen2, ToInch(171.5), pyDeliveryHdr, ToInch(171.5), pyDeliveryHdr + rowHeight2);

         g.DrawString("NF-e", hdrRegular7, brBlack, ToInch(188), pyDeliveryHdr + ToInch(1.5), ctFormat);
         g.DrawString("N°", hdrRegular7, brBlack, ToInch(174), pyDeliveryHdr + ToInch(5));
         g.DrawString($"{ds.NotaFiscal:###,###,###}", hdrCourierBold9, brBlack, ToInch(184), pyDeliveryHdr + ToInch(5));
         g.DrawString("SÉRIE", hdrRegular7, brBlack, ToInch(174), pyDeliveryHdr + ToInch(9.5));
         g.DrawString($"{ds.SerieNF:000}", hdrCourierBold9, brBlack, ToInch(184), pyDeliveryHdr + ToInch(9));
         // fim - canhoto recebimento

         var penDotted = new Pen(brBlack, 1f) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash, DashPattern = new float[] { 6, 3 } };

         g.DrawLine(penDotted, px1, dashedLineStartYPos, px2, dashedLineStartYPos);

         // cabeçalho da nota
         g.DrawImage(logo, ToInch(10.5), ToInch(30));

         var rcEmprNome = new RectangleF(ToInch(25), ToInch(29), ToInch(62), hdrRegularBold8.GetHeight(g));
         var rcEmprEnd1 = new RectangleF(ToInch(25), ToInch(35), ToInch(62), hdrRegular6.GetHeight(g));
         var rcEmprEnd2 = new RectangleF(ToInch(25), ToInch(39), ToInch(62), hdrRegular6.GetHeight(g));
         var rcEmprEnd3 = new RectangleF(ToInch(25), ToInch(43), ToInch(62), hdrRegular6.GetHeight(g));

         g.DrawString($"{empr.Nome}", hdrRegularBold8, brBlack, rcEmprNome, ctFormat);
         g.DrawString($"{empr.Endereco.Logradouro}, {empr.Endereco.Numero} {empr.Endereco.Complemento}", hdrRegular6, brBlack, rcEmprEnd1, ctFormat);
         g.DrawString($"{empr.Endereco.Bairro} - {empr.Endereco.Cidade}/{empr.Endereco.Estado}", hdrRegular6, brBlack, rcEmprEnd2, ctFormat);
         g.DrawString($"{empr.Endereco.Cep.Insert(5, "-")}   {empr.Telefone.PhoneNumber()}", hdrRegular6, brBlack, rcEmprEnd3, ctFormat);

         // linhas horizontais
         g.DrawLine(pen2, ToInch(92), pyInvoiceHdr, px2, pyInvoiceHdr);
         g.DrawLine(pen2, ToInch(125), pyInvoiceHdr + rowHeight2, px2, pyInvoiceHdr + rowHeight2);
         g.DrawLine(pen2, ToInch(125), pyInvoiceHdr + rowHeight3, px2, pyInvoiceHdr + rowHeight3);
         g.DrawLine(pen2, px1, pyInvoiceHdr2, px2, pyInvoiceHdr2);

         // linhas verticais
         g.DrawLine(pen2, ToInch(92), pyInvoiceHdr, ToInch(92), pyInvoiceHdr2);
         g.DrawLine(pen2, ToInch(125), pyInvoiceHdr, ToInch(125), pyInvoiceHdr2);

         var rcDanfe0 = new RectangleF(ToInch(93), pyInvoiceHdr + ToInch(1.5), ToInch(31), hdrRegular7.GetHeight(g));
         var rcDanfe1 = new RectangleF(ToInch(93), pyInvoiceHdr + ToInch(4.5), ToInch(31), hdrRegular7.GetHeight(g));
         var rcDanfe2 = new RectangleF(ToInch(93), pyInvoiceHdr + ToInch(7), ToInch(31), hdrRegular7.GetHeight(g));
         var rcDanfe3 = new RectangleF(ToInch(93), pyInvoiceHdr + ToInch(9.5), ToInch(31), hdrRegular7.GetHeight(g));

         //var rcBarCode = new RectangleF(ToInch(126), pyInvoiceHdr + 5, ToInch(77), ToInch(11));
         
         Image barCode = new Code128CBarCode().Create(nfe.NFeId, 295, 43, out _);
         g.DrawImage(barCode, ToInch(128), pyInvoiceHdr + 5);

         g.DrawString("DANFE", hdrRegular7, brBlack, rcDanfe0, ctFormat);
         g.DrawString("DOCUMENTO AUXILIAR", hdrRegular6, brBlack, rcDanfe1, ctFormat);
         g.DrawString("DA NOTA FISCAL", hdrRegular6, brBlack, rcDanfe2, ctFormat);
         g.DrawString("ELETRÔNICA", hdrRegular6, brBlack, rcDanfe3, ctFormat);

         g.DrawString("0 - ENTRADA", hdrRegular5, brBlack, ToInch(94.5), pyInvoiceHdr + ToInch(13.5));
         g.DrawString("1 - SAÍDA", hdrRegular5, brBlack, ToInch(94.5), pyInvoiceHdr + ToInch(16));

         var rcTipo = new RectangleF(ToInch(116), pyInvoiceHdr + ToInch(13), ToInch(6), ToInch(6));
         g.DrawRectangle(Pens.Black, Rectangle.Truncate(rcTipo));
         rcTipo.Offset(ToInch(0.25), ToInch(0.5));
         g.DrawString($"{ds.Finalidade}", hdrCourierBold9, brBlack, rcTipo, ccFormat);

         g.DrawString("N.°:", hdrRegular5, brBlack, ToInch(94.5), pyInvoiceHdr + ToInch(20.5));
         g.DrawString($"{ds.NotaFiscal:###,###,###}", hdrCourierBold8, brBlack, ToInch(113), pyInvoiceHdr + ToInch(20), ctFormat);
         g.DrawString("SÉRIE:", hdrRegular5, brBlack, ToInch(94.5), pyInvoiceHdr + ToInch(23.5));
         g.DrawString($"{ds.SerieNF:000}", hdrCourierBold8, brBlack, ToInch(113), pyInvoiceHdr + ToInch(23), ctFormat);
         g.DrawString("FOLHA:", hdrRegular5, brBlack, ToInch(94.5), pyInvoiceHdr + ToInch(26.5));
         g.DrawString($"{page}/{pages}", hdrCourierBold8, brBlack, ToInch(113), pyInvoiceHdr + ToInch(26), ctFormat);

         g.DrawString("CHAVE DE ACESSO", hdrRegular5, brBlack, ToInch(126.5), pyInvoiceHdr + ToInch(14));
         g.DrawString($"{nfe.NFeId}", hdrRegular8, brBlack, ToInch(130), pyInvoiceHdr + ToInch(16));

         var rcCons1 = new RectangleF(ToInch(125), pyInvoiceHdr + ToInch(22), ToInch(79), hdrRegularBold7.GetHeight(g));
         var rcCons2 = new RectangleF(ToInch(125), pyInvoiceHdr + ToInch(25), ToInch(79), hdrRegularBold7.GetHeight(g));

         g.DrawString("Consulta de autenticidade no portal nacional da NF-e", hdrRegularBold7, brBlack, rcCons1, ctFormat);
         g.DrawString("www.nfe.fazenda.gov.br/portal ou no site da Sefaz Autorizadora", hdrRegularBold7, brBlack, rcCons2, ctFormat);

         g.DrawString("NATUREZA DA OPERAÇÃO", hdrRegular5, brBlack, px1, pyInvoiceHdr2 + spcToStartRow);
         g.DrawString($"{ds.CFOP} {ds.CfopDescricao}", hdrCourierBold8, brBlack, px1 + ToInch(3), pyInvoiceHdr2 + spcToStartRow2);

         g.DrawLine(pen2, ToInch(125), pyInvoiceHdr2, ToInch(125), pyInvoiceHdr2 + rowHeight);
         g.DrawString("PROTOCOLO DE AUTORIZAÇÃO DE USO", hdrRegular5, brBlack, ToInch(126), pyInvoiceHdr2 + spcToStartRow);
         g.DrawString($"{nfe.Recibo} {nfe.DataEvento:dd/MM/yyyy HH:mm:ss}", hdrCourierBold8, brBlack, ToInch(135), pyInvoiceHdr2 + spcToStartRow2);

         g.DrawLine(pen2, px1, pyInvoiceHdr2 + rowHeight, px2, pyInvoiceHdr2 + rowHeight);

         g.DrawString("INSCRIÇÃO ESTADUAL", hdrRegular5, brBlack, px1, pyInvoiceHdr2 + rowHeight + spcToStartRow);
         g.DrawString($"{empr.InscricaoEstadual}", hdrCourierBold8, brBlack, px1 + ToInch(3), pyInvoiceHdr2 + rowHeight + spcToStartRow2);

         g.DrawLine(pen2, ToInch(75), pyInvoiceHdr2 + rowHeight, ToInch(75), pyInvoiceHdr2 + rowHeight2);
         g.DrawString("INSCRIÇÃO ESTADUAL DO SUBST. TRIBUT.", hdrRegular5, brBlack, ToInch(76), pyInvoiceHdr2 + rowHeight + spcToStartRow);

         cnpjMask.Set(empr.CNPJ);

         g.DrawLine(pen2, ToInch(147), pyInvoiceHdr2 + rowHeight, ToInch(147), pyInvoiceHdr2 + rowHeight2);
         g.DrawString("CNPJ", hdrRegular5, brBlack, ToInch(148), pyInvoiceHdr2 + rowHeight + spcToStartRow);
         g.DrawString($"{cnpjMask}", hdrCourierBold8, brBlack, ToInch(150), pyInvoiceHdr2 + rowHeight + spcToStartRow2);

         g.DrawLine(pen2, px1, pyInvoiceHdr2 + rowHeight2, px2, pyInvoiceHdr2 + rowHeight2);

         g.DrawString("DESTINATÁRIO / REMETENTE", hdrRegular6, brBlack, px1, secDest);
         g.DrawLine(pen2, px1, pyDest, px2, pyDest);

         g.DrawString("NOME / RAZÃO SOCIAL", hdrRegular5, brBlack, px1, pyDest + spcToStartRow);
         g.DrawString($"{cli.Nome}", hdrCourierBold8, brBlack, px1 + ToInch(3), pyDest + spcToStartRow2);

         cnpjMask.Set(cli.CNPJ);

         g.DrawLine(pen2, ToInch(130), pyDest, ToInch(130), pyDest + rowHeight);
         g.DrawString("C.N.P.J. / C.P.F.", hdrRegular5, brBlack, ToInch(131), pyDest + spcToStartRow);
         g.DrawString($"{cnpjMask}", hdrCourierBold8, brBlack, ToInch(131), pyDest + spcToStartRow2);

         g.DrawLine(pen2, ToInch(174), pyDest, ToInch(174), pyDest + rowHeight);
         g.DrawString("DATA DA EMISSÃO", hdrRegular5, brBlack, ToInch(175), pyDest + spcToStartRow);
         g.DrawString($"{ds.DataFatura:dd/MM/yyyy}", hdrCourierBold8, brBlack, ToInch(182), pyDest + spcToStartRow2);

         g.DrawLine(pen2, px1, pyDest + rowHeight, px2, pyDest + rowHeight);

         var rcLogr = new RectangleF(px1 + ToInch(3), pyDest + rowHeight + spcToStartRow2, ToInch(97), hdrCourierBold8.GetHeight(g));
         g.DrawString("ENDEREÇO", hdrRegular5, brBlack, px1, pyDest + rowHeight + spcToStartRow);
         g.DrawString($"{cli.Endereco.Logradouro}, {cli.Endereco.Numero}", hdrCourierBold8, brBlack, rcLogr);

         var rcBairro = new RectangleF(ToInch(110), pyDest + rowHeight + spcToStartRow2, ToInch(43), hdrCourierBold8.GetHeight(g));
         g.DrawLine(pen2, ToInch(108), pyDest + rowHeight, ToInch(108), pyDest + rowHeight2);
         g.DrawString("BAIRRO / DISTRITO", hdrRegular5, brBlack, ToInch(109), pyDest + rowHeight + spcToStartRow);
         g.DrawString($"{cli.Endereco.Bairro}", hdrCourierBold8, brBlack, rcBairro);

         g.DrawLine(pen2, ToInch(154), pyDest + rowHeight, ToInch(154), pyDest + rowHeight2);
         g.DrawString("CEP", hdrRegular5, brBlack, ToInch(155), pyDest + rowHeight + spcToStartRow);
         g.DrawString($"{cli.Endereco.Cep.Insert(5, "-")}", hdrCourierBold8, brBlack, ToInch(156.5), pyDest + rowHeight + spcToStartRow2);

         g.DrawLine(pen2, ToInch(174), pyDest + rowHeight, ToInch(174), pyDest + rowHeight2);
         g.DrawString("DATA DA ENTRADA / SAÍDA", hdrRegular5, brBlack, ToInch(175), pyDest + rowHeight + spcToStartRow);
         g.DrawString($"{ds.DataFatura:dd/MM/yyyy}", hdrCourierBold8, brBlack, ToInch(182), pyDest + rowHeight + spcToStartRow2);

         g.DrawLine(pen2, px1, pyDest + rowHeight2, px2, pyDest + rowHeight2);

         var rcCidade = new RectangleF(px1 + ToInch(3), pyDest + rowHeight2 + spcToStartRow2, ToInch(97), hdrCourierBold8.GetHeight(g));
         g.DrawString("MUNICÍPIO", hdrRegular5, brBlack, px1, pyDest + rowHeight2 + spcToStartRow);
         g.DrawString($"{cli.Endereco.Cidade}", hdrCourierBold8, brBlack, rcCidade);

         var cliContato = cli.Contatos.ToArray().FirstOrDefault();
         g.DrawLine(pen2, ToInch(82), pyDest + rowHeight2, ToInch(82), pyDest + rowHeight3);
         g.DrawString("FONE / FAX", hdrRegular5, brBlack, ToInch(83), pyDest + rowHeight2 + spcToStartRow);
         g.DrawString($"{cliContato?.Telefone.PhoneNumber()}", hdrCourierBold8, brBlack, ToInch(84), pyDest + rowHeight2 + spcToStartRow2);

         g.DrawLine(pen2, ToInch(120), pyDest + rowHeight2, ToInch(120), pyDest + rowHeight3);
         g.DrawString("UF", hdrRegular5, brBlack, ToInch(121), pyDest + rowHeight2 + spcToStartRow);
         g.DrawString($"{cli.Endereco.Estado}", hdrCourierBold8, brBlack, ToInch(122), pyDest + rowHeight2 + spcToStartRow2);

         var rcIE = new RectangleF(ToInch(131), pyDest + rowHeight2 + spcToStartRow2, ToInch(42), hdrCourierBold8.GetHeight(g));
         g.DrawLine(pen2, ToInch(130), pyDest + rowHeight2, ToInch(130), pyDest + rowHeight3);
         g.DrawString("INSCRIÇÃO ESTADUAL", hdrRegular5, brBlack, ToInch(131), pyDest + rowHeight2 + spcToStartRow);
         g.DrawString($"{cli.Inscricao}", hdrCourierBold8, brBlack, rcIE, ctFormat);

         g.DrawLine(pen2, ToInch(174), pyDest + rowHeight2, ToInch(174), pyDest + rowHeight3);
         g.DrawString("HORA DA SAÍDA", hdrRegular5, brBlack, ToInch(175), pyDest + rowHeight2 + spcToStartRow);
         g.DrawString($"{ds.DataFatura:HH:mm:ss}", hdrCourierBold8, brBlack, ToInch(182), pyDest + rowHeight2 + spcToStartRow2);

         g.DrawLine(pen2, px1, pyDest + rowHeight3, px2, pyDest + rowHeight3);

         g.DrawString("FATURA / DUPLICATA", hdrRegular6, brBlack, px1, secDupl);
         g.DrawLine(pen2, px1, pyDupl, px2, pyDupl);

         int dRow = 0;
         int dCol = 0;
         float pyRowDupl;

         foreach (FaturaDuplicata d in ds.Duplicatas)
         {
            if (int.Parse(d.Numero.Substring(0, 2)) != ds.Empresa) continue;

            if (dCol > 2)
            {
               dRow++;
               dCol = 0;
            }

            float pxDoc = ToInch(64) * dCol + ToInch(10);
            float pxData = ToInch(64) * dCol + ToInch(30);
            float pxVlr = ToInch(64) * dCol + ToInch(69);
            pyRowDupl = pyStartDupl + dRow * spcToStartRow2;

            g.DrawString($"{d.Numero}", hdrCourierBold8, brBlack, pxDoc, pyRowDupl);
            g.DrawString($"{d.Vencimento:dd/MM/yyyy}", hdrCourierBold8, brBlack, pxData, pyRowDupl);
            g.DrawString($"{d.Valor:###,###,##0.00}", hdrCourierBold8, brBlack, pxVlr, pyRowDupl, rtFormat);

            dCol++;
         }

         g.DrawLine(pen2, px1, pyEndDupl, px2, pyEndDupl);

         g.DrawString("CÁLCULO DO IMPOSTO", hdrRegular6, brBlack, px1, secTax);
         g.DrawLine(pen2, px1, pyTaxCalc, px2, pyTaxCalc);

         var rcBCICMS = new RectangleF(px1 + ToInch(3), pyTaxCalc + spcToStartRow2, ToInch(43), hdrCourierBold8.GetHeight(g));
         g.DrawString("BASE DE CÁLCULO DO ICMS", hdrRegular5, brBlack, px1, pyTaxCalc + spcToStartRow);
         g.DrawString($"{ds.BaseCalculoICMS:###,###,##0.00}", hdrCourierBold8, brBlack, rcBCICMS, ctFormat);

         g.DrawLine(pen2, ToInch(54), pyTaxCalc, ToInch(54), pyTaxCalc + rowHeight);
         g.DrawString("VALOR DO ICMS", hdrRegular5, brBlack, ToInch(55), pyTaxCalc + spcToStartRow);
         g.DrawString($"{ds.ValorIcms:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(87), pyTaxCalc + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, ToInch(93), pyTaxCalc, ToInch(93), pyTaxCalc + rowHeight);
         g.DrawString("BASE DE CÁLCULO DO I.C.M.S. S.T.", hdrRegular5, brBlack, ToInch(94), pyTaxCalc + spcToStartRow);
         //g.DrawString($"{ds.Impostos.BaseCalculoICMSST:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(124), pyTaxCalc + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, ToInch(130), pyTaxCalc, ToInch(130), pyTaxCalc + rowHeight);
         g.DrawString("VALOR DO I.C.M.S. SUBSTITUIÇÃO", hdrRegular5, brBlack, ToInch(131), pyTaxCalc + spcToStartRow);
         //g.DrawString($"{ds.Impostos.TotalICMSST:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(168), pyTaxCalc + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, ToInch(174), pyTaxCalc, ToInch(174), pyTaxCalc + rowHeight);
         g.DrawString("VALOR TOTAL DOS PRODUTOS", hdrRegular5, brBlack, ToInch(175), pyTaxCalc + spcToStartRow);
         g.DrawString($"{ds.ValorNF - ds.ValorIpi:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(204), pyTaxCalc + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, px1, pyTaxCalc + rowHeight, px2, pyTaxCalc + rowHeight);

         g.DrawString("VALOR DO FRETE", hdrRegular5, brBlack, px1, pyTaxCalc + rowHeight + spcToStartRow);
         g.DrawString($"{ds.ValorFrete:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(40), pyTaxCalc + rowHeight + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, ToInch(46), pyTaxCalc + rowHeight, ToInch(46), pyTaxCalc + rowHeight2);
         g.DrawString("VALOR DO SEGURO", hdrRegular5, brBlack, ToInch(47), pyTaxCalc + rowHeight + spcToStartRow);
         //g.DrawString($"{ds.Impostos.TotalSeguro:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(75), pyTaxCalc + rowHeight + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, ToInch(81), pyTaxCalc + rowHeight, ToInch(81), pyTaxCalc + rowHeight2);
         g.DrawString("DESCONTO", hdrRegular5, brBlack, ToInch(82), pyTaxCalc + rowHeight + spcToStartRow);
         g.DrawString($"{ds.ValorDesconto:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(108), pyTaxCalc + rowHeight + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, ToInch(109), pyTaxCalc + rowHeight, ToInch(109), pyTaxCalc + rowHeight2);
         g.DrawString("OUTRAS DESPESAS ACESSÓRIAS", hdrRegular5, brBlack, ToInch(110), pyTaxCalc + rowHeight + spcToStartRow);
         //g.DrawString($"{ds.Impostos.TotalOutrasDespesas:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(138), pyTaxCalc + rowHeight + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, ToInch(141), pyTaxCalc + rowHeight, ToInch(141), pyTaxCalc + rowHeight2);
         g.DrawString("VALOR TOTAL DO IPI", hdrRegular5, brBlack, ToInch(142), pyTaxCalc + rowHeight + spcToStartRow);
         g.DrawString($"{ds.ValorIpi:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(171), pyTaxCalc + rowHeight + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, ToInch(174), pyTaxCalc + rowHeight, ToInch(174), pyTaxCalc + rowHeight2);
         g.DrawString("VALOR TOTAL DA NOTA", hdrRegular5, brBlack, ToInch(175), pyTaxCalc + rowHeight + spcToStartRow);
         g.DrawString($"{ds.ValorNF:###,###,##0.00}", hdrCourierBold8, brBlack, ToInch(204), pyTaxCalc + rowHeight + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, px1, pyTaxCalc + rowHeight2, px2, pyTaxCalc + rowHeight2);

         g.DrawString("TRANSPORTADOR / VOLUMES TRANSPORTADOS", hdrRegular6, brBlack, px1, secTransport);
         g.DrawLine(pen2, px1, pyTransport, px2, pyTransport);

         var rcTranspNome = new RectangleF(px1 + ToInch(3), pyTransport + spcToStartRow2, ToInch(82), hdrCourierBold8.GetHeight(g));
         g.DrawString("NOME / RAZÃO SOCIAL", hdrRegular5, brBlack, px1, pyTransport + spcToStartRow);
         g.DrawString($"{transp?.Nome}", hdrCourierBold8, brBlack, rcTranspNome);

         var modFrete = ModalidadeFrete.GetItem(ds.ModalidadeFrete);
         g.DrawLine(pen2, ToInch(94), pyTransport, ToInch(94), pyTransport + rowHeight);
         g.DrawString("FRETE POR CONTA", hdrRegular5, brBlack, ToInch(95), pyTransport + spcToStartRow);
         g.DrawString($"{modFrete.Codigo:0}. {modFrete.Responsavel}", hdrCourierBold8, brBlack, ToInch(96), pyTransport + spcToStartRow2);

         g.DrawLine(pen2, ToInch(125), pyTransport, ToInch(125), pyTransport + rowHeight);
         g.DrawString("CÓDIGO ANTT", hdrRegular5, brBlack, ToInch(126), pyTransport + spcToStartRow);

         g.DrawLine(pen2, ToInch(141), pyTransport, ToInch(141), pyTransport + rowHeight);
         g.DrawString("PLACA DO VEÍCULO", hdrRegular5, brBlack, ToInch(142), pyTransport + spcToStartRow);
         g.DrawString($"{ds.Placa.Insert(3, "-")}", hdrCourierBold8, brBlack, ToInch(144), pyTransport + spcToStartRow2);

         g.DrawLine(pen2, ToInch(162), pyTransport, ToInch(162), pyTransport + rowHeight);
         g.DrawString("UF", hdrRegular5, brBlack, ToInch(163), pyTransport + spcToStartRow);
         g.DrawString($"{ds.PlacaUF}", hdrCourierBold8, brBlack, ToInch(164), pyTransport + spcToStartRow2);

         string transpCnpj = transp?.CNPJ ?? "";
         cnpjMask.Set(transpCnpj);

         g.DrawLine(pen2, ToInch(169), pyTransport, ToInch(169), pyTransport + rowHeight);
         g.DrawString("CNPJ", hdrRegular5, brBlack, ToInch(170), pyTransport + spcToStartRow);
         g.DrawString(string.IsNullOrWhiteSpace(transpCnpj) ? "" : $"{cnpjMask}", hdrCourierBold8, brBlack, ToInch(170), pyTransport + spcToStartRow2);

         g.DrawLine(pen2, px1, pyTransport + rowHeight, px2, pyTransport + rowHeight);

         rcLogr = new RectangleF(px1 + ToInch(3), pyTransport + rowHeight + spcToStartRow2, ToInch(83), hdrCourierBold8.GetHeight(g));
         g.DrawString("ENDEREÇO", hdrRegular5, brBlack, px1, pyTransport + rowHeight + spcToStartRow);

         if (transp != null && !transp.Endereco.IsEmpty())
            g.DrawString($"{transp?.Endereco.Logradouro}, {transp?.Endereco.Numero} {transp?.Endereco.Complemento}", hdrCourierBold8, brBlack, rcLogr);

         rcCidade = new RectangleF(ToInch(95), pyTransport + rowHeight + spcToStartRow2, ToInch(67), hdrCourierBold8.GetHeight(g));
         g.DrawLine(pen2, ToInch(94), pyTransport + rowHeight, ToInch(94), pyTransport + rowHeight2);
         g.DrawString("MUNICÍPIO", hdrRegular5, brBlack, ToInch(95), pyTransport + rowHeight + spcToStartRow);
         g.DrawString($"{transp?.Endereco.Cidade}", hdrCourierBold8, brBlack, rcCidade);

         g.DrawLine(pen2, ToInch(162), pyTransport + rowHeight, ToInch(162), pyTransport + rowHeight2);
         g.DrawString("UF", hdrRegular5, brBlack, ToInch(163), pyTransport + rowHeight + spcToStartRow);
         g.DrawString($"{transp?.Endereco.Estado}", hdrCourierBold8, brBlack, ToInch(164), pyTransport + rowHeight + spcToStartRow2);

         g.DrawLine(pen2, ToInch(169), pyTransport + rowHeight, ToInch(169), pyTransport + rowHeight2);
         g.DrawString("INSCRIÇÃO ESTADUAL", hdrRegular5, brBlack, ToInch(170), pyTransport + rowHeight + spcToStartRow);
         g.DrawString($"{transp?.Inscricao}", hdrCourierBold8, brBlack, ToInch(170), pyTransport + rowHeight + spcToStartRow2);

         g.DrawLine(pen2, px1, pyTransport + rowHeight2, px2, pyTransport + rowHeight2);

         g.DrawString("QUANTIDADE", hdrRegular5, brBlack, px1, pyTransport + rowHeight2 + spcToStartRow);
         g.DrawString($"{ds.VolumeQuantidade}", hdrCourierBold8, brBlack, px1 + ToInch(3), pyTransport + rowHeight2 + spcToStartRow2);

         var rcVolEsp = new RectangleF(ToInch(42), pyTransport + rowHeight2 + spcToStartRow2, ToInch(25), hdrCourierBold8.GetHeight(g));
         g.DrawLine(pen2, ToInch(41), pyTransport + rowHeight2, ToInch(41), pyTransport + rowHeight3);
         g.DrawString("ESPÉCIE", hdrRegular5, brBlack, ToInch(42), pyTransport + rowHeight2 + spcToStartRow);
         g.DrawString($"{ds.VolumeEspecie}", hdrCourierBold8, brBlack, rcVolEsp, ctFormat);

         var rcVolMarca = new RectangleF(ToInch(71), pyTransport + rowHeight2 + spcToStartRow2, ToInch(22), hdrCourierBold8.GetHeight(g));
         g.DrawLine(pen2, ToInch(70), pyTransport + rowHeight2, ToInch(70), pyTransport + rowHeight3);
         g.DrawString("MARCA", hdrRegular5, brBlack, ToInch(71), pyTransport + rowHeight2 + spcToStartRow);
         g.DrawString($"{ds.VolumeMarca}", hdrCourierBold8, brBlack, rcVolMarca, ctFormat);

         g.DrawLine(pen2, ToInch(94), pyTransport + rowHeight2, ToInch(94), pyTransport + rowHeight3);
         g.DrawString("NÚMERO", hdrRegular5, brBlack, ToInch(95), pyTransport + rowHeight2 + spcToStartRow);
         g.DrawString($"{ds.VolumeNumero}", hdrCourierBold8, brBlack, ToInch(96), pyTransport + rowHeight2 + spcToStartRow2);

         g.DrawLine(pen2, ToInch(141), pyTransport + rowHeight2, ToInch(141), pyTransport + rowHeight3);
         g.DrawString("PESO BRUTO", hdrRegular5, brBlack, ToInch(142), pyTransport + rowHeight2 + spcToStartRow);
         g.DrawString($"{ds.PesoBruto:###,###,##0.000}", hdrCourierBold8, brBlack, ToInch(168), pyTransport + rowHeight2 + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, ToInch(169), pyTransport + rowHeight2, ToInch(169), pyTransport + rowHeight3);
         g.DrawString("PESO LÍQUIDO", hdrRegular5, brBlack, ToInch(170), pyTransport + rowHeight2 + spcToStartRow);
         g.DrawString($"{ds.PesoLiquido:###,###,##0.000}", hdrCourierBold8, brBlack, ToInch(204), pyTransport + rowHeight2 + spcToStartRow2, rtFormat);

         g.DrawLine(pen2, px1, pyTransport + rowHeight3, px2, pyTransport + rowHeight3);

         g.DrawString("DADOS DOS PRODUTOS / SERVIÇOS", hdrRegular6, brBlack, px1, secProduct);
         g.DrawLine(pen2, px1, pyProduct, px2, pyProduct);

         g.DrawString("COD. PROD.", hdrRegular5, brBlack, px1, pyProduct + ToInch(1.5));
         g.DrawString("DESCRIÇÃO DOS PRODUTOS / SERVIÇOS", hdrRegular5, brBlack, ToInch(22), pyProduct + ToInch(1.5));
         g.DrawString("NCM / SH", hdrRegular5, brBlack, ToInch(72), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("CST", hdrRegular5, brBlack, ToInch(83.5), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("CFOP", hdrRegular5, brBlack, ToInch(91), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("UNID.", hdrRegular5, brBlack, ToInch(98.5), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("QUANT.", hdrRegular5, brBlack, ToInch(110), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("VALOR UNITÁRIO", hdrRegular5, brBlack, ToInch(126), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("VALOR TOTAL", hdrRegular5, brBlack, ToInch(143), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("B. CALC. ICMS", hdrRegular5, brBlack, ToInch(160), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("VALOR ICMS", hdrRegular5, brBlack, ToInch(175), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("VALOR IPI", hdrRegular5, brBlack, ToInch(188), pyProduct + ToInch(1.5), ctFormat);
         g.DrawString("ALÍQUOTAS", hdrRegular4, brBlack, ToInch(199.5), pyProduct + ToInch(1), ctFormat);
         g.DrawString("ICMS", hdrRegular4, brBlack, ToInch(195.5), pyProduct + ToInch(3));
         g.DrawString("IPI", hdrRegular4, brBlack, ToInch(201), pyProduct + ToInch(3));

         g.DrawLine(Pens.Black, ToInch(194), pyProduct + ToInch(2.5), px2, pyProduct + ToInch(2.5));
         g.DrawLine(pen2, px1, pyProduct + ToInch(5), px2, pyProduct + ToInch(5));

         // linhas verticais
         float pyEndProduct = secIssqn - sectionHeight;
         g.DrawLine(pen2, px1 + ToInch(16), pyProduct, px1 + ToInch(16), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(59), pyProduct, px1 + ToInch(59), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(75), pyProduct, px1 + ToInch(75), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(81.5), pyProduct, px1 + ToInch(81.5), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(90), pyProduct, px1 + ToInch(90), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(96.5), pyProduct, px1 + ToInch(96.5), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(112), pyProduct, px1 + ToInch(112), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(129.5), pyProduct, px1 + ToInch(129.5), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(146.5), pyProduct, px1 + ToInch(146.5), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(163.5), pyProduct, px1 + ToInch(163.5), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(177), pyProduct, px1 + ToInch(177), pyEndProduct);
         g.DrawLine(pen2, px1 + ToInch(189), pyProduct, px1 + ToInch(189), pyEndProduct);
         g.DrawLine(Pens.Black, px1 + ToInch(195), pyProduct + ToInch(2.5), px1 + ToInch(195), pyEndProduct);

         float pyCurItem = pyItems;
         float fontSz4 = hdrCourierBold4.GetHeight(g);
         float fontSz7 = hdrCourierBold7.GetHeight(g);

         for (; record < ds.ItemsA.Count; )
         {
            var item = ds.ItemsA.GetItem(record);
            ncmMask.Set(item.NCM);

            if (!itemPrinted)
            {
               var rcProdNome = new RectangleF(px1 + ToInch(17), pyCurItem, ToInch(41), fontSz7);

               g.DrawString($"{item.Produto:0'.'00'.'0000}", hdrCourierBold6, brBlack, px1, pyCurItem);
               g.DrawString($"{item.NomeProduto}", hdrCourierBold6, brBlack, rcProdNome);
               g.DrawString($"{ncmMask}", hdrCourierBold6, brBlack, px1 + ToInch(60), pyCurItem);
               g.DrawString($"{item.CstOrigem}{item.STICMS:00}", hdrCourierBold6, brBlack, px1 + ToInch(76), pyCurItem);
               g.DrawString($"{ds.CFOP}", hdrCourierBold6, brBlack, px1 + ToInch(82.5), pyCurItem);
               g.DrawString($"{item.Unidade}", hdrCourierBold6, brBlack, px1 + ToInch(91), pyCurItem);
               g.DrawString($"{item.Quantidade:###,##0.000}", hdrCourierBold6, brBlack, px1 + ToInch(111.5), pyCurItem, rtFormat);
               g.DrawString($"{item.PrecoUnitario:###,##0.00}", hdrCourierBold6, brBlack, px1 + ToInch(129), pyCurItem, rtFormat);
               g.DrawString($"{(double)item.PrecoUnitario * item.Quantidade:###,##0.00}", hdrCourierBold6, brBlack, px1 + ToInch(146), pyCurItem, rtFormat);
               g.DrawString($"{item.BaseCalculoICMS:###,##0.00}", hdrCourierBold6, brBlack, px1 + ToInch(163), pyCurItem, rtFormat);
               g.DrawString($"{item.ValorICMS:###,##0.00}", hdrCourierBold6, brBlack, px1 + ToInch(176.5), pyCurItem, rtFormat);
               g.DrawString($"{item.ValorIPI:###,##0.00}", hdrCourierBold6, brBlack, px1 + ToInch(188.5), pyCurItem, rtFormat);
               g.DrawString($"{item.AliquotaICMS:##0.#}", hdrCourierBold6, brBlack, px1 + ToInch(194.5), pyCurItem, rtFormat);
               g.DrawString($"{item.AliquotaIPI:##0.#}", hdrCourierBold6, brBlack, px2, pyCurItem, rtFormat);

               pyCurItem += fontSz7;
               itemPrinted = true;

               if (pyCurItem >= pyEndProduct)
                  break;
            }

            bool nextPage = false;

            if (!lotePrinted)
            {
               var lotes = ds.Lotes.Where(a => a.ItemFatura == item.Item).ToList();
               
               for (; curLote < lotes.Count; curLote++)
               {
                  var l = lotes[curLote];
                  g.DrawString($"Lote: {l.Lote} Val.: {l.DataValidade:dd/MM/yy} Qtde: ", hdrCourierBold4, brBlack, px1 + ToInch(17), pyCurItem);
                  g.DrawString($"{l.Quantidade:###,###0.000}", hdrCourierBold4, brBlack, px1 + ToInch(58), pyCurItem, rtFormat);

                  pyCurItem += fontSz4 * 1.15f;

                  if (pyCurItem >= pyEndProduct)
                  {
                     nextPage = true;
                     break;
                  }
               }
            }

            if (nextPage) break;
            lotePrinted = true;

            if (!fciPrinted && !string.IsNullOrWhiteSpace(item.NumeroFCI))
            {
               g.DrawString($"FCI: {item.NumeroFCI}", hdrCourierBold4, brBlack, px1 + ToInch(17), pyCurItem);
               fciPrinted = true;
               pyCurItem += fontSz4 * 1.15f;

               if (pyCurItem >= pyEndProduct)
                  break;
            }

            if (!fecpPrinted && ds.CFOP.StartsWith("5"))
            {
               fecpPrinted = true;
               pyCurItem += fontSz4 * 1.15f;

               if (pyCurItem >= pyEndProduct)
                  break;
            }

            if (!ctrlPrinted && item.Controlado)
            {
               g.DrawString("PRODUTO CONTROLADO PELA POLÍCIA FEDERAL", hdrCourierBold4, brBlack, px1 + ToInch(17), pyCurItem);
               pyCurItem += fontSz4 * 1.15f;

               if (pyCurItem >= pyEndProduct)
                  break;
            }

            pyCurItem += spcToStartRow;

            float szItem = GetSizeItem(item, g);
            
            if (pyCurItem + szItem >= pyEndProduct)
               break;

            itemPrinted = false;
            lotePrinted = false;
            fciPrinted = false;
            fecpPrinted = false;
            ctrlPrinted = false;
            curLote = 0;
            record++;
         }

         // ISSQN
         g.DrawLine(pen2, px1, secIssqn - sectionHeight, px2, secIssqn - sectionHeight);

         g.DrawString("CÁLCULO DO ISSQN", hdrRegular6, brBlack, px1, secIssqn);
         g.DrawLine(pen2, px1, pyIssqnHdr, px2, pyIssqnHdr);

         g.DrawString("INSCRIÇÃO MUNICIPAL", hdrRegular5, brBlack, px1, pyIssqnHdr + spcToStartRow);

         g.DrawLine(pen2, ToInch(64), pyIssqnHdr, ToInch(64), pyIssqnHdr + rowHeight);
         g.DrawString("VALOR TOTAL DOS SERVIÇOS", hdrRegular5, brBlack, ToInch(65), pyIssqnHdr + spcToStartRow);

         g.DrawLine(pen2, ToInch(120), pyIssqnHdr, ToInch(120), pyIssqnHdr + rowHeight);
         g.DrawString("BASE DE CÁLCULO DO ISSQN", hdrRegular5, brBlack, ToInch(121), pyIssqnHdr + spcToStartRow);

         g.DrawLine(pen2, ToInch(170), pyIssqnHdr, ToInch(170), pyIssqnHdr + rowHeight);
         g.DrawString("VALOR DO ISSQN", hdrRegular5, brBlack, ToInch(171), pyIssqnHdr + spcToStartRow);

         g.DrawLine(pen2, px1, pyIssqnHdr + rowHeight, px2, pyIssqnHdr + rowHeight);

         // INFO ADICIONAL

         g.DrawString("DADOS ADICIONAIS", hdrRegular6, brBlack, px1, secInfoAdic);
         g.DrawLine(pen2, px1, pyInfoAdic, px2, pyInfoAdic);

         var rcInfoAdic = new RectangleF(ToInch(8), pyInfoAdic + spcToStartRow2, ToInch(128), ToInch(25));
         g.DrawString("INFORMAÇÕES COMPLEMENTARES", hdrRegular5, brBlack, px1, pyInfoAdic + spcToStartRow);
         g.DrawString($"{ds.Observacao.Replace("|", "\n\r")}", hdrCourierBold7, brBlack, rcInfoAdic);

         g.DrawLine(pen2, ToInch(138), pyInfoAdic, ToInch(138), pyEndPage);
         g.DrawString("RESERVADO AO FISCO", hdrRegular6, brBlack, ToInch(139), pyInfoAdic + spcToStartRow);

         g.DrawLine(pen2, px1, pyEndPage, px2, pyEndPage);

         if (record < ds.ItemsA.Count)
            AddPage(e);
      }

      #endregion

      private float GetNeededSizeItems(Graphics gr)
      {
         float fontSz4 = hdrCourierBold4.GetHeight(gr);
         float fontSz7 = hdrCourierBold7.GetHeight(gr);

         var items = ds.ItemsA.ToArray();

         int countLotes = ds.Lotes.Count;       // items.Sum(a => a.Lotes.Count);
         int countFci = items.Count(a => !string.IsNullOrWhiteSpace(a.NumeroFCI));
         int countFecp = items.Count(a => ds.CFOP.StartsWith("5"));
         int countCtrl = items.Count(a => a.Controlado);

         float szItems = fontSz7 * ds.ItemsA.Count;
         float szLotes = fontSz4 * 1.15f * countLotes;
         float szFCI = fontSz4 * 1.15f * countFci;
         float szFecp = fontSz4 * 1.15f * countFecp;
         float szCtrl = fontSz4 * 1.15f * countCtrl;
         float spaceBetweenItems = (items.Length - 1) * spcToStartRow;

         return szItems + szLotes + szFCI + szFecp + szCtrl + spaceBetweenItems;
      }

      private float GetSizeItem(FaturaItem item, Graphics gr)
      {
         float fontSz4 = hdrCourierBold4.GetHeight(gr);
         float fontSz7 = hdrCourierBold7.GetHeight(gr);
         int countSubItems = ds.Lotes.Count(a => a.ItemFatura == item.Item); // item.Lotes.Count;

         if (!string.IsNullOrWhiteSpace(item.NumeroFCI))
            countSubItems++;

         if (ds.CFOP.StartsWith("5"))
            countSubItems++;

         if (item.Controlado)
            countSubItems++;

         return fontSz7 + fontSz4 * 1.15f * countSubItems + spcToStartRow;
      }
   }
}
