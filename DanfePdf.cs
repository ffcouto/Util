using EmissorNFe.Models;
using Root.Reports;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Windows.Forms;

namespace EmissorNFe.Printing
{
   internal class DanfePdf : Report
   {
      private readonly Fatura _fatura;
      private readonly NFeEvento _evento;
      private readonly Empresa _empresa;
      private readonly Cliente _cliente;
      private readonly Transportadora _transportadora;

      //private int pages = 0;
      private int record = 0;
      private int curLote = 0;
      private bool itemPrinted = false;
      private bool lotePrinted = false;
      private bool fciPrinted = false;
      private bool fecpPrinted = false;
      private bool ctrlPrinted = false;

      // define os limites da página
      private readonly static double px1 = 7;
      private readonly static double px2 = 204;

      // define as variáveis de posicionamento e tamanho
      private readonly static double rowHeight = 6.5;
      private readonly static double rowHeight2 = rowHeight * 2;
      private readonly static double rowHeight3 = rowHeight * 3;
      private readonly static double rowHeight4 = rowHeight * 4;

      private readonly static double spcToStartRow = 2;
      private readonly static double spcToStartRow2 = 4;
      private readonly static double sectionHeight = 4;

      // define as variáveis de posicionamento fixo das seções e campos da nf
      private readonly static double pyDeliveryHdr = 9;                                         // início do canhoto
      private readonly static double dashedLineStartYPos = pyDeliveryHdr + rowHeight2 + spcToStartRow;    // linha tracejada
      private readonly static double pyInvoiceHdr = dashedLineStartYPos + spcToStartRow;        // logo + info nfe
      private readonly static double pyInvoiceHdr2 = pyInvoiceHdr + rowHeight4 + 3.5;           // operação
      private readonly static double secDest = pyInvoiceHdr2 + rowHeight2 + sectionHeight;      // seção destinatário
      private readonly static double pyDest = secDest + spcToStartRow;                          // início destinatário
      private readonly static double secDupl = pyDest + rowHeight3 + sectionHeight;             // seção faturas/duplicatas
      private readonly static double pyDupl = secDupl + spcToStartRow;                          // início faturas
      private readonly static double pyStartDupl = pyDupl + 1.5;                                // início duplicatas
      private readonly static double secIssqn = 243;                                            // seção cálculo ISSQN
      private readonly static double pyIssqnHdr = secIssqn + spcToStartRow;                     // início seção cálculo ISSQN
      private readonly static double secInfoAdic = pyIssqnHdr + rowHeight + sectionHeight;      // seção info adicional
      private readonly static double pyInfoAdic = secInfoAdic + spcToStartRow;                  // início seção info adicional
      private readonly static double pyEndPage = 290;

      // define as variáveis de posicionamento variável das seções e campos da nf
      private static int countLineDupls;           // n° de linhas para a seção faturas
      private static double pyEndDupl;             // fim duplicatas
      private static double secTax;                // seção impostos
      private static double pyTaxCalc;             // início dos tributos
      private static double secTransport;          // seção transporte
      private static double pyTransport;           // início seção transporte
      private static double secProduct;            // seção produtos/serviços
      private static double pyProduct;             // início seção produtos/serviços
      private static double pyItems;               // início dos items

      private static double szProductSection;      // altura da área dos itens da nota fiscal

      public DanfePdf(
         Fatura fatura,
         NFeEvento nfe,
         Empresa empresa,
         Cliente cliente,
         Transportadora transportadora)
      {
         _fatura = fatura;
         _evento = nfe;
         _empresa = empresa;
         _cliente = cliente;
         _transportadora = transportadora;



         // calcula o números de linhas para a seção fatura/duplicata
         countLineDupls = (_fatura.Parcelas - 1) / 3 + 1;

         // calcula a posição após a última linha da seção fatura/duplicata
         double pyLastRowDupl = pyStartDupl + countLineDupls * spcToStartRow2;

         // atribui o posicionamento das seções variáveis
         pyEndDupl = (pyLastRowDupl > pyStartDupl + rowHeight) ? pyLastRowDupl : pyStartDupl + rowHeight;

         secTax = pyEndDupl + spcToStartRow2;
         pyTaxCalc = secTax + spcToStartRow;
         secTransport = pyTaxCalc + rowHeight2 + spcToStartRow2;
         pyTransport = secTransport + spcToStartRow;
         secProduct = pyTransport + rowHeight3 + spcToStartRow2;
         pyProduct = secProduct + spcToStartRow;
         pyItems = pyProduct + 5 + spcToStartRow;

         szProductSection = secIssqn - (sectionHeight * 2) - pyItems;

         itemPrinted = false;
         lotePrinted = false;
         fciPrinted = false;
         fecpPrinted = false;
         ctrlPrinted = false;
         record = 0;
      }

      protected override void Create()
      {
         FontDef fontArial = new FontDef(this, FontDef.StandardFont.Helvetica);
         FontDef fontCourier = new FontDef(this, FontDef.StandardFont.Courier);

         FontProp hdrRegular4 = new FontPropPoint(fontArial, 3.5);
         FontProp hdrRegular5 = new FontPropPoint(fontArial, 4.5);
         FontProp hdrRegular6 = new FontPropPoint(fontArial, 5.5);
         FontProp hdrRegular7 = new FontPropPoint(fontArial, 6.5);
         FontProp hdrRegularBold7 = new FontPropPoint(fontArial, 6.5) { bBold = true };
         FontProp hdrRegular8 = new FontPropPoint(fontArial, 7.5);
         FontProp hdrRegularBold8 = new FontPropPoint(fontArial, 7.5) { bBold = true };
         FontProp hdrCourierBold4 = new FontPropPoint(fontCourier, 4) { bBold = true };
         FontProp hdrCourierBold6 = new FontPropPoint(fontCourier, 5) { bBold = true };
         FontProp hdrCourierBold7 = new FontPropPoint(fontCourier, 6) { bBold = true };
         FontProp hdrCourierBold8 = new FontPropPoint(fontCourier, 7) { bBold = true };
         FontProp hdrCourierBold9 = new FontPropPoint(fontCourier, 8) { bBold = true };

         PenProp pen1 = new PenPropMM(this, 0.1, Color.Black);
         PenProp pen2 = new PenPropMM(this, 0.4, Color.Black);
         PenProp penDotted = new PenPropMM(this, 0.1, Color.Black, 2, 4);

         //BrushProp brBlack = new BrushProp(this, Color.Black);

         MaskedTextProvider cnpjMask = new MaskedTextProvider("00,000,000/0000-00");
         MaskedTextProvider cpfMask = new MaskedTextProvider("000,000,000-00");
         MaskedTextProvider ncmMask = new MaskedTextProvider("0000,00,00");

         Page page = new Page(this);

         // canhoto recebimento
         page.AddMM(px1, pyDeliveryHdr, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyDeliveryHdr + spcToStartRow, new RepString(hdrRegular5,
            $"RECEBEMOS DE {_empresa.Nome.ToUpper()}, OS PRODUTOS E/OU SERVIÇOS " +
            "CONSTANTES DA NOTA FISCAL ELETRÔNICA INDICADO AO LADO"));

         page.AddMM(px1, pyDeliveryHdr + rowHeight, new RepLineMM(pen2, 171.5 - px1, 0));

         page.AddMM(px1, pyDeliveryHdr + rowHeight + spcToStartRow, new RepString(hdrRegular5, "DATA DE RECEBIMENTO"));
         page.AddMM(41, pyDeliveryHdr + rowHeight + spcToStartRow, new RepString(hdrRegular5, "IDENTIFICAÇÃO E ASSINATURA DO RECEBEDOR"));

         page.AddMM(px1, pyDeliveryHdr + rowHeight2, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(40, pyDeliveryHdr + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(171.5, pyDeliveryHdr, new RepLineMM(pen2, 0, -rowHeight2));

         page.AddMM(186, pyDeliveryHdr + spcToStartRow * 1.25, new RepString(hdrRegular7, "NF-e"));
         page.AddMM(174, pyDeliveryHdr + spcToStartRow * 3.5, new RepString(hdrRegular7, "N°"));
         page.AddMM(184, pyDeliveryHdr + spcToStartRow * 3.5, new RepString(hdrCourierBold9, $"{_fatura.NotaFiscal:###,###,###}"));
         page.AddMM(174, pyDeliveryHdr + spcToStartRow * 5.5, new RepString(hdrRegular7, "SÉRIE"));
         page.AddMM(184, pyDeliveryHdr + spcToStartRow * 5.5, new RepString(hdrCourierBold9, $"{_fatura.SerieNF:000}"));
         // fim - canhoto recebimento

         // linha pontilhada
         page.AddMM(px1, dashedLineStartYPos, new RepLineMM(penDotted, px2 - px1, 0));

         // cabeçalho da nota
         Bitmap imgLogo = Properties.Resources.logoDanfe;
         MemoryStream msLogo = new MemoryStream();
         imgLogo.Save(msLogo, imgLogo.RawFormat);
         
         page.AddLT_MM(10.5, 30, new RepImageMM(msLogo, double.NaN, double.NaN));

         page.AddCT_MM(62, 29, new RepString(hdrRegularBold8, $"{_empresa.Nome}"));
         page.AddCT_MM(62, 35, new RepString(hdrRegular6, $"{_empresa.Endereco.Logradouro}, {_empresa.Endereco.Numero} {_empresa.Endereco.Complemento}"));
         page.AddCT_MM(62, 39, new RepString(hdrRegular6, $"{_empresa.Endereco.Bairro} - {_empresa.Endereco.Cidade}/{_empresa.Endereco.Estado}"));
         page.AddCT_MM(62, 43, new RepString(hdrRegular6, $"{_empresa.Endereco.Cep.Insert(5, "-")}   {_empresa.Telefone.PhoneNumber()}"));

         // linhas horizontais
         page.AddMM(92, pyInvoiceHdr, new RepLineMM(pen2, px2 - 92, 0));
         page.AddMM(125, pyInvoiceHdr + rowHeight2, new RepLineMM(pen2, px2 - 125, 0));
         page.AddMM(125, pyInvoiceHdr + rowHeight3, new RepLineMM(pen2, px2 - 125, 0));
         page.AddMM(px1, pyInvoiceHdr2, new RepLineMM(pen2, px2 - px1, 0));

         // linhas verticais
         page.AddMM(92, pyInvoiceHdr, new RepLineMM(pen2, 0, pyInvoiceHdr - pyInvoiceHdr2));
         page.AddMM(125, pyInvoiceHdr, new RepLineMM(pen2, 0, pyInvoiceHdr - pyInvoiceHdr2));

         //var rcBarCode = new RectangleF(ToInch(126), pyInvoiceHdr + 5, ToInch(77), ToInch(11));

         Image barCode = new Code128CBarCode().Create(_evento.NFeId, 290, 43, out _);

         MemoryStream msBarCode = new MemoryStream();
         barCode.Save(msBarCode, ImageFormat.Jpeg);

         page.AddLT_MM(127, pyInvoiceHdr + 1, new RepImageMM(msBarCode, double.NaN, double.NaN));

         page.AddCT_MM(108.5, pyInvoiceHdr + 1.5, new RepString(hdrRegular7, "DANFE"));
         page.AddCT_MM(108.5, pyInvoiceHdr + 4.5, new RepString(hdrRegular6, "DOCUMENTO AUXILIAR"));
         page.AddCT_MM(108.5, pyInvoiceHdr + 7.0, new RepString(hdrRegular6, "DA NOTA FISCAL"));
         page.AddCT_MM(108.5, pyInvoiceHdr + 9.5, new RepString(hdrRegular6, "ELETRÔNICA"));

         page.AddMM(94.5, pyInvoiceHdr + 13.5, new RepString(hdrRegular5, "0 - ENTRADA"));
         page.AddMM(94.5, pyInvoiceHdr + 16.0, new RepString(hdrRegular5, "1 - SAÍDA"));

         page.AddMM(116, pyInvoiceHdr + 12, new RepRectMM(pen1, 6, -6));
         page.AddCC_MM(119, pyInvoiceHdr + 15, new RepString(hdrCourierBold9, $"{_fatura.Finalidade}"));

         page.AddMM(94.5, pyInvoiceHdr + 20.5, new RepString(hdrRegular5, "N.°:"));
         page.AddCT_MM(113, pyInvoiceHdr + 19.5, new RepString(hdrCourierBold8, $"{_fatura.NotaFiscal:###,###,###}"));
         page.AddMM(94.5, pyInvoiceHdr + 23.5, new RepString(hdrRegular5, "SÉRIE:"));
         page.AddCT_MM(113, pyInvoiceHdr + 22.5, new RepString(hdrCourierBold8, $"{_fatura.SerieNF:000}"));
         page.AddMM(94.5, pyInvoiceHdr + 26.5, new RepString(hdrRegular5, "FOLHA:"));
         page.AddCT_MM(113, pyInvoiceHdr + 25.5, new RepString(hdrCourierBold8, $"{page.iPageNo}/{iPageCount}"));

         page.AddMM(126.5, pyInvoiceHdr + 13 + spcToStartRow, new RepString(hdrRegular5, "CHAVE DE ACESSO"));
         page.AddMM(130, pyInvoiceHdr + 16 + spcToStartRow, new RepString(hdrRegular8, $"{_evento.NFeId}"));

         page.AddCT_MM(165, pyInvoiceHdr + 22.5, new RepString(hdrRegularBold7, "Consulta de autenticidade no portal nacional da NF-e"));
         page.AddCT_MM(165, pyInvoiceHdr + 25.5, new RepString(hdrRegularBold7, "www.nfe.fazenda.gov.br/portal ou no site da Sefaz Autorizadora"));

         page.AddMM(px1, pyInvoiceHdr2 + spcToStartRow, new RepString(hdrRegular5, "NATUREZA DA OPERAÇÃO"));
         page.AddMM(px1 + 3, pyInvoiceHdr2 + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_fatura.CFOP} {_fatura.CfopDescricao}"));

         page.AddMM(125, pyInvoiceHdr2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(126, pyInvoiceHdr2 + spcToStartRow, new RepString(hdrRegular5, "PROTOCOLO DE AUTORIZAÇÃO DE USO"));
         page.AddMM(135, pyInvoiceHdr2 + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_evento.Recibo} {_evento.DataEvento:dd/MM/yyyy HH:mm:ss}"));

         page.AddMM(px1, pyInvoiceHdr2 + rowHeight, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyInvoiceHdr2 + rowHeight + spcToStartRow, new RepString(hdrRegular5, "INSCRIÇÃO ESTADUAL"));
         page.AddMM(px1 + 3, pyInvoiceHdr2 + rowHeight + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_empresa.InscricaoEstadual}"));

         page.AddMM(75, pyInvoiceHdr2 + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(76, pyInvoiceHdr2 + rowHeight + spcToStartRow, new RepString(hdrRegular5, "INSCRIÇÃO ESTADUAL DO SUBST. TRIBUT."));

         cnpjMask.Set(_empresa.CNPJ);

         page.AddMM(147, pyInvoiceHdr2 + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(148, pyInvoiceHdr2 + rowHeight + spcToStartRow, new RepString(hdrRegular5, "CNPJ"));
         page.AddMM(150, pyInvoiceHdr2 + rowHeight + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{cnpjMask}"));

         page.AddMM(px1, pyInvoiceHdr2 + rowHeight2, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, secDest + 1, new RepString(hdrRegular6, "DESTINATÁRIO / REMETENTE"));
         page.AddMM(px1, pyDest, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyDest + spcToStartRow, new RepString(hdrRegular5, "NOME / RAZÃO SOCIAL"));
         page.AddMM(px1 + 3, pyDest + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_cliente.Nome}"));

         cnpjMask.Set(_cliente.CNPJ);
         cpfMask.Set(_cliente.CNPJ);

         string cliPIN = _fatura.CFOP.StartsWith("7") ? _cliente.CNPJ :
            _cliente.CNPJ.Length == 14 ? $"{cnpjMask}" : $"{cpfMask}";

         page.AddMM(130, pyDest, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(131, pyDest + spcToStartRow, new RepString(hdrRegular5, "C.N.P.J. / C.P.F."));
         page.AddMM(131, pyDest + spcToStartRow2 + 1, new RepString(hdrCourierBold8, cliPIN));

         page.AddMM(174, pyDest, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(175, pyDest + spcToStartRow, new RepString(hdrRegular5, "DATA DA EMISSÃO"));
         page.AddMM(182, pyDest + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_fatura.DataFatura:dd/MM/yyyy}"));

         page.AddMM(px1, pyDest + rowHeight, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyDest + rowHeight + spcToStartRow, new RepString(hdrRegular5, "ENDEREÇO"));
         page.AddMM(px1 + 3, pyDest + rowHeight + spcToStartRow2 + 1,
            new RepString(hdrCourierBold8, $"{_cliente.Endereco.Logradouro}, {_cliente.Endereco.Numero}"));

         page.AddMM(108, pyDest + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(109, pyDest + rowHeight + spcToStartRow, new RepString(hdrRegular5, "BAIRRO / DISTRITO"));
         page.AddMM(110, pyDest + rowHeight + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_cliente.Endereco.Bairro}"));

         page.AddMM(154, pyDest + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(155, pyDest + rowHeight + spcToStartRow, new RepString(hdrRegular5, "CEP"));
         page.AddMM(156.5, pyDest + rowHeight + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_cliente.Endereco.Cep.Insert(5, "-")}"));

         page.AddMM(174, pyDest + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(175, pyDest + rowHeight + spcToStartRow, new RepString(hdrRegular5, "DATA DA ENTRADA / SAÍDA"));
         page.AddMM(182, pyDest + rowHeight + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_fatura.DataFatura:dd/MM/yyyy}"));

         page.AddMM(px1, pyDest + rowHeight2, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyDest + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "MUNICÍPIO"));
         page.AddMM(px1 + 3, pyDest + rowHeight2 + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_cliente.Endereco.Cidade}"));

         var cliContato = _cliente.Contatos.ToArray().FirstOrDefault();

         page.AddMM(82, pyDest + rowHeight2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(83, pyDest + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "FONE / FAX"));
         page.AddMM(84, pyDest + rowHeight2 + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{cliContato?.Telefone.PhoneNumber()}"));

         page.AddMM(120, pyDest + rowHeight2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(121, pyDest + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "UF"));
         page.AddMM(122, pyDest + rowHeight2 + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_cliente.Endereco.Estado}"));

         page.AddMM(130, pyDest + rowHeight2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(131, pyDest + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "INSCRIÇÃO ESTADUAL"));
         page.AddMM(131, pyDest + rowHeight2 + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_cliente.Inscricao}"));

         page.AddMM(174, pyDest + rowHeight2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(175, pyDest + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "HORA DA SAÍDA"));
         page.AddMM(182, pyDest + rowHeight2 + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_fatura.DataFatura:HH:mm:ss}"));

         page.AddMM(px1, pyDest + rowHeight3, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, secDupl + 1, new RepString(hdrRegular6, "FATURA / DUPLICATA"));
         page.AddMM(px1, pyDupl, new RepLineMM(pen2, px2 - px1, 0));

         int dRow = 0;
         int dCol = 0;
         double pyRowDupl;

         foreach (FaturaDuplicata d in _fatura.Duplicatas)
         {
            if (int.Parse(d.Numero.Substring(0, 2)) != _fatura.Empresa) continue;

            if (dCol > 2)
            {
               dRow++;
               dCol = 0;
            }

            double pxDoc = 64 * dCol + 10;
            double pxData = 64 * dCol + 30;
            double pxVlr = 64 * dCol + 69;
            pyRowDupl = pyStartDupl + dRow * spcToStartRow2;

            page.AddLT_MM(pxDoc, pyRowDupl, new RepString(hdrCourierBold8, $"{d.Numero}"));
            page.AddLT_MM(pxData, pyRowDupl, new RepString(hdrCourierBold8, $"{d.Vencimento:dd/MM/yyyy}"));
            page.AddRT_MM(pxVlr, pyRowDupl, new RepString(hdrCourierBold8, $"{d.Valor:###,###,##0.00}"));

            dCol++;
         }

         page.AddMM(px1, pyEndDupl, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, secTax + 1, new RepString(hdrRegular6, "CÁLCULO DO IMPOSTO"));
         page.AddMM(px1, pyTaxCalc, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyTaxCalc + spcToStartRow, new RepString(hdrRegular5, "BASE DE CÁLCULO DO ICMS"));
         page.AddRC_MM(40, pyTaxCalc + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.BaseCalculoICMS:###,###,##0.00}"));

         page.AddMM(54, pyTaxCalc, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(55, pyTaxCalc + spcToStartRow, new RepString(hdrRegular5, "VALOR DO ICMS"));
         page.AddRC_MM(87, pyTaxCalc + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.ValorIcms:###,###,##0.00}"));

         page.AddMM(93, pyTaxCalc, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(94, pyTaxCalc + spcToStartRow, new RepString(hdrRegular5, "BASE DE CÁLCULO DO I.C.M.S. S.T."));
         //page.AddRT_MM(124, pyTaxCalc + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.BaseCalculoICMSST:###,###,##0.00}"));

         page.AddMM(130, pyTaxCalc, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(131, pyTaxCalc + spcToStartRow, new RepString(hdrRegular5, "VALOR DO I.C.M.S. SUBSTITUIÇÃO"));
         //page.AddRT_MM(168, pyTaxCalc + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.TotalICMSST:###,###,##0.00}"));

         page.AddMM(174, pyTaxCalc, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(175, pyTaxCalc + spcToStartRow, new RepString(hdrRegular5, "VALOR TOTAL DOS PRODUTOS"));
         page.AddRC_MM(204, pyTaxCalc + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.ValorNF - _fatura.ValorIpi:###,###,##0.00}"));

         page.AddMM(px1, pyTaxCalc + rowHeight, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyTaxCalc + rowHeight + spcToStartRow, new RepString(hdrRegular5, "VALOR DO FRETE"));
         page.AddRC_MM(40, pyTaxCalc + rowHeight + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.ValorFrete:###,###,##0.00}"));

         page.AddMM(46, pyTaxCalc + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(47, pyTaxCalc + rowHeight + spcToStartRow, new RepString(hdrRegular5, "VALOR DO SEGURO"));
         //page.AddRT_MM(75, pyTaxCalc + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.ValorSeguro:###,###,##0.00}"));

         page.AddMM(81, pyTaxCalc + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(82, pyTaxCalc + rowHeight + spcToStartRow, new RepString(hdrRegular5, "DESCONTO"));
         page.AddRC_MM(107, pyTaxCalc + rowHeight + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.ValorDesconto:###,###,##0.00}"));

         page.AddMM(109, pyTaxCalc + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(110, pyTaxCalc + rowHeight + spcToStartRow, new RepString(hdrRegular5, "OUTRAS DESPESAS ACESSÓRIAS"));
         //page.AddRT_MM(138, pyTaxCalc + rowHeight + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.TotalOutrasDespesas:###,###,##0.00}"));

         page.AddMM(141, pyTaxCalc + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(142, pyTaxCalc + rowHeight + spcToStartRow, new RepString(hdrRegular5, "VALOR TOTAL DO IPI"));
         page.AddRC_MM(171, pyTaxCalc + rowHeight + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.ValorIpi:###,###,##0.00}"));

         page.AddMM(174, pyTaxCalc + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(175, pyTaxCalc + rowHeight + spcToStartRow, new RepString(hdrRegular5, "VALOR TOTAL DA NOTA"));
         page.AddRC_MM(204, pyTaxCalc + rowHeight + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.ValorNF:###,###,##0.00}"));

         page.AddMM(px1, pyTaxCalc + rowHeight2, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, secTransport + 1, new RepString(hdrRegular6, "TRANSPORTADOR / VOLUMES TRANSPORTADOS"));
         page.AddMM(px1, pyTransport, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyTransport + spcToStartRow, new RepString(hdrRegular5, "NOME / RAZÃO SOCIAL"));
         page.AddMM(px1 + 3, pyTransport + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_transportadora?.Nome}"));

         var modFrete = ModalidadeFrete.GetItem(_fatura.ModalidadeFrete);
         page.AddMM(94, pyTransport, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(95, pyTransport + spcToStartRow, new RepString(hdrRegular5, "FRETE POR CONTA"));
         page.AddMM(96, pyTransport + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{modFrete.Codigo:0}. {modFrete.Responsavel}"));

         page.AddMM(125, pyTransport, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(126, pyTransport + spcToStartRow, new RepString(hdrRegular5, "CÓDIGO ANTT"));

         page.AddMM(141, pyTransport, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(142, pyTransport + spcToStartRow, new RepString(hdrRegular5, "PLACA DO VEÍCULO"));
         page.AddMM(144, pyTransport + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_fatura.Placa.Insert(3, "-")}"));

         page.AddMM(162, pyTransport, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(163, pyTransport + spcToStartRow, new RepString(hdrRegular5, "UF"));
         page.AddMM(164, pyTransport + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_fatura.PlacaUF}"));

         string transpCnpj = _transportadora?.CNPJ ?? "";
         cnpjMask.Set(transpCnpj);

         page.AddMM(169, pyTransport, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(170, pyTransport + spcToStartRow, new RepString(hdrRegular5, "CNPJ"));
         page.AddMM(170, pyTransport + spcToStartRow2 + 1, new RepString(hdrCourierBold8, string.IsNullOrWhiteSpace(transpCnpj) ? "" : $"{cnpjMask}"));

         page.AddMM(px1, pyTransport + rowHeight, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyTransport + rowHeight + spcToStartRow, new RepString(hdrRegular5, "ENDEREÇO"));
         page.AddMM(px1 + 3, pyTransport + rowHeight + spcToStartRow2 + 1,
            new RepString(hdrCourierBold8, _transportadora != null && !_transportadora.Endereco.IsEmpty() ?
               $"{_transportadora?.Endereco.Logradouro}, {_transportadora?.Endereco.Numero} {_transportadora?.Endereco.Complemento}" : ""));

         page.AddMM(94, pyTransport + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(95, pyTransport + rowHeight + spcToStartRow, new RepString(hdrRegular5, "MUNICÍPIO"));
         page.AddMM(95, pyTransport + rowHeight + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_transportadora?.Endereco.Cidade}"));

         page.AddMM(162, pyTransport + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(163, pyTransport + rowHeight + spcToStartRow, new RepString(hdrRegular5, "UF"));
         page.AddMM(164, pyTransport + rowHeight + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_transportadora?.Endereco.Estado}"));

         page.AddMM(169, pyTransport + rowHeight, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(170, pyTransport + rowHeight + spcToStartRow, new RepString(hdrRegular5, "INSCRIÇÃO ESTADUAL"));
         page.AddMM(170, pyTransport + rowHeight + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_transportadora?.Inscricao}"));

         page.AddMM(px1, pyTransport + rowHeight2, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyTransport + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "QUANTIDADE"));
         page.AddMM(px1 + 3, pyTransport + rowHeight2 + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_fatura.VolumeQuantidade}"));

         page.AddMM(41, pyTransport + rowHeight2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(42, pyTransport + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "ESPÉCIE"));
         page.AddCT_MM(64.5, pyTransport + rowHeight2 + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.VolumeEspecie}"));

         page.AddMM(70, pyTransport + rowHeight2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(71, pyTransport + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "MARCA"));
         page.AddCT_MM(82, pyTransport + rowHeight2 + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.VolumeMarca}"));

         page.AddMM(94, pyTransport + rowHeight2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(95, pyTransport + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "NÚMERO"));
         page.AddMM(96, pyTransport + rowHeight2 + spcToStartRow2 + 1, new RepString(hdrCourierBold8, $"{_fatura.VolumeNumero}"));

         page.AddMM(141, pyTransport + rowHeight2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(142, pyTransport + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "PESO BRUTO"));
         page.AddRC_MM(168, pyTransport + rowHeight2 + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.PesoBruto:###,###,##0.000}"));

         page.AddMM(169, pyTransport + rowHeight2, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(170, pyTransport + rowHeight2 + spcToStartRow, new RepString(hdrRegular5, "PESO LÍQUIDO"));
         page.AddRC_MM(204, pyTransport + rowHeight2 + spcToStartRow2, new RepString(hdrCourierBold8, $"{_fatura.PesoLiquido:###,###,##0.000}"));

         page.AddMM(px1, pyTransport + rowHeight3, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, secProduct + 1, new RepString(hdrRegular6, "DADOS DOS PRODUTOS / SERVIÇOS"));
         page.AddMM(px1, pyProduct, new RepLineMM(pen2, px2 - px1, 0));

         page.AddLT_MM(px1, pyProduct + 2, new RepString(hdrRegular5, "COD. PROD."));
         page.AddLT_MM(22, pyProduct + 2, new RepString(hdrRegular5, "DESCRIÇÃO DOS PRODUTOS / SERVIÇOS"));
         page.AddCT_MM(72, pyProduct + 2, new RepString(hdrRegular5, "NCM / SH"));
         page.AddCT_MM(83.5, pyProduct + 2, new RepString(hdrRegular5, "CST"));
         page.AddCT_MM(91, pyProduct + 2, new RepString(hdrRegular5, "CFOP"));
         page.AddCT_MM(98.5, pyProduct + 2, new RepString(hdrRegular5, "UNID."));
         page.AddCT_MM(110, pyProduct + 2, new RepString(hdrRegular5, "QUANT."));
         page.AddCT_MM(126, pyProduct + 2, new RepString(hdrRegular5, "VALOR UNITÁRIO"));
         page.AddCT_MM(143, pyProduct + 2, new RepString(hdrRegular5, "VALOR TOTAL"));
         page.AddCT_MM(160, pyProduct + 2, new RepString(hdrRegular5, "B. CALC. ICMS"));
         page.AddCT_MM(175, pyProduct + 2, new RepString(hdrRegular5, "VALOR ICMS"));
         page.AddCT_MM(188, pyProduct + 2, new RepString(hdrRegular5, "VALOR IPI"));
         page.AddCT_MM(199.5, pyProduct + 1, new RepString(hdrRegular5, "ALÍQUOTAS"));
         page.AddMM(195.5, pyProduct + 4, new RepString(hdrRegular4, "ICMS"));
         page.AddMM(201, pyProduct + 4, new RepString(hdrRegular4, "IPI"));

         page.AddMM(194, pyProduct + 2.5, new RepLineMM(pen1, px2 - 194, 0));
         page.AddMM(px1, pyProduct + 5, new RepLineMM(pen2, px2 - px1, 0));

         // linhas verticais
         double pyEndProduct = secIssqn - sectionHeight;

         page.AddMM(px1 + 14, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 57, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 73, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 79.5, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 88, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 94.5, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 110, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 127.5, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 144.5, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 161.5, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 175, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));
         page.AddMM(px1 + 187, pyProduct, new RepLineMM(pen2, 0, pyProduct - pyEndProduct));

         page.AddMM(px1 + 193, pyProduct + 2.5, new RepLineMM(pen1, 0, pyProduct+2.5 - pyEndProduct));

         double pyCurItem = pyItems;
         double fontSz4 = hdrCourierBold4.rSizeMM + 1;
         double fontSz7 = hdrCourierBold7.rSizeMM + 1;

         for (; record < _fatura.ItemsA.Count;)
         {
            var item = _fatura.ItemsA.GetItem(record);
            ncmMask.Set(item.NCM);

            if (!itemPrinted)
            {
               //var rcProdNome = new RectangleF(px1 + ToInch(17), pyCurItem, ToInch(41), fontSz7);

               page.AddLT_MM(px1, pyCurItem, new RepString(hdrCourierBold6, $"{item.Produto:0'.'00'.'0000}"));
               page.AddLT_MM(px1 + 15, pyCurItem, new RepString(hdrCourierBold6, $"{item.NomeProduto}"));
               page.AddLT_MM(px1 + 58, pyCurItem, new RepString(hdrCourierBold6, $"{ncmMask}"));
               page.AddLT_MM(px1 + 74, pyCurItem, new RepString(hdrCourierBold6, $"{item.CstOrigem}{item.STICMS:00}"));
               page.AddLT_MM(px1 + 80.5, pyCurItem, new RepString(hdrCourierBold6, $"{_fatura.CFOP}"));
               page.AddLT_MM(px1 + 89, pyCurItem, new RepString(hdrCourierBold6, $"{item.Unidade}"));
               page.AddRT_MM(px1 + 109.5, pyCurItem, new RepString(hdrCourierBold6, $"{item.Quantidade:###,##0.000}"));
               page.AddRT_MM(px1 + 127, pyCurItem, new RepString(hdrCourierBold6, $"{item.PrecoUnitario:###,##0.00}"));
               page.AddRT_MM(px1 + 144, pyCurItem, new RepString(hdrCourierBold6, $"{(double)item.PrecoUnitario * item.Quantidade:###,##0.00}"));
               page.AddRT_MM(px1 + 161, pyCurItem, new RepString(hdrCourierBold6, $"{item.BaseCalculoICMS:###,##0.00}"));
               page.AddRT_MM(px1 + 174.5, pyCurItem, new RepString(hdrCourierBold6, $"{item.ValorICMS:###,##0.00}"));
               page.AddRT_MM(px1 + 186.5, pyCurItem, new RepString(hdrCourierBold6, $"{item.ValorIPI:###,##0.00}"));
               page.AddRT_MM(px1 + 192.5, pyCurItem, new RepString(hdrCourierBold6, $"{item.AliquotaICMS:##0.#}"));
               page.AddRT_MM(px2, pyCurItem, new RepString(hdrCourierBold6, $"{item.AliquotaIPI:##0.#}"));

               pyCurItem += fontSz7;
               itemPrinted = true;

               if (pyCurItem >= pyEndProduct)
                  break;
            }

            bool nextPage = false;

            if (!lotePrinted)
            {
               var lotes = _fatura.Lotes.Where(a => a.ItemFatura == item.Item).ToList();

               for (; curLote < lotes.Count; curLote++)
               {
                  var l = lotes[curLote];

                  page.AddLT_MM(px1 + 15, pyCurItem, new RepString(hdrCourierBold4, $"Lote: {l.Lote} Val.: {l.DataValidade:dd/MM/yy} Qtde: "));
                  page.AddRT_MM(px1 + 56, pyCurItem, new RepString(hdrCourierBold4, $"{l.Quantidade:###,###0.000}"));

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
               page.AddMM(px1 + 15, pyCurItem + 0.5, new RepString(hdrCourierBold4, $"FCI: {item.NumeroFCI}"));
               fciPrinted = true;
               pyCurItem += fontSz4 * 1.15f;

               if (pyCurItem >= pyEndProduct)
                  break;
            }

            if (!fecpPrinted && _fatura.CFOP.StartsWith("5"))
            {
               fecpPrinted = true;
               pyCurItem += fontSz4 * 1.15f;

               if (pyCurItem >= pyEndProduct)
                  break;
            }

            if (!ctrlPrinted && item.Controlado)
            {
               page.Add(px1 + 17, pyCurItem + 0.5, new RepString(hdrCourierBold4, "PRODUTO CONTROLADO PELA POLÍCIA FEDERAL"));
               pyCurItem += fontSz4 * 1.15f;

               if (pyCurItem >= pyEndProduct)
                  break;
            }

            pyCurItem += spcToStartRow / 2;

            double szItem = GetSizeItem(item, hdrCourierBold4, hdrCourierBold7);

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
         page.AddMM(px1, secIssqn - sectionHeight, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, secIssqn + 1, new RepString(hdrRegular6, "CÁLCULO DO ISSQN"));
         page.AddMM(px1, pyIssqnHdr, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyIssqnHdr + spcToStartRow, new RepString(hdrRegular5, "INSCRIÇÃO MUNICIPAL"));

         page.AddMM(64, pyIssqnHdr, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(65, pyIssqnHdr + spcToStartRow, new RepString(hdrRegular5, "VALOR TOTAL DOS SERVIÇOS"));

         page.AddMM(120, pyIssqnHdr, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(121, pyIssqnHdr + spcToStartRow, new RepString(hdrRegular5, "BASE DE CÁLCULO DO ISSQN"));

         page.AddMM(170, pyIssqnHdr, new RepLineMM(pen2, 0, -rowHeight));
         page.AddMM(171, pyIssqnHdr + spcToStartRow, new RepString(hdrRegular5, "VALOR DO ISSQN"));

         page.AddMM(px1, pyIssqnHdr + rowHeight, new RepLineMM(pen2, px2 - px1, 0));

         // INFO ADICIONAL
         page.AddMM(px1, secInfoAdic + 1, new RepString(hdrRegular6, "DADOS ADICIONAIS"));
         page.AddMM(px1, pyInfoAdic, new RepLineMM(pen2, px2 - px1, 0));

         page.AddMM(px1, pyInfoAdic + spcToStartRow, new RepString(hdrRegular5, "INFORMAÇÕES COMPLEMENTARES"));

         string infoCompl = $"{_fatura.Observacao.Replace("|", "\n\r")}";

         FontProp fp_BestFit = hdrCourierBold7.fontProp_GetBestFitMM(infoCompl, 125, 30, 1);
         int iStart = 0;
         double rY = pyInfoAdic + spcToStartRow2 + 1;

         while (iStart <= infoCompl.Length)
         {
            string textLine = fp_BestFit.sGetTextLineMM(infoCompl, 125, ref iStart, TextSplitMode.Line);
            page.AddMM(8, rY, new RepString(fp_BestFit, textLine));
            rY += fp_BestFit.rLineFeedMM;
         }

         page.AddMM(138, pyInfoAdic, new RepLineMM(pen2, 0, pyInfoAdic - pyEndPage));
         page.AddMM(139, pyInfoAdic + spcToStartRow, new RepString(hdrRegular6, "RESERVADO AO FISCO"));

         page.AddMM(px1, pyEndPage, new RepLineMM(pen2, px2 - px1, 0));
      }


      private double GetSizeItem(FaturaItem item, FontProp f4, FontProp f7)
      {
         double fontSz4 = f4.rSizeMM;   //GetHeight(gr);
         double fontSz7 = f7.rSizeMM;   //GetHeight(gr);
         int countSubItems = _fatura.Lotes.Count(a => a.ItemFatura == item.Item); // item.Lotes.Count;

         if (!string.IsNullOrWhiteSpace(item.NumeroFCI))
            countSubItems++;

         if (_fatura.CFOP.StartsWith("5"))
            countSubItems++;

         if (item.Controlado)
            countSubItems++;

         return fontSz7 + fontSz4 * 1.15f * countSubItems + spcToStartRow;
      }

   }
}
