using System;
using System.Drawing;

namespace EmissorNFe.Printing
{
   public enum TextJustification : int
   {
      Left,
      Center,
      Right,
      Full
   }

   internal class PrintUtility
   {
      public static int ToMilimeter(float inch)
      {
         return (int)Math.Round(inch * 25.4 / 100);
      }

      public static float ToInch(int mm)
      {
         return (float)Math.Round(mm / 25.4 * 100);
      }

      public static float ToInch(double mm)
      {
         return (float)Math.Round(mm / 25.4 * 100);
      }

      public static int ToInchI(int mm)
      {
         return (int)Math.Round(mm / 25.4 * 100);
      }

      public static int ToPixel(float mm, int ppi = 96)
      {
         return (int)(mm / 10 * (ppi / 2.54f));
      }

      // Draw justified text on the Graphics object in the indicated Rectangle.
      public static RectangleF DrawParagraphs(Graphics gr, RectangleF rect, Font font, Brush brush, string text,
          TextJustification justification, float line_spacing, float indent, float paragraph_spacing)
      {
         // split the text into paragraphs.
         string[] paragraphs = text.Split('\n');

         // draw each paragraph.
         foreach (string paragraph in paragraphs)
         {
            // draw the paragraph keeping track of remaining space.
            rect = DrawParagraph(gr, rect, font, brush, paragraph, justification,
               line_spacing, indent, paragraph_spacing);

            // see if there's any room left.
            if (rect.Height < font.Size) break;
         }

         return rect;
      }

      // Draw a paragraph by lines inside the Rectangle.
      // Return a RectangleF representing any unused space in the original RectangleF.
      public static RectangleF DrawParagraph(Graphics gr, RectangleF rect, Font font, Brush brush, string text,
          TextJustification justification, float line_spacing, float indent, float extra_paragraph_spacing)
      {
         // returns the rect if there's no text to print 
         if (string.IsNullOrWhiteSpace(text))
            return rect;
         
         // get the coordinates for the first line.
         float y = rect.Top;

         // break the text into words.
         string[] words = text.Split(' ');
         int start_word = 0;

         // repeat until we run out of text or room.
         for (;;)
         {
            // see how many words will fit.
            // start with just the next word.
            string line = words[start_word];

            // add more words until the line won't fit.
            int end_word = start_word + 1;
            while (end_word < words.Length)
            {
               // see if the next word fits.
               string test_line = line + " " + words[end_word];
               var line_size = gr.MeasureString(test_line, font);

               if (line_size.Width > (rect.Width - indent))
               {
                  // the line is too wide. Don't use the last word.
                  end_word--;
                  break;
               }
               else
               {
                  // the word fits. Save the test line.
                  line = test_line;
               }

               // try the next word.
               end_word++;
            }

            // see if this is the last line in the paragraph.
            if ((end_word == words.Length) && (justification == TextJustification.Full))
            {
               // this is the last line. Don't justify it.
               DrawLine(gr, line, font, brush, rect.Left + indent, y,
                  rect.Width - indent, TextJustification.Left);
            }
            else
            {
               // this is not the last line. Justify it.
               DrawLine(gr, line, font, brush, rect.Left + indent, y,
                   rect.Width - indent, justification);
            }

            // move down to draw the next line.
            y += font.Height * line_spacing;

            // make sure there's room for another line.
            //if (font.Size > rect.Height) break;

            // start the next line at the next word.
            start_word = end_word + 1;
            if (start_word >= words.Length) break;

            // don't indent subsequent lines in this paragraph.
            indent = 0;
         }

         // add a gap after the paragraph.
         y += font.Height * extra_paragraph_spacing;

         // return a RectangleF representing any unused space in the original RectangleF.
         float height = y - rect.Top;  //rect.Bottom - y;
         if (height < 0) height = 0;
         return new RectangleF(rect.X, y, rect.Width, height);
      }

      // Draw a line of text.
      public static void DrawLine(Graphics gr, string line, Font font, Brush brush, float x, float y, float width, TextJustification justification)
      {
         // make a rectangle to hold the text.
         var rect = new RectangleF(x, y, width, font.Height);

         // see if we should use full justification.
         if (justification == TextJustification.Full)
         {
            // justify the text.
            DrawJustifiedLine(gr, rect, font, brush, line);
         }
         else
         {
            // make a StringFormat to align the text.
            var sf = new StringFormat
            {
               // use the appropriate alignment.
               Alignment = (StringAlignment)justification
            };

            //switch (justification)
            //{
            //   case TextJustification.Left:
            //      sf.Alignment = StringAlignment.Near;
            //      break;
            //   case TextJustification.Right:
            //      sf.Alignment = StringAlignment.Far;
            //      break;
            //   case TextJustification.Center:
            //      sf.Alignment = StringAlignment.Center;
            //      break;
            //}


            gr.DrawString(line, font, brush, rect, sf);
         }
      }

      // Draw justified text on the Graphics object in the indicated Rectangle.
      public static void DrawJustifiedLine(Graphics gr, RectangleF rect, Font font, Brush brush, string text)
      {
         // break the text into words.
         string[] words = text.Split(' ');

         // add a space to each word and get their lengths.
         float[] word_width = new float[words.Length];
         float total_width = 0;

         for (int i = 0; i < words.Length; i++)
         {
            // see how wide this word is.
            var size = gr.MeasureString(words[i], font);
            word_width[i] = size.Width;
            total_width += word_width[i];
         }

         // get the additional spacing between words.
         float extra_space = rect.Width - total_width;
         int num_spaces = words.Length - 1;
         if (words.Length > 1) extra_space /= num_spaces;

         // draw the words.
         float x = rect.Left;
         float y = rect.Top;

         for (int i = 0; i < words.Length; i++)
         {
            // draw the word.
            gr.DrawString(words[i], font, brush, x, y);

            // move right to draw the next word.
            x += word_width[i] + extra_space;
         }
      }

      public static string NumberInWords(decimal number, bool isCurrency)
      {
         // se o número for menor que zero ou superior a 1 quatrilhão sai da função
         if (number < 0 || number >= 1000000000000000)
            return "Valor não suportado pelo sistema.";

         // unidades
         string[] unidades = new string[] { "zero", "um ", "dois ", "tres ", "quatro ", "cinco ", "seis ",
            "sete ", "oito ",  "nove ", "dez ", "onze ", "doze ", "treze ", "quatorze ", "quinze ",
            "dezesseis ", "dezessete ", "dezoito ", "dezenove " };

         // dezenas
         string[] dezenas = new string[] { "", "dez ", "vinte ", "trinta ", "quarenta ", "cinquenta ",
            "sessenta ", "setenta ", "oitenta ", "noventa " };

         // centenas
         string[] centenas = new string[] { "cem ", "cento ", "duzentos ", "trezentos ", "quatrocentos ",
            "quinhentos ", "seiscentos ", "setecentos ", "oitocentos ", "novecentos " };

         // milhares, milhões, etc...
         string[] milhar = new string[] { "", "mil ", "milhão ", "bilhão ", "trilhão ", "quatrilhão " };
         string[] milhares = new string[] { "", "mil ", "milhões ", "bilhões ", "trilhões ", "quatrilhões " };

         // formata o número para a variável local
         string valor = string.Format("{0:000000000000000.00}", number);

         // divide o número em grupos de 3 dígitos
         string[] grupo = new string[6];

         // define os locais para armazena o texto de cada grupo
         string[] texto = new string[6];

         for (int i = 0, j = 0; i < valor.Length; i += 3, j++)
         {
            grupo[j] = valor.Substring(i, 3).Replace(".", "").Replace(",", "");
            if (grupo[j].Length < 3) grupo[j] = "0" + grupo[j];
         }

         int g = 0;  // define o grupo em uso

         // realiza a leitura dos grupos
         foreach (var parte in grupo)
         {
            // converte o grupo em valor
            int vn = int.Parse(parte);

            // verifica o tamanho do grupo de números
            int tam = (vn == 0) ? 0 : (vn < 10) ? 1 : (vn < 100) ? 2 : 3;

            // obtém os dígitos de cada posição do número
            int c = int.Parse(parte.Substring(0, 1));
            int d = int.Parse(parte.Substring(1, 2));
            int u = int.Parse(parte.Substring(2));

            // o tamanho é 3
            if (tam == 3)
            {
               // caso os 2 últimos algarismos forem 00, trata-se de uma centena inteira
               if (parte.EndsWith("00"))
               {
                  // passa para o texto qual centena pertence o número
                  if (c == 1) c = 0;
                  texto[g] += centenas[c];
               }
               else
               {
                  // passa para o texto qual centena se refere
                  texto[g] += centenas[c] + "e ";

                  // diminui o tamanho para 2
                  tam = 2;
               }
            }

            // o tamanho é 2
            if (tam == 2)
            {
               // verifica se os 2 últimos algarismos são menores que 20
               // se positivo, informa qual unidade pertence o número
               if (vn < 20)
               {
                  // adiciona ao texto a unidade referente ao número
                  texto[g] += unidades[vn];
               }
               else
               {
                  // adicona ao texto a dezena referente ao número
                  texto[g] += dezenas[d];

                  // não é uma dezena exata
                  if (parte.Substring(1, 1) != "0")
                  {
                     // adicona ao texto a palavra 'e' e diminui o tamanho para 1
                     texto[g] += "e ";
                     tam = 1;
                  }
               }
            }

            // o tamanho é 1
            if (tam == 1)
            {
               // adicona ao texto qual unidade represena o número
               texto[g] += unidades[u];
            }

            g++;
         }

         for (int i = 0; i < 5; i++)
         {
            // converte para número
            int vg = int.Parse(grupo[i]);

            // verifica se cada grupo possui valor; se positivo
            // adiciona a palavra representativo dos grupos na forma singular ou plural
            if (vg != 0) texto[i] += (vg == 1) ? milhar[4 - i] : milhares[4 - i];
         }

         // variáveis auxiliares para ajustes finais do texto
         _ = int.TryParse(grupo[3], out int milh);
         _ = int.TryParse(grupo[4], out int unid);

         // adiciona o conectivo 'e' quando há uma milhar e
         // as unidades são menores que 100
         if (milh > 0 && unid < 100)
            texto[3] += "e ";

         // o número é um valor monetário
         if (isCurrency)
         {
            // complementa com o valor referente a moeda
            texto[4] += (number == 1) ? "real " : "reais ";

            // quando o grupo das unidades é milhares for 0, acrescenta a preposição 'de'
            if (milh == 0 && unid == 0) texto[3] += "de ";

            // se moeda, verifica os centavos
            int cent = int.Parse(grupo[5]);

            if (cent != 0)
            {
               // se é 1, então adiciona o texto no singular, senão, no plural
               texto[5] += (cent == 1) ? "centavo" : "centavos";

               // adiciona o complemento 'e' ao final do grupo das unidades
               texto[4] += "e ";
            }
         }

         // retorna o número por extenso
         return string.Join("", texto).Trim();
      }
   }
}