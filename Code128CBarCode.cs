using System;
using System.Collections.Generic;
using System.Drawing;

namespace EmissorNFe.Printing
{
   /// <summary>
   /// Generate code-128C barcodes.
   /// </summary>
   public class Code128CBarCode
   {
      /// <summary>
      /// The encoding array for all allowed characters.
      /// </summary>
      private static readonly string[] bc = new string[] {
         "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312", "132212", "221213",
         "221312", "231212", "112232", "122132", "122231", "113222", "123122", "123221", "223211", "221132",
         "221231", "213212", "223112", "312131", "311222", "321122", "321221", "312212", "322112", "322211",
         "212123", "212321", "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
         "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121", "313121", "211331",
         "231131", "213113", "213311", "213131", "311123", "311321", "331121", "312113", "312311", "332111",
         "314111", "221411", "431111", "111224", "111422", "121124", "121421", "141122", "141221", "112214",
         "112412", "122114", "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
         "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112", "421211", "212141",
         "214121", "412121", "111143", "111341", "131141", "114113", "114311", "411113", "411311", "113141",
         "114131", "311141", "411131", "211412", "211214",
         "211232",   // START C
         "2331112",  // STOP
         "211133"
      };

      /// <summary>
      /// Generate a code-128C barcode.
      /// </summary>
      /// <param name="content">Barcode input sequence.</param>
      /// <param name="errorMessage">Stores an error message if an error occurs.</param>
      /// <returns>Returns an <see cref="Image"/> of the barcode.</returns>
      public Image Create(string content, int width, int height, out string errorMessage)
      {
         if (content.Length % 2 == 1)
            content = $"0{content}";

         int calcDV = CalculateDV(content);
         
         if (calcDV == -1)
         {
            errorMessage = "Não foi possível gerar o código de barras! Erro na geração do dígito verificador.";
            return null;
         }

         //****************** Creates 128C Bar Code pattern *******************
         //int checkSum = 105;
         var barcodeArray = new List<string> { bc[105] };

         for (int m = 0; m < content.Length; m += 2)
         {
            string twoDigitString = content.Substring(m, 2);
            int twoDigitNumber = int.Parse(twoDigitString);
            //checkSum += twoDigitNumber * (m / 2 + 1);
            barcodeArray.Add(bc[twoDigitNumber]);
         }

         //int reminder = calcDV % 103;
         barcodeArray.Add(bc[calcDV]);
         barcodeArray.Add(bc[106]);
         //********************************************************************


         //************************* Prepare the image ************************
         //int dW = pObjeto.ScaleHeight / 40; // espaco entre as barras
         //pObjeto.Width = 1.1 * Len(new_string) * (15 * dW) * pObjeto.Width / pObjeto.ScaleWidth

         Image canvas = new Bitmap(width, height);
         var graphics = Graphics.FromImage(canvas);

         //tH = pObjeto.TextHeight(pSequenciaCodBar)   'altura do texto
         //tW = pObjeto.TextWidth(pSequenciaCodBar)    'largura do texto
         //string new_string = (char)1 + input + (char)2;

         //int y1 = 1;
         //int y2 = height - 2;
         //********************************************************************

         graphics.FillRectangle(Brushes.White, new Rectangle(0, 0, width, height));

         // desenha cada caractere na string do codigo de barras
         float posLeft = 0;   // pObjeto.ScaleLeft;
         float dW = Math.Max(1, height / 40); // (content.Length + 3);

         // margem clara...5 espaços no início
         graphics.FillRectangle(Brushes.White, posLeft, 0, posLeft + 5 * dW, height);
         posLeft += 5;

         // para cada padrão faça...Ex.:(0) = "211232" ... (1) = "2331112" ...
         foreach (string code in barcodeArray)
         {
            // combinação de barras: B = barra preta e S = espaço (barra branca)
            // CODE C = B S B S B S

            // para cada item do padrão faça...Ex: padrão = 211232 ... (1)= 2 ... (2) = 1 ... (3) = 1 ... (4) = 2 ...
            for (int n = 0; n < code.Length; n++)
            {
               // próximo caracter do padrão...
               string c = code.Substring(n, 1);
               float v = int.Parse(c);

               var rectangle = new RectangleF(posLeft, 0, v * dW, height - 1);
               graphics.FillRectangle((n % 2 == 0) ? Brushes.Black : Brushes.White, rectangle);
               posLeft += v * dW;
            }
         }

         // margem clara...5 espaços no final
         graphics.FillRectangle(Brushes.White, posLeft, 0, posLeft + 5 * dW, height);

         errorMessage = string.Empty;
         return canvas;
      }

      /// <summary>
      /// Calculate the check digit of the 128C barcode.
      /// </summary>
      /// <param name="input">Barcode input sequence.</param>
      /// <returns>Returns a positive integer with the code's check digit.
      /// Returns the value -1 in case of error.
      /// </returns>
      private static int CalculateDV(string input)
      {
         if (input.Length % 2 == 1)
            input = $"0{input}";

         int checkSum = 105;    // START value

         for (int m = 0; m < input.Length; m += 2)
         {
            string twoDigitString = input.Substring(m, 2);
            int twoDigitNumber = int.Parse(twoDigitString);
            checkSum += twoDigitNumber * (m / 2 + 1);
         }

         int reminder = checkSum % 103;

         //// The CODE-128C Barcode Character Set only goes up to 104 positions
         //if (reminder > 102)
         //   return -1;

         return reminder;
      }
   }
}
