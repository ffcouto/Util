using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;

namespace HiTech
{
   [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "<Pending>")]
   public static partial class Utility
   {
#pragma warning disable SYSLIB1045 // Convert to 'GeneratedRegexAttribute'.
      private static readonly Regex accentRegex = new(@"\p{M}");
#pragma warning restore SYSLIB1045 // Convert to 'GeneratedRegexAttribute'.

      #region Math extensions

      public static double Truncate(this double value, int digits)
      {
         return Truncate(value, digits, false);
      }

      public static double Truncate(this double value, int digits, bool round)
      {
         if (round)
         {
            return Math.Round(value, digits, MidpointRounding.ToEven);
         }
         else
         {
            double p = Math.Pow(10, digits);
            return Math.Truncate(value * p) / p;
         }
      }

      public static decimal Truncate(this decimal value, int digits)
      {
         return Truncate(value, digits, false);
      }

      public static decimal Truncate(this decimal value, int digits, bool round)
      {
         if (round)
         {
            return Math.Round(value, digits, MidpointRounding.ToEven);
         }
         else
         {
            decimal p = (decimal)Math.Pow(10, digits);
            return Math.Truncate(value * p) / p;
         }
      }

      public static decimal ToRound(this decimal value, int decimals)
      {
         // converte para string
         string moeda = value.ToString("0.00##################");

         // obtém o separador decimal utilizado pelo sistema
         string decSep = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;

         // verifica qual a posição da vírgula dentro da variável
         int iPos = moeda.IndexOf(decSep) + 1;

         // se não há vírgula o número é inteiro, então a posição é o tamanho da string
         if (iPos == 0) iPos = moeda.Length - 1;

         // retorna o número até a segunda casa decimal
         return Convert.ToDecimal(moeda[..(iPos + decimals)]);
      }

      public static double ToRound(this double value, int decimals)
      {
         // converte para string
         string moeda = value.ToString("0.00##################");

         // obtém o separador decimal utilizado pelo sistema
         string decSep = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;

         // verifica qual a posição da vírgula dentro da variável
         int iPos = moeda.IndexOf(decSep) + 1;

         // se não há vígula o número é inteiro, então a posição é o tamanho da string
         if (iPos == 0) iPos = moeda.Length - 1;

         // retorna o número até a segunda casa decimal
         return Convert.ToDouble(moeda[..(iPos + decimals)]);
      }

      #endregion

      #region Date and time

      public static int DeadlineFromToday(this DateTime theDate)
      {
         return (int)theDate.Date.Subtract(DateTime.Today).TotalDays;
      }

      public static DateTime GetTime(this DateTime theDate)
      {
         return new DateTime(1, 1, 1, theDate.Hour, theDate.Minute, theDate.Second);
      }

      public static DateTime FirstDateOfWeek(int year, int weekOfYear)
      {
         DateTime jan1 = new(year, 1, 1);
         int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

         // Use first Thursday in January to get first week of the year as
         // it will never be in Week 52/53
         DateTime firstThursday = jan1.AddDays(daysOffset);
         
         int firstWeek = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
         int  weekNum = weekOfYear;
         
         // As we're adding days to a date in Week 1,
         // we need to subtract 1 in order to get the right date for week #1
         if (firstWeek == 1)
            weekNum -= 1;

         // Using the first Thursday as starting week ensures that we are starting in the right year
         // then we add number of weeks multiplied with days
         DateTime result = firstThursday.AddDays(weekNum * 7);

         // Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
         return result.AddDays(-3);
      }

      public static string LocalTimeToInternetTime(this DateTime dateTime)
      {
         // Obtém o fuso horário local
         TimeSpan offset = TimeZoneInfo.Local.GetUtcOffset(dateTime);

         // Formata a data e hora no padrão ISO 8601
         return dateTime.ToString("yyyy-MM-dd'T'HH:mm:ss") +
                (offset.TotalMinutes < 0 ? "-" : "+") +
                offset.ToString("hh\\:mm");
      }

      #endregion

      #region Strings

      public static string RemoveFormat(this string input)
      {
         string formatChars = "(),-./:@";
         string result = input;

         foreach (char s in formatChars)
            result = result.Replace(s.ToString(), "");

         return result;
      }

      public static string PhoneNumber(this string number)
      {
         if (string.IsNullOrWhiteSpace(number))
            return string.Empty;

         int tam = number.Length;
         string num = number.RemoveFormat().Replace("+", "");

         if (tam < 10)
            return string.Format("{0}-{1}", num[..(tam - 4)], num[^4..]);

         else if (tam == 10)
            return string.Format("({0}){1}-{2}", num[..2], num[2..6], num[6..]);

         else if (tam == 11)
            return string.Format("({0}){1}-{2}", num[..2], num[2..7], num[7..]);

         else if (tam == 12)
            return string.Format("+{0}({1}){2}-{3}", num[..1], num[1..3], num[3..8], num[8..]);

         else if (tam == 13)
            return string.Format("+{0}({1}){2}-{3}", num[..2], num[2..4], num[4..9], num[9..]);

         else if (tam == 14)
            return string.Format("+{0}({1}){2}-{3}", num[..3], num[3..5], num[5..10], num[10..]);

         else if (tam == 15)
            return string.Format("+{0}({1}){2}-{3}", num[..3], num[3..6], num[6..11], num[11..]);

         else
            return number;
      }

      public static string RemoveAccents(this string text)
      {
         if (string.IsNullOrWhiteSpace(text))
            return text;

         string normalized = text.Normalize(NormalizationForm.FormD);
         return accentRegex.Replace(normalized, "").Normalize(NormalizationForm.FormC);
      }

      #endregion

      #region Image

      public static byte[] ImageToByteArray(Image imageIn)
      {
         var ms = new MemoryStream();
         imageIn.Save(ms, ImageFormat.Jpeg);
         return ms.ToArray();
      }

      public static byte[] ImageToByteArray(string filename)
      {
         var img = Image.FromFile(filename);
         return ImageToByteArray(img);
      }

      public static Image ByteArrayToImage(byte[] byteArrayIn)
      {
         var ms = new MemoryStream(byteArrayIn);
         return Image.FromStream(ms);
      }

      public static object ByteArrayToObject(byte[] byteArrayIn)
      {
         return ByteArrayToImage(byteArrayIn);
      }

      public static Image GetThumbnailImage(Image original, int maxSize)
      {
         double factor;

         // largura e altura da imagem original
         int origWidth = original.Width;
         int origHeight = original.Height;

         // Cálcula o melhor fator para escala, baseado na dimensão mais larga
         if (origWidth > origHeight)
            factor = (double)maxSize / origWidth;
         else
            factor = (double)maxSize / origHeight;

         // largura e altura da nova imagem
         int destWidth = (int)(origWidth * factor);
         int destHeight = (int)(origHeight * factor);

         // cria uma nova imagem
         var bmp = new Bitmap(destWidth, destHeight);

         using (var gr = Graphics.FromImage(bmp))
         {
            gr.SmoothingMode = SmoothingMode.HighQuality;
            gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
            gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
            gr.DrawImage(original, 0, 0, destWidth, destHeight);
         }

         // retorna a nova imagem
         return bmp;
      }

      #endregion

      #region Null and empty

      public static string EmptyIfZero(this int value)
      {
         if (value == 0) return string.Empty;
         return value.ToString();
      }

      public static object NullIfZero(this double value)
      {
         if (value == 0.0 || double.IsNaN(value)) return null!;
         return value;
      }

      public static object NullIfZero(this decimal value)
      {
         if (value == 0M) return null!;
         return value;
      }

      public static object NullIfZero(this int value)
      {
         if (value == 0) return null!;
         return value;
      }

      public static object NullIfZero(this long value)
      {
         if (value == 0) return null!;
         return value;
      }

      public static object NullIfZero(this short value)
      {
         if (value == 0) return null!;
         return value;
      }

      #endregion

      #region Others

      public static int LineNumber(this Exception ex)
      {
         const string lineSearch = ":line ";

         bool foundValue = false;
         int lineNumber = 0;
         int index = ex.StackTrace!.LastIndexOf(lineSearch);

         if (index != -1)
         {
            string lineNumberText = ex.StackTrace[(index + lineSearch.Length)..];
            foundValue = int.TryParse(lineNumberText, out lineNumber);
         }

         if (!foundValue)
         {
            var st = new StackTrace(ex, true);
            var frame = st.GetFrame(0)!;
            lineNumber = frame.GetFileLineNumber();
         }

         return lineNumber;
      }

      #endregion

      #region Validation docs

      /// <summary>
      /// Verifica a validade de um número de CNPJ informado.
      /// </summary>
      /// <param name="nro">Número a ser validado.</param>
      /// <returns>True se válido; caso contrário False.</returns>
      public static bool ValidarCnpj(string nro)
      {
         int[] multiplicador1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
         int[] multiplicador2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
         int soma;
         int resto;
         string digito;
         string tempCnpj;

         // remove espaços e caracteres especias; mantém somente os números
         nro = nro.Trim().Replace(".", "").Replace("-", "").Replace("/", "");

         // o número deve possuir 14 dígitos
         if (nro.Length != 14)
            return false;

         // exclui os dígitos verificadores para cálculo
         tempCnpj = nro[..12];

         // calcula o 1° dígito verificador
         soma = 0;
         for (int i = 0; i < 12; i++)
            soma += int.Parse(tempCnpj[i].ToString()) * multiplicador1[i];

         resto = soma % 11;

         if (resto < 2)
            resto = 0;
         else
            resto = 11 - resto;

         // digito 1 encontrado
         digito = resto.ToString();

         // calculo o 2° digito verificado; inclui o 1° digito pois ele faz parte do calculo
         tempCnpj += digito;
         soma = 0;
         for (int i = 0; i < 13; i++)
            soma += int.Parse(tempCnpj[i].ToString()) * multiplicador2[i];

         resto = soma % 11;

         if (resto < 2)
            resto = 0;
         else
            resto = 11 - resto;

         // digito 2 encontrado
         digito += resto.ToString();

         // retorna sucesso caso os valores do digito do número
         // e o calculado forem iguais
         return nro.EndsWith(digito);
      }

      /// <summary>
      /// Verifica a validade de um número de CPF informado.
      /// </summary>
      /// <param name="nro">Número a ser validado.</param>
      /// <returns>True se válido; caso contrário False.</returns>
      public static bool ValidarCpf(string nro)
      {
         int[] multiplicador1 = [10, 9, 8, 7, 6, 5, 4, 3, 2];
         int[] multiplicador2 = [11, 10, 9, 8, 7, 6, 5, 4, 3, 2];
         string tempCpf;
         string digito;
         int soma;
         int resto;

         // remove espaços e caracteres especiais; mantém somente os números
         nro = nro.Trim().Replace(".", "").Replace("-", "");

         // o número deve possuir 11 digitios
         if (nro.Length != 11)
            return false;

         // exclui os digitos verificadores
         tempCpf = nro[..9];

         // calcul o 1° digito verificador
         soma = 0;
         for (int i = 0; i < 9; i++)
            soma += int.Parse(tempCpf[i].ToString()) * multiplicador1[i];

         resto = soma % 11;

         if (resto < 2)
            resto = 0;
         else
            resto = 11 - resto;

         // digito 1 encontrdo
         digito = resto.ToString();

         // calculo o 2° digito verificado; inclui o 1° digito pois ele faz parte do calculo
         tempCpf += digito;
         soma = 0;
         for (int i = 0; i < 10; i++)
            soma += int.Parse(tempCpf[i].ToString()) * multiplicador2[i];

         resto = soma % 11;

         if (resto < 2)
            resto = 0;
         else
            resto = 11 - resto;

         // digito 2 encontrado
         digito += resto.ToString();

         // retorna sucesso caso os valores do digito do número
         // e o calculado forem iguais
         return nro.EndsWith(digito);
      }

      #endregion

      #region Check sum

      /// <summary>
      /// Retorna o dígito verificador de base 10.
      /// </summary>
      /// <param name="value">Valor a ser verificado.</param>
      /// <returns>Número do DV.</returns>
      public static int DVBase10(string value)
      {
         int dig;
         int dv = 0;
         int flag = 2;

         for (int i = value.Length - 1; i >= 0; i--)
         {
            dig = int.Parse(value.Substring(i, 1)) * flag;

            if (dig > 9)
               dig = 1 + (dig - 10);

            dv += dig;
            flag = flag == 2 ? 1 : 2;
         }

         dig = 10 - dv % 10;
         if (dig > 9) dig = 0;
         return dig;
      }


      #endregion

      public static string SerializeToXml<T>(this T value) //where T : class, T
      {
         if (value == null)
            return string.Empty;

         try
         {
            var xmlserializer = new XmlSerializer(typeof(T));
            var stringWriter = new StringWriter();
            using var writer = XmlWriter.Create(stringWriter);
            xmlserializer.Serialize(writer, value);
            return stringWriter.ToString();
         }
         catch (Exception ex)
         {
            throw new Exception("An error occurred", ex);
         }
      }

      public static T DeserializeToXml<T>(this string input)
      {
         try
         {
            var xmlSerializer = new XmlSerializer(typeof(T));
            var stringReader = new StringReader(input);
            using var reader = XmlReader.Create(stringReader);
            return (T)xmlSerializer.Deserialize(reader)!;
         }
         catch (Exception ex)
         {
            throw new Exception("An error ccurred", ex);
         }
      }
   }
}
