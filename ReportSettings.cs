using System;
using System.Collections.Generic;
using System.Drawing;

namespace EmissorNFe.Printing
{
   public sealed class ReportSettings : IPrintSettings
   {
      public ReportSettings()
      {
         UserData = new Dictionary<string, object>();
      }

      public ReportSettings(string name)
         : this()
      {
         ReportName = name;
      }

      public ReportSettings(string name, string title)
         : this()
      {
         ReportName = name;
         ReportTitle = title;
      }

      /// <summary>
      /// The action to read the options chosen by the user to print the report.
      /// </summary>
      public Action ReadOptions { get; set; }

      /// <summary>
      /// The application name.
      /// </summary>
      public string AppTitle { get; set; } = null;

      /// <summary>
      /// The string value that represents the current version of the application.
      /// </summary>
      public string AppVersion { get; set; } = null;

      /// <summary>
      /// The name of report.
      /// </summary>
      public string ReportName { get; set; } = null;

      /// <summary>
      /// The title of report.
      /// </summary>
      public string ReportTitle { get; set; } = null;

      /// <summary>
      /// Image used as a logo in reports.
      /// </summary>
      public Image ImageLogo { get; set; }

      /// <summary>
      /// An instance of the Windows Form object that is the source of the report.
      /// </summary>
      public object Window { get; set; } = null;

      /// <summary>
      /// All users data required for report printing.
      /// </summary>
      public Dictionary<string, object> UserData { get; }
   }
}
