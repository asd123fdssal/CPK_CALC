using CpkTool;
using MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes;
using PdfSharpCore.Fonts;
using PdfSharpCore.Utils;
using SixLabors.ImageSharp.PixelFormats;

namespace CPK_CALC
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            GlobalFontSettings.FontResolver ??= new KoreanFontResolver();
            ImageSource.ImageSourceImpl = new ImageSharpImageSource<Rgba32>();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                Application.Run(new CPKAnalysisForm());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"애플리케이션 실행 중 오류가 발생했습니다: {ex.Message}\n\n{ex.StackTrace}",
                              "심각한 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}