using Microsoft.Office.Interop.Word;
using System.Reflection.Metadata;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

using System.Xml.Linq;

namespace DocxToPdfConvertService.Services
{
    public class ConvertService : IConvertService
    {
        public async Task<string> ConvertToPdf(string filepath)
        {
          
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                object oMissing = System.Reflection.Missing.Value;
                FileInfo wordfile = new FileInfo(filepath);
                word.Visible = false;
                word.ScreenUpdating = false;
                Object filename = (Object)wordfile.FullName;
                Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref oMissing, ref oMissing,
                  ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing);

                doc.Activate();

                object outputfilename = wordfile.FullName.Replace(".docx", ".pdf");
                object fileformat = WdSaveFormat.wdFormatPDF;

                doc.SaveAs2(ref outputfilename, ref fileformat, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                    , ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                object savechanges = WdSaveOptions.wdSaveChanges;
                ((_Document)doc).Close(ref savechanges, ref oMissing, oMissing);
                doc = null;

                ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                return outputfilename.ToString();
     

        }
    }
}
