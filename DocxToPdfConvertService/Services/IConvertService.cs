namespace DocxToPdfConvertService.Services
{
    public interface IConvertService
    {
        Task<string> ConvertToPdf(string filepath);
    }
}
