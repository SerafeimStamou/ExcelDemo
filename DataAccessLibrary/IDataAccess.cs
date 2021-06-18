using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace DataAccessLibrary
{
    public interface IDataAccess
    {
        Task CreateExcelFile(FileInfo file);
        Task CreateFile();
        Task<List<Person>> LoadExcelFile(FileInfo file);
        Task<List<Person>> LoadFile();
        List<Person> SetMockData();
    }
}