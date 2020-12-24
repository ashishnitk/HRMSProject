using HRMS.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace HRReporting.Services
{

    public interface ICosmosDbService
    {
        Task<IEnumerable<Employee>> GetItemsAsync(string query);
        Task<Employee> GetItemAsync(string id);
        Task AddItemAsync(Employee item);
        Task UpdateItemAsync(string id, Employee item);
        Task DeleteItemAsync(string id);
    }
}
