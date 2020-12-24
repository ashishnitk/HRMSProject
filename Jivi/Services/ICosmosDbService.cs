using HRMS.Model;
using HRReporting.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace HRReporting.Services
{

    public interface ICosmosDbService
    {
        Task<List<Employee>> GetItemsAsync(string query);
        Task<Employee> GetItemAsync(string id);
        Task AddItemAsync(Employee item);
        Task UpdateItemAsync(string id, Employee item);
        Task DeleteItemAsync(string id);
        Task<BulkInviteResponseModel> createBulkItemAsync(List<Employee> items);
    }
}
