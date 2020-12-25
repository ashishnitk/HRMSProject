using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using HRReporting.Model;
using HRMS.Model;
using Microsoft.Azure.Cosmos;
using Microsoft.Azure.Cosmos.Fluent;
using Microsoft.Extensions.Configuration;

namespace HRReporting.Services
{
    public class CosmosDbService : ICosmosDbService
    {
        private Container _container;
        private readonly int _batchPerTransectionSize = 99;

        public CosmosDbService(CosmosClient dbClient, string databaseName, string containerName)
        {
            this._container = dbClient.GetContainer(databaseName, containerName);
        }

        public async Task AddItemAsync(Employee item)
        {
            await this._container.CreateItemAsync<Employee>(item, new PartitionKey(item.Month));
        }


        public async Task DeleteItemAsync(string id)
        {
            await this._container.DeleteItemAsync<Employee>(id, new PartitionKey(id));
        }

        public async Task<Employee> GetItemAsync(string id)
        {
            try
            {
                ItemResponse<Employee> response = await this._container.ReadItemAsync<Employee>(id, new PartitionKey(id));
                return response.Resource;
            }
            catch (CosmosException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                return null;
            }

        }

        public async Task<List<Employee>> GetItemsAsync(QueryDefinition inputQuery)
        {
            try
            {
                var query = this._container.GetItemQueryIterator<Employee>(inputQuery);
                List<Employee> results = new List<Employee>();
                while (query.HasMoreResults)
                {
                    var response = await query.ReadNextAsync();
                    results.AddRange(response.ToList());
                }
                return results;
            }
            catch (Exception e)
            {
                throw;
            }
        }

        public async Task UpdateItemAsync(string id, Employee item)
        {
            await this._container.UpsertItemAsync<Employee>(item, new PartitionKey(id));
        }

        public async Task<BulkInviteResponseModel> createBulkItemAsync(List<Employee> items)
        {
            List<TransactionalBatch> transactionalBatches = new List<TransactionalBatch>();
            int numberOfBatches = items.Count / 90 + 1;
            // Form the transactional batches and add it to batch list
            for (int i = 0; i < numberOfBatches; i++)
            {
                transactionalBatches.Add(this._container.CreateTransactionalBatch(new PartitionKey(items[0].Month)));
            }


            // Parse through each visit and create Items to batch based on batch transection size.  
            foreach (var item in items.Select((value, i) => new { i, value }))
            {
                var value = item.value;
                int index = item.i;
                int batchNo = index / _batchPerTransectionSize;
                transactionalBatches[batchNo].CreateItem<Employee>(value); // create each visitDoc in respective batch
            }

            List<TransactionalBatchResponse> result = new List<TransactionalBatchResponse>();
            foreach (TransactionalBatch i in transactionalBatches)
            {
                // execute batches one by one and accumulate the result of each batch transaction
                TransactionalBatchResponse res = await i.ExecuteAsync();
                result.Add(res);
            }

            List<BatchResponse> batchResponse = new List<BatchResponse>();
            List<string> failedVisits = new List<string>();
            // foreach (var item in result) // result would contain for the number of batches created and executed
            foreach (var item in result.Select((value, i) => new { i, value }))
            {
                TransactionalBatchResponse value = item.value;
                using (value)
                {
                    batchResponse.Add(new BatchResponse()
                    {
                        StatusCode = value.StatusCode,
                        IsSuccessStatusCode = value.IsSuccessStatusCode,
                        RetryAfter = value.RetryAfter,
                        ErrorMessage = value.ErrorMessage,
                        ActivityId = value.ActivityId
                    });
                    if (!value.IsSuccessStatusCode)
                    {
                        failedVisits.AddRange(items.Skip(_batchPerTransectionSize * item.i).Take(value.Count).Select(a => a.Id));
                        //TODO - Logging on failure of batch
                    }
                }
            }

            return new BulkInviteResponseModel() { response = batchResponse, failedVisits = failedVisits };
        }

    }
}
