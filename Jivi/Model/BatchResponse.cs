using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace HRReporting.Model
{

    /// <summary>
    /// response model for Bulk invite
    /// </summary>
    public class BulkInviteResponseModel
    {
        /// <summary>
        /// Responses
        /// </summary>
        public List<BatchResponse> response { get; set; }
        /// <summary>
        /// Failed Visits DocIds
        /// </summary>
        //public List<Visit> failedVisits { get; set; }
        public List<string> failedVisits { get; set; }
    }

    /// <summary>
    /// Batch Response
    /// </summary>
    public class BatchResponse
    {
        /// <summary>
        /// Gets the ActivityId that identifies the server request made to execute the batch.
        /// </summary>
        public string ActivityId { get; set; }
        /// <summary>
        /// Responses
        /// </summary>
        public HttpStatusCode StatusCode { get; set; }
        /// <summary>
        /// Created DocIds
        /// </summary>
        public bool IsSuccessStatusCode { get; set; }
        /// <summary>
        ///   Gets the reason for failure of the batch request.
        /// </summary>
        public string ErrorMessage { get; set; }
        /// <summary>
        /// Gets the number of operation results.
        /// </summary>
        public int Count { get; set; }
        /// <summary>
        ///  Gets the amount of time to wait before retrying this or any other request within
        ///  Cosmos container or collection due to throttling.
        /// </summary>
        public TimeSpan? RetryAfter { get; set; }
    }
}
