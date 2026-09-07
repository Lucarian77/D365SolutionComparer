using System;
using System.Collections.Generic;

namespace D365SolutionComparer.Infrastructure
{
    /// <summary>Operation-scoped, thread-safe request instrumentation. It does not own or retain a service client.</summary>
    public sealed class DataverseRequestCounter
    {
        private readonly object sync = new object();
        private readonly Dictionary<string, int> executeCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> queryCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private int executeRequests;
        private int queryRequests;

        public int ExecuteRequests { get { lock (sync) return executeRequests; } }
        public int QueryRequests { get { lock (sync) return queryRequests; } }
        public int TotalRequests { get { lock (sync) return executeRequests + queryRequests; } }

        public int GetExecuteCount(string requestName)
        {
            if (string.IsNullOrWhiteSpace(requestName)) throw new ArgumentException("A request name is required.", nameof(requestName));
            lock (sync) return GetCount(executeCounts, requestName);
        }

        public int GetQueryCount(string entityLogicalName)
        {
            if (string.IsNullOrWhiteSpace(entityLogicalName)) throw new ArgumentException("An entity logical name is required.", nameof(entityLogicalName));
            lock (sync) return GetCount(queryCounts, entityLogicalName);
        }

        internal void RecordExecute(string requestName)
        {
            lock (sync)
            {
                executeRequests++;
                Increment(executeCounts, string.IsNullOrWhiteSpace(requestName) ? "(unknown)" : requestName);
            }
        }

        internal void RecordQuery(string entityLogicalName)
        {
            lock (sync)
            {
                queryRequests++;
                Increment(queryCounts, string.IsNullOrWhiteSpace(entityLogicalName) ? "(unknown)" : entityLogicalName);
            }
        }

        private static int GetCount(IDictionary<string, int> values, string key)
        {
            int count;
            return values.TryGetValue(key, out count) ? count : 0;
        }

        private static void Increment(IDictionary<string, int> values, string key)
        {
            values[key] = GetCount(values, key) + 1;
        }
    }
}
