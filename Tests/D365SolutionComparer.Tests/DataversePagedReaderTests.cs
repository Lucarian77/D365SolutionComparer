using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class DataversePagedReaderTests
    {
        private static QueryExpression Query() => new QueryExpression("solutioncomponent")
        {
            ColumnSet = new ColumnSet("componenttype", "objectid")
        };

        private static EntityCollection Page(bool more, string cookie, params Entity[] records)
        {
            var page = new EntityCollection { MoreRecords = more, PagingCookie = cookie };
            page.Entities.AddRange(records);
            return page;
        }

        [TestMethod]
        public void ReadsAllPagesWithCookiesAndDeterministicOrderWithoutMutatingCaller()
        {
            var solutionId = Guid.NewGuid(); var query = Query();
            query.Criteria.AddCondition("solutionid", ConditionOperator.Equal, solutionId);
            query.AddOrder("componenttype", OrderType.Descending);
            query.PageInfo = new PagingInfo { PageNumber = 9, PagingCookie = "caller", Count = 17 };
            var first = new Entity("solutioncomponent", Guid.NewGuid());
            var second = new Entity("solutioncomponent", Guid.NewGuid());
            var third = new Entity("solutioncomponent", Guid.NewGuid());
            var service = new FakeOrganizationService();
            service.RetrievePage = request =>
            {
                Assert.AreNotSame(query, request);
                Assert.AreNotSame(query.Criteria, request.Criteria);
                Assert.AreEqual("solutioncomponent", request.EntityName);
                Assert.AreEqual(solutionId, request.Criteria.Conditions[0].Values[0]);
                CollectionAssert.AreEqual(new[] { "componenttype", "objectid" }, request.ColumnSet.Columns.ToArray());
                Assert.AreEqual(2, request.Orders.Count);
                Assert.AreEqual(OrderType.Descending, request.Orders[0].OrderType);
                Assert.AreEqual("solutioncomponentid", request.Orders[1].AttributeName);
                Assert.AreEqual(OrderType.Ascending, request.Orders[1].OrderType);
                Assert.AreEqual(2, request.PageInfo.Count);
                Assert.AreEqual(service.Calls, request.PageInfo.PageNumber);
                if (service.Calls == 1)
                {
                    Assert.IsNull(request.PageInfo.PagingCookie);
                    return Page(true, "<cookie page='1' />", first, second);
                }
                Assert.AreEqual("<cookie page='1' />", request.PageInfo.PagingCookie);
                return Page(false, null, third);
            };
            var progress = new List<RetrievalProgress>();
            var result = new DataversePagedReader(service).ReadAll(query, "solutioncomponentid", CancellationToken.None, progress.Add, 2);
            CollectionAssert.AreEqual(new[] { first, second, third }, new List<Entity>(result));
            Assert.AreEqual(2, progress.Count);
            Assert.AreEqual(1, progress[0].PagesRetrieved);
            Assert.AreEqual(2, progress[0].RecordsRetrieved);
            Assert.AreEqual(2, progress[1].PagesRetrieved);
            Assert.AreEqual(3, progress[1].RecordsRetrieved);
            Assert.AreEqual(9, query.PageInfo.PageNumber);
            Assert.AreEqual(17, query.PageInfo.Count);
            Assert.AreEqual("caller", query.PageInfo.PagingCookie);
            Assert.AreEqual(1, query.Orders.Count);
        }

        [DataTestMethod]
        [DataRow(null, OrderType.Ascending, 1)]
        [DataRow("componenttype", OrderType.Descending, 2)]
        [DataRow("solutioncomponentid", OrderType.Ascending, 1)]
        [DataRow("solutioncomponentid", OrderType.Descending, 1)]
        [DataRow("SoLuTiOnCoMpOnEnTiD", OrderType.Ascending, 1)]
        [DataRow("SoLuTiOnCoMpOnEnTiD", OrderType.Descending, 1)]
        public void PrimaryKeyOrderIsAddedOnlyWhenAbsentAndCallerQueryIsUnchanged(
            string existingAttribute, OrderType existingDirection, int expectedOrderCount)
        {
            var query = Query();
            if (existingAttribute != null) query.AddOrder(existingAttribute, existingDirection);
            var originalOrders = query.Orders;
            var originalOrder = query.Orders.SingleOrDefault();
            var originalPageInfo = new PagingInfo { Count = 17, PageNumber = 9, PagingCookie = "caller" };
            query.PageInfo = originalPageInfo;
            var service = new FakeOrganizationService
            {
                RetrievePage = request =>
                {
                    Assert.AreNotSame(query, request);
                    Assert.AreNotSame(originalOrders, request.Orders);
                    Assert.AreEqual(expectedOrderCount, request.Orders.Count);
                    if (existingAttribute != null)
                    {
                        Assert.AreNotSame(originalOrder, request.Orders[0]);
                        Assert.AreEqual(existingAttribute, request.Orders[0].AttributeName);
                        Assert.AreEqual(existingDirection, request.Orders[0].OrderType);
                    }
                    if (existingAttribute == null || expectedOrderCount == 2)
                    {
                        Assert.AreEqual("solutioncomponentid", request.Orders.Last().AttributeName);
                        Assert.AreEqual(OrderType.Ascending, request.Orders.Last().OrderType);
                    }
                    Assert.AreEqual(1, request.Orders.Count(order => string.Equals(
                        order.AttributeName, "solutioncomponentid", StringComparison.OrdinalIgnoreCase)));
                    return Page(false, null);
                }
            };

            new DataversePagedReader(service).ReadAll(query, "solutioncomponentid", CancellationToken.None);

            Assert.AreEqual(1, service.Calls);
            Assert.AreSame(originalOrders, query.Orders);
            Assert.AreEqual(existingAttribute == null ? 0 : 1, query.Orders.Count);
            if (originalOrder != null)
            {
                Assert.AreSame(originalOrder, query.Orders[0]);
                Assert.AreEqual(existingAttribute, query.Orders[0].AttributeName);
                Assert.AreEqual(existingDirection, query.Orders[0].OrderType);
            }
            Assert.AreSame(originalPageInfo, query.PageInfo);
            Assert.AreEqual(17, query.PageInfo.Count);
            Assert.AreEqual(9, query.PageInfo.PageNumber);
            Assert.AreEqual("caller", query.PageInfo.PagingCookie);
        }

        [TestMethod]
        public void EmptyFirstPageIsSuccessfulEmptyInventory()
        {
            var service = new FakeOrganizationService { RetrievePage = q => Page(false, null) };
            var reports = new List<RetrievalProgress>();
            var results = new DataversePagedReader(service).ReadAll(Query(), "solutioncomponentid", CancellationToken.None, reports.Add);
            Assert.AreEqual(0, results.Count);
            Assert.AreEqual(1, service.Calls);
            Assert.AreEqual(0, reports[0].RecordsRetrieved);
        }

        [TestMethod]
        public void CancellationBeforeReadMakesNoServiceCall()
        {
            var service = new FakeOrganizationService();
            Assert.ThrowsException<OperationCanceledException>(() => new DataversePagedReader(service).ReadAll(Query(), "solutioncomponentid", new CancellationToken(true)));
            Assert.AreEqual(0, service.Calls);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void CancellationDuringServiceCallNeverReturnsPartialOrFinalSuccess(bool moreRecords)
        {
            using (var cancellation = new CancellationTokenSource())
            {
                var service = new FakeOrganizationService { RetrievePage = q =>
                {
                    cancellation.Cancel();
                    return Page(moreRecords, "cookie", new Entity("solutioncomponent", Guid.NewGuid()));
                }};
                bool reported = false;
                Assert.ThrowsException<OperationCanceledException>(() => new DataversePagedReader(service).ReadAll(
                    Query(), "solutioncomponentid", cancellation.Token, p => reported = true));
                Assert.IsFalse(reported);
                Assert.AreEqual(1, service.Calls);
            }
        }

        [TestMethod]
        public void CancellationFromProgressPreventsNextRequest()
        {
            using (var cancellation = new CancellationTokenSource())
            {
                var service = new FakeOrganizationService { RetrievePage = q => Page(true, "cookie", new Entity()) };
                Assert.ThrowsException<OperationCanceledException>(() => new DataversePagedReader(service).ReadAll(
                    Query(), "solutioncomponentid", cancellation.Token, p => cancellation.Cancel()));
                Assert.AreEqual(1, service.Calls);
            }
        }

        [TestMethod]
        public void LaterPageServiceFaultPropagatesOriginalException()
        {
            var fault = new FaultException<OrganizationServiceFault>(new OrganizationServiceFault { ErrorCode = -1, Message = "Read denied" });
            var service = new FakeOrganizationService();
            service.RetrievePage = q => service.Calls == 1 ? Page(true, "cookie", new Entity()) : throw fault;
            var actual = Assert.ThrowsException<FaultException<OrganizationServiceFault>>(() => new DataversePagedReader(service).ReadAll(
                Query(), "solutioncomponentid", CancellationToken.None));
            Assert.AreSame(fault, actual);
            Assert.AreEqual(2, service.Calls);
        }

        [TestMethod]
        public void ProgressCallbackFaultPropagates()
        {
            var fault = new InvalidOperationException("Callback failed");
            var service = new FakeOrganizationService { RetrievePage = q => Page(true, "cookie", new Entity()) };
            var actual = Assert.ThrowsException<InvalidOperationException>(() => new DataversePagedReader(service).ReadAll(
                Query(), "solutioncomponentid", CancellationToken.None, p => throw fault));
            Assert.AreSame(fault, actual);
            Assert.AreEqual(1, service.Calls);
        }

        [TestMethod]
        public void MissingCookieFailsInsteadOfSilentlyTruncatingOrFallingBack()
        {
            var service = new FakeOrganizationService { RetrievePage = q => Page(true, null, new Entity()) };
            Assert.ThrowsException<NotSupportedException>(() => new DataversePagedReader(service).ReadAll(Query(), "solutioncomponentid", CancellationToken.None));
            Assert.AreEqual(1, service.Calls);
        }

        [TestMethod]
        public void RepeatedCookiesFailInsteadOfLoopingForever()
        {
            var service = new FakeOrganizationService { RetrievePage = q => Page(true, "same", new Entity()) };
            Assert.ThrowsException<InvalidOperationException>(() => new DataversePagedReader(service).ReadAll(Query(), "solutioncomponentid", CancellationToken.None));
            Assert.AreEqual(2, service.Calls);
        }

        [TestMethod]
        public void NullOrNonAdvancingResponseFails()
        {
            var service = new FakeOrganizationService { RetrievePage = q => null };
            var reader = new DataversePagedReader(service);
            Assert.ThrowsException<InvalidOperationException>(() => reader.ReadAll(Query(), "solutioncomponentid", CancellationToken.None));
            service.RetrievePage = q => Page(true, "cookie");
            Assert.ThrowsException<InvalidOperationException>(() => reader.ReadAll(Query(), "solutioncomponentid", CancellationToken.None));
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(5001)]
        public void InvalidPageSizesFailBeforeServiceCall(int pageSize)
        {
            var service = new FakeOrganizationService();
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new DataversePagedReader(service).ReadAll(
                Query(), "solutioncomponentid", CancellationToken.None, pageSize: pageSize));
            Assert.AreEqual(0, service.Calls);
        }

        [TestMethod]
        public void TopCountDistinctAndInvalidInputsFailBeforeServiceCall()
        {
            var service = new FakeOrganizationService(); var reader = new DataversePagedReader(service);
            Assert.ThrowsException<ArgumentNullException>(() => new DataversePagedReader(null));
            Assert.ThrowsException<ArgumentNullException>(() => reader.ReadAll(null, "solutioncomponentid", CancellationToken.None));
            Assert.ThrowsException<ArgumentException>(() => reader.ReadAll(Query(), "alias.id", CancellationToken.None));
            var query = Query(); query.TopCount = 10;
            Assert.ThrowsException<ArgumentException>(() => reader.ReadAll(query, "solutioncomponentid", CancellationToken.None));
            query.TopCount = null; query.Distinct = true;
            Assert.ThrowsException<ArgumentException>(() => reader.ReadAll(query, "solutioncomponentid", CancellationToken.None));
            query.Distinct = false; query.AddLink("solution", "solutionid", "solutionid");
            Assert.ThrowsException<ArgumentException>(() => reader.ReadAll(query, "solutioncomponentid", CancellationToken.None));
            Assert.AreEqual(0, service.Calls);
        }
    }
}
