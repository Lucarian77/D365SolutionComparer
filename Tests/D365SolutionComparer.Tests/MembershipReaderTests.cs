using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MembershipReaderTests
    {
        [TestMethod]
        public void EmptyPresentSolutionIsCompleteAndUsesSelectedSolutionQuery()
        {
            var solution = Solution();
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solution")
                {
                    Assert.AreEqual(2, query.TopCount);
                    Assert.AreEqual("uniquename", query.Criteria.Conditions.Single().AttributeName);
                    Assert.AreEqual(solution.UniqueName, query.Criteria.Conditions.Single().Values.Single());
                    return Rows(SolutionRow(solution));
                }
                Assert.AreEqual("solutioncomponent", query.EntityName);
                Assert.AreEqual(solution.SolutionId, query.Criteria.Conditions.Single().Values.Single());
                Assert.AreEqual("solutionid", query.Criteria.Conditions.Single().AttributeName);
                Assert.AreEqual("solutioncomponentid", query.Orders.Single().AttributeName);
                Assert.AreEqual(OrderType.Ascending, query.Orders.Single().OrderType);
                CollectionAssert.AreEquivalent(new[] { "solutioncomponentid", "componenttype", "objectid", "rootcomponentbehavior",
                    "rootsolutioncomponentid", "ismetadata", "solutionid" }, query.ColumnSet.Columns.ToArray());
                return Rows();
            });
            var result = new DataverseSolutionMembershipReader().Read(service, solution, CancellationToken.None);
            Assert.AreEqual(MembershipSnapshotState.Complete, result.State);
            Assert.AreEqual(solution.SolutionId, result.Solution.SolutionId);
            Assert.AreEqual(0, result.Components.Count);
            Assert.AreEqual(1, service.ExecuteCalls);
        }

        [TestMethod]
        public void MultiplePagesPreserveEveryRawFieldIncludingUnknownCodesAndNulls()
        {
            var solution = Solution();
            var first = ComponentRow(solution, 987654); var second = ComponentRow(solution, 2); var third = ComponentRow(solution, 10);
            second.Attributes.Remove("objectid"); second.Attributes.Remove("ismetadata");
            second.Attributes.Remove("rootcomponentbehavior"); second.Attributes.Remove("rootsolutioncomponentid");
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solution") return Rows(SolutionRow(solution));
                int page = query.PageInfo.PageNumber;
                Assert.AreEqual(page == 1 ? null : "cookie" + (page - 1), query.PageInfo.PagingCookie);
                var response = Rows(page == 1 ? first : page == 2 ? second : third);
                response.MoreRecords = page < 3;
                response.PagingCookie = response.MoreRecords ? "cookie" + page : null;
                return response;
            });
            var progress = new List<RetrievalProgress>();
            var result = new DataverseSolutionMembershipReader().Read(service, solution, CancellationToken.None, progress.Add);
            Assert.AreEqual(3, result.Components.Count);
            Assert.AreEqual(3, progress.Last().RecordsRetrieved);
            Assert.AreEqual(3, progress.Last().PagesRetrieved);
            var raw = result.Components[0].Record;
            Assert.AreEqual(first.Id, raw.SolutionComponentId);
            Assert.AreEqual(987654, raw.ComponentType);
            Assert.AreEqual(first.GetAttributeValue<Guid>("objectid"), raw.ObjectId);
            Assert.AreEqual(first.GetAttributeValue<Guid>("rootsolutioncomponentid"), raw.RootSolutionComponentId);
            Assert.AreEqual(2, raw.RootComponentBehavior);
            Assert.AreEqual(false, raw.IsMetadata);
            Assert.IsNull(result.Components[1].Record.ObjectId);
            Assert.IsNull(result.Components[1].Record.RootComponentBehavior);
            Assert.IsNull(result.Components[1].Record.RootSolutionComponentId);
            Assert.IsNull(result.Components[1].Record.IsMetadata);
            Assert.IsTrue(result.Components.All(c => c.Status == IdentityResolutionStatus.Unresolved));
        }

        [TestMethod]
        public void AbsentSolutionReturnsAbsentWithoutMembershipQuery()
        {
            var solution = Solution(); var service = Service(solution, query =>
            {
                Assert.AreEqual("solution", query.EntityName);
                return Rows();
            });
            var result = new DataverseSolutionMembershipReader().Read(service, solution.Environment, solution.UniqueName, CancellationToken.None);
            Assert.AreEqual(MembershipSnapshotState.SolutionAbsent, result.State);
            Assert.IsNull(result.Solution);
            Assert.AreEqual(1, service.Calls);
        }

        [TestMethod]
        public void StaleSolutionIdAndAmbiguousSelectionFailBeforeMembershipRead()
        {
            var solution = Solution(); var service = Service(solution, q => Rows(SolutionRow(Solution())));
            var reader = new DataverseSolutionMembershipReader();
            Assert.ThrowsException<InvalidOperationException>(() => reader.Read(service, solution, CancellationToken.None));
            service.RetrievePage = q => Rows(SolutionRow(solution), SolutionRow(Solution()));
            Assert.ThrowsException<InvalidOperationException>(() => reader.Read(service, solution, CancellationToken.None));
        }

        [TestMethod]
        public void ServiceEnvironmentMismatchFailsBeforeSolutionQuery()
        {
            var solution = Solution(); var service = Service(Solution());
            Assert.ThrowsException<InvalidOperationException>(() => new DataverseSolutionMembershipReader().Read(service, solution, CancellationToken.None));
            Assert.AreEqual(0, service.Calls);
        }

        [DataTestMethod]
        [DataRow("id")]
        [DataRow("type")]
        [DataRow("solution")]
        public void MalformedRawRowsFailInsteadOfFabricatingIdentifiers(string invalidField)
        {
            var solution = Solution(); var component = ComponentRow(solution);
            if (invalidField == "id") component.Id = Guid.Empty;
            if (invalidField == "type") component.Attributes.Remove("componenttype");
            if (invalidField == "solution") component["solutionid"] = new EntityReference("solution", Guid.NewGuid());
            var service = Service(solution, q => q.EntityName == "solution" ? Rows(SolutionRow(solution)) : Rows(component));
            Assert.ThrowsException<InvalidOperationException>(() => new DataverseSolutionMembershipReader().Read(service, solution, CancellationToken.None));
        }

        [TestMethod]
        public void PageTwoFaultPropagatesAndStateAlternativeIsUnavailableNotPartial()
        {
            var solution = Solution();
            var fault = new FaultException<OrganizationServiceFault>(new OrganizationServiceFault { Message = "Denied" }, new FaultReason("Denied"));
            var service = Service(solution, query =>
            {
                if (query.EntityName == "solution") return Rows(SolutionRow(solution));
                if (query.PageInfo.PageNumber == 2) throw fault;
                var response = Rows(ComponentRow(solution)); response.MoreRecords = true; response.PagingCookie = "cookie";
                return response;
            });
            var reader = new DataverseSolutionMembershipReader();
            Assert.AreSame(fault, Assert.ThrowsException<FaultException<OrganizationServiceFault>>(() => reader.Read(service, solution, CancellationToken.None)));
            var result = reader.ReadOrUnavailable(service, solution.Environment, solution.UniqueName, CancellationToken.None);
            Assert.AreEqual(MembershipSnapshotState.Unavailable, result.State);
            Assert.AreEqual(0, result.Components.Count);
            StringAssert.Contains(result.Diagnostic, "Denied");
        }

        [TestMethod]
        public void DisconnectedServiceCanBeRepresentedAsUnavailable()
        {
            var solution = Solution();
            var result = new DataverseSolutionMembershipReader().ReadOrUnavailable(null, solution.Environment, solution.UniqueName, CancellationToken.None);
            Assert.AreEqual(MembershipSnapshotState.Unavailable, result.State);
        }

        [TestMethod]
        public void CancellationBeforeReadDoesNotExecuteAnyDataverseCall()
        {
            var solution = Solution(); var service = Service(solution);
            var reader = new DataverseSolutionMembershipReader();
            Assert.ThrowsException<OperationCanceledException>(() => reader.Read(service, solution, new CancellationToken(true)));
            Assert.ThrowsException<OperationCanceledException>(() => reader.ReadOrUnavailable(service, solution.Environment, solution.UniqueName, new CancellationToken(true)));
            Assert.AreEqual(0, service.Calls);
            Assert.AreEqual(0, service.ExecuteCalls);
        }

        [TestMethod]
        public void CancellationBetweenMembershipPagesNeverCompletesSnapshot()
        {
            var solution = Solution();
            using (var cancellation = new CancellationTokenSource())
            {
                var service = Service(solution, query =>
                {
                    if (query.EntityName == "solution") return Rows(SolutionRow(solution));
                    var result = Rows(ComponentRow(solution)); result.MoreRecords = true; result.PagingCookie = "cookie";
                    return result;
                });
                Assert.ThrowsException<OperationCanceledException>(() => new DataverseSolutionMembershipReader().Read(service, solution,
                    cancellation.Token, progress => cancellation.Cancel()));
                Assert.AreEqual(2, service.Calls);
            }
        }
    }
}
