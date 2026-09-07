using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ProcessWorkflowIdentityResolverTests
    {
        [TestMethod]
        public void DirectUniqueNameResolutionRemainsUnchangedAndUsesOneGroupedRequest()
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); int requests = 0;
            var service = Service(solution, query =>
            {
                requests++;
                AssertWorkflowColumns(query, "workflowid", "uniquename", "type", "parentworkflowid");
                return Rows(Workflow(objectId, "new_Direct", 1));
            });

            var result = Resolve(service, solution, objectId);

            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Status);
            Assert.AreEqual("new_Direct", result.ComparisonKey);
            Assert.AreEqual(string.Empty, result.Diagnostic);
            Assert.AreEqual(1, requests);
        }

        [TestMethod]
        public void ActivationInheritsUniqueNameOnlyFromConfirmedParentDefinition()
        {
            var solution = Solution(); var activationId = Guid.NewGuid(); var parentId = Guid.NewGuid();
            var service = WorkflowService(solution, new[] { Workflow(activationId, null, 2, parentId) },
                new[] { Workflow(parentId, "new_Parent", 1) });

            var result = Resolve(service, solution, activationId);

            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Status);
            Assert.AreEqual("new_Parent", result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "inherited from the parent workflow definition");
            Assert.IsFalse(result.ComparisonKey.Contains(activationId.ToString("D")));
            Assert.IsFalse(result.ComparisonKey.Contains(parentId.ToString("D")));
            Assert.AreEqual(2, service.Calls);
        }

        [TestMethod]
        public void ActivationWithoutParentRemainsUnresolvedWithoutParentQuery()
        {
            var solution = Solution(); var activationId = Guid.NewGuid();
            var service = WorkflowService(solution, new[] { Workflow(activationId, null, 2) }, new Entity[0]);

            var result = Resolve(service, solution, activationId);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "activation has no parent workflow definition");
            Assert.AreEqual(1, service.Calls);
        }

        [TestMethod]
        public void MissingParentDefinitionRemainsUnresolved()
        {
            var solution = Solution(); var activationId = Guid.NewGuid(); var parentId = Guid.NewGuid();
            var service = WorkflowService(solution, new[] { Workflow(activationId, null, 2, parentId) },
                new Entity[0]);

            var result = Resolve(service, solution, activationId);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            StringAssert.Contains(result.Diagnostic, "Parent workflow definition was not found");
            Assert.AreEqual(2, service.Calls);
        }

        [TestMethod]
        public void ParentDefinitionWithBlankUniqueNameRemainsUnresolved()
        {
            var solution = Solution(); var activationId = Guid.NewGuid(); var parentId = Guid.NewGuid();
            var service = WorkflowService(solution, new[] { Workflow(activationId, null, 2, parentId) },
                new[] { Workflow(parentId, " ", 1) });

            var result = Resolve(service, solution, activationId);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            StringAssert.Contains(result.Diagnostic, "Parent workflow definition has a blank uniquename");
        }

        [TestMethod]
        public void DefinitionWithBlankUniqueNameRemainsUnresolved()
        {
            var solution = Solution(); var definitionId = Guid.NewGuid();
            var service = WorkflowService(solution, new[] { Workflow(definitionId, null, 1) }, new Entity[0]);

            var result = Resolve(service, solution, definitionId);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            StringAssert.Contains(result.Diagnostic, "Workflow definition has a blank uniquename");
            Assert.AreEqual(1, service.Calls);
        }

        [DataTestMethod]
        [DataRow(3)]
        [DataRow(99)]
        public void UnsupportedWorkflowRecordTypeDoesNotUseParentFallback(int workflowType)
        {
            var solution = Solution(); var objectId = Guid.NewGuid(); var parentId = Guid.NewGuid();
            var service = WorkflowService(solution, new[] { Workflow(objectId, null, workflowType, parentId) },
                new[] { Workflow(parentId, "must_not_be_used", 1) });

            var result = Resolve(service, solution, objectId);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "Unsupported workflow record type");
            Assert.AreEqual(1, service.Calls);
        }

        [TestMethod]
        public void RepeatedActivationsShareOneParentLookupButRemainSeparateCandidates()
        {
            var solution = Solution(); var first = Guid.NewGuid(); var second = Guid.NewGuid();
            var parentId = Guid.NewGuid();
            var service = WorkflowService(solution,
                new[] { Workflow(first, null, 2, parentId), Workflow(second, null, 2, parentId) },
                new[] { Workflow(parentId, "new_Shared", 1) });

            var snapshot = ResolveSnapshot(service, solution, first, second);

            Assert.AreEqual(2, snapshot.Components.Count);
            Assert.IsTrue(snapshot.Components.All(item => item.Status == IdentityResolutionStatus.Resolved));
            Assert.IsTrue(snapshot.Components.All(item => item.ComparisonKey == "new_Shared"));
            Assert.AreNotSame(snapshot.Components[0].Record, snapshot.Components[1].Record);
            Assert.AreEqual(2, service.Calls);
            var compared = new SolutionMembershipComparer().Compare(snapshot,
                MembershipSnapshot.Absent(solution.Environment, solution.UniqueName, DateTimeOffset.UtcNow));
            Assert.IsTrue(compared.All(item => item.Source.Status == IdentityResolutionStatus.Ambiguous));
            Assert.IsTrue(compared.All(item => item.Presence == MembershipPresence.Indeterminate));
        }

        [TestMethod]
        public void MixedDirectAndInheritedWorkflowsUseTwoGroupedRequests()
        {
            var solution = Solution(); var directId = Guid.NewGuid(); var activationId = Guid.NewGuid();
            var parentId = Guid.NewGuid();
            var service = WorkflowService(solution,
                new[] { Workflow(directId, "new_Direct", 1), Workflow(activationId, null, 2, parentId) },
                new[] { Workflow(parentId, "new_Inherited", 1) });

            var snapshot = ResolveSnapshot(service, solution, directId, activationId);

            CollectionAssert.AreEqual(new[] { "new_Direct", "new_Inherited" },
                snapshot.Components.Select(item => item.ComparisonKey).ToArray());
            Assert.AreEqual(string.Empty, snapshot.Components[0].Diagnostic);
            StringAssert.Contains(snapshot.Components[1].Diagnostic, "inherited");
            Assert.AreEqual(2, service.Calls);
        }

        [TestMethod]
        public void DirectAndInheritedSameIdentityRemainDuplicateCandidates()
        {
            var solution = Solution(); var directId = Guid.NewGuid(); var activationId = Guid.NewGuid();
            var parentId = Guid.NewGuid();
            var service = WorkflowService(solution,
                new[] { Workflow(directId, "new_Same", 1), Workflow(activationId, null, 2, parentId) },
                new[] { Workflow(parentId, "new_Same", 1) });
            var snapshot = ResolveSnapshot(service, solution, directId, activationId);

            Assert.AreEqual(2, snapshot.Components.Count(item => item.Status == IdentityResolutionStatus.Resolved));
            var compared = new SolutionMembershipComparer().Compare(snapshot,
                MembershipSnapshot.Absent(solution.Environment, solution.UniqueName, DateTimeOffset.UtcNow));
            Assert.IsTrue(compared.All(item => item.Source.Status == IdentityResolutionStatus.Ambiguous));
        }

        [TestMethod]
        public void ParentLookupFaultLeavesActivationUnresolved()
        {
            var solution = Solution(); var activationId = Guid.NewGuid(); var parentId = Guid.NewGuid();
            int queryNumber = 0;
            var service = Service(solution, query =>
            {
                queryNumber++;
                if (queryNumber == 1) return Rows(Workflow(activationId, null, 2, parentId));
                throw new FaultException("Parent denied");
            });

            var result = Resolve(service, solution, activationId);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            StringAssert.Contains(result.Diagnostic, "Parent workflow identity read failed: Parent denied");
        }

        [TestMethod]
        public void CancellationDuringParentLookupPropagates()
        {
            var solution = Solution(); var activationId = Guid.NewGuid(); var parentId = Guid.NewGuid();
            using (var cancellation = new CancellationTokenSource())
            {
                int queryNumber = 0;
                var service = Service(solution, query =>
                {
                    queryNumber++;
                    if (queryNumber == 1) return Rows(Workflow(activationId, null, 2, parentId));
                    cancellation.Cancel();
                    return Rows(Workflow(parentId, "new_Parent", 1));
                });
                Assert.ThrowsException<OperationCanceledException>(() =>
                    new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                        Record(activationId), cancellation.Token));
            }
        }

        [TestMethod]
        public void DistinctParentsAreGroupedAndRespectBatchSize()
        {
            var solution = Solution();
            var activations = Enumerable.Range(0, 201).Select(index => new
            {
                Child = Guid.NewGuid(), Parent = Guid.NewGuid(), Index = index
            }).ToArray();
            var rawIds = new HashSet<Guid>(activations.Select(item => item.Child));
            int rawRequests = 0; int parentRequests = 0;
            var service = Service(solution, query =>
            {
                var ids = QueryIds(query).ToArray();
                Assert.IsTrue(ids.Length <= 200);
                if (ids.All(rawIds.Contains))
                {
                    rawRequests++;
                    return Rows(activations.Where(item => ids.Contains(item.Child))
                        .Select(item => Workflow(item.Child, null, 2, item.Parent)).ToArray());
                }
                parentRequests++;
                return Rows(activations.Where(item => ids.Contains(item.Parent))
                    .Select(item => Workflow(item.Parent, "new_" + item.Index, 1)).ToArray());
            });

            var snapshot = ResolveSnapshot(service, solution, activations.Select(item => item.Child).ToArray());

            Assert.AreEqual(201, snapshot.Components.Count);
            Assert.IsTrue(snapshot.Components.All(item => item.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(2, rawRequests);
            Assert.AreEqual(2, parentRequests);
        }

        [TestMethod]
        public void AmbiguousParentLookupDoesNotProducePortableIdentity()
        {
            var solution = Solution(); var activationId = Guid.NewGuid(); var parentId = Guid.NewGuid();
            var service = WorkflowService(solution, new[] { Workflow(activationId, null, 2, parentId) },
                new[] { Workflow(parentId, "first", 1), Workflow(parentId, "second", 1) });

            var result = Resolve(service, solution, activationId);

            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Status);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "Parent workflow definition lookup returned multiple records");
        }

        [TestMethod]
        public void ParentRecordMustBeAConfirmedDefinition()
        {
            var solution = Solution(); var activationId = Guid.NewGuid(); var parentId = Guid.NewGuid();
            var service = WorkflowService(solution, new[] { Workflow(activationId, null, 2, parentId) },
                new[] { Workflow(parentId, "unsafe_activation_name", 2) });

            var result = Resolve(service, solution, activationId);

            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.IsNull(result.ComparisonKey);
            StringAssert.Contains(result.Diagnostic, "not a confirmed definition");
        }

        [TestMethod]
        public void RawWorkflowFaultIsUnresolvedAndUnexpectedExceptionsPropagate()
        {
            var solution = Solution(); var objectId = Guid.NewGuid();
            var faulted = Service(solution, query => throw new FaultException("Workflow denied"));
            var faultedResult = Resolve(faulted, solution, objectId);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, faultedResult.Status);
            StringAssert.Contains(faultedResult.Diagnostic, "Workflow identity read failed: Workflow denied");

            var unexpected = Service(solution, query => throw new InvalidOperationException("Malformed response"));
            Assert.ThrowsException<InvalidOperationException>(() => Resolve(unexpected, solution, objectId));
        }

        [TestMethod]
        public void RequestCounterRecordsDirectRepeatedParentDistinctParentAndMixedCosts()
        {
            AssertWorkflowRequestCount(new[]
            {
                Workflow(Guid.NewGuid(), "direct_one", 1), Workflow(Guid.NewGuid(), "direct_two", 1)
            }, new Entity[0], 1);

            var sharedParent = Guid.NewGuid();
            AssertWorkflowRequestCount(new[]
            {
                Workflow(Guid.NewGuid(), null, 2, sharedParent),
                Workflow(Guid.NewGuid(), null, 2, sharedParent)
            }, new[] { Workflow(sharedParent, "shared", 1) }, 2);

            var firstParent = Guid.NewGuid(); var secondParent = Guid.NewGuid();
            AssertWorkflowRequestCount(new[]
            {
                Workflow(Guid.NewGuid(), null, 2, firstParent),
                Workflow(Guid.NewGuid(), null, 2, secondParent)
            }, new[] { Workflow(firstParent, "first", 1), Workflow(secondParent, "second", 1) }, 2);

            var mixedParent = Guid.NewGuid();
            AssertWorkflowRequestCount(new[]
            {
                Workflow(Guid.NewGuid(), "direct", 1), Workflow(Guid.NewGuid(), null, 2, mixedParent)
            }, new[] { Workflow(mixedParent, "inherited", 1) }, 2);
        }

        private static ComponentIdentity Resolve(FakeOrganizationService service, SolutionIdentity solution,
            Guid objectId) => new DataverseComponentIdentityResolver().Resolve(service, solution.Environment,
                Record(objectId), CancellationToken.None);

        private static MembershipSnapshot ResolveSnapshot(FakeOrganizationService service,
            SolutionIdentity solution, params Guid[] objectIds) =>
            ResolveSnapshot(service, solution, null, objectIds);

        private static MembershipSnapshot ResolveSnapshot(FakeOrganizationService service,
            SolutionIdentity solution, DataverseRequestCounter requestCounter, params Guid[] objectIds)
        {
            var input = MembershipSnapshot.Complete(solution, objectIds.Select(id =>
                new ComponentIdentity(Record(id), IdentityResolutionStatus.Unresolved)), DateTimeOffset.UtcNow);
            return new DataverseComponentIdentityResolver().ResolveSnapshot(service, input,
                CancellationToken.None, requestCounter);
        }

        private static void AssertWorkflowRequestCount(IReadOnlyCollection<Entity> rawRows,
            IReadOnlyCollection<Entity> parentRows, int expectedWorkflowRequests)
        {
            var solution = Solution(); var counter = new DataverseRequestCounter();
            var service = WorkflowService(solution, rawRows, parentRows);
            var result = ResolveSnapshot(service, solution, counter, rawRows.Select(item => item.Id).ToArray());
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Resolved));
            Assert.AreEqual(expectedWorkflowRequests, counter.GetQueryCount("workflow"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
        }

        private static FakeOrganizationService WorkflowService(SolutionIdentity solution,
            IReadOnlyCollection<Entity> rawRows, IReadOnlyCollection<Entity> parentRows)
        {
            var rawIds = new HashSet<Guid>(rawRows.Select(item => item.Id));
            return Service(solution, query =>
            {
                var ids = QueryIds(query).ToArray();
                var isRawQuery = ids.All(rawIds.Contains);
                AssertWorkflowColumns(query, isRawQuery
                    ? new[] { "workflowid", "uniquename", "type", "parentworkflowid" }
                    : new[] { "workflowid", "uniquename", "type" });
                var source = isRawQuery ? rawRows : parentRows;
                return Rows(source.Where(item => ids.Contains(item.Id)).ToArray());
            });
        }

        private static SolutionComponentRecord Record(Guid objectId) =>
            new SolutionComponentRecord(Guid.NewGuid(), 29, objectId);

        private static Entity Workflow(Guid id, string uniqueName, int? type, Guid? parentId = null)
        {
            var row = new Entity("workflow", id);
            if (uniqueName != null) row["uniquename"] = uniqueName;
            if (type.HasValue) row["type"] = new OptionSetValue(type.Value);
            if (parentId.HasValue) row["parentworkflowid"] = new EntityReference("workflow", parentId.Value);
            return row;
        }

        private static IEnumerable<Guid> QueryIds(QueryExpression query)
        {
            Assert.AreEqual("workflow", query.EntityName);
            Assert.AreEqual("workflowid", query.Criteria.Conditions.Single().AttributeName);
            Assert.AreEqual(ConditionOperator.In, query.Criteria.Conditions.Single().Operator);
            return query.Criteria.Conditions.Single().Values.Cast<Guid>();
        }

        private static void AssertWorkflowColumns(QueryExpression query, params string[] expected) =>
            CollectionAssert.AreEquivalent(expected, query.ColumnSet.Columns.ToArray());
    }
}
