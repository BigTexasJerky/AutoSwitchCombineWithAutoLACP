using HarmonyLib;
using Il2Cpp;
using Il2CppInterop.Runtime.Injection;
using MelonLoader;
using MelonLoader.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using UnityEngine;
using Component = UnityEngine.Component;

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "2.19.0", "Big Texas Jerky")]
[assembly: MelonGame("Waseku", "Data Center")]

namespace AutoSwitch
{
    public sealed class AutoSwitchMod : MelonMod
    {
        private const float ScanIntervalSeconds = 6.0f;

        private const float SameFabricXZTolerance = 0.40f;
        private const float AdjacentYTolerance = 0.040f;
        private const float FallbackExpectedStep = 0.0444f;

        private const float OriginXZRejectRadius = 1.25f;
        private const float MinUsefulWorldY = 1.0f;
        private const float MaxUsefulWorldY = 4.5f;

        private static readonly Regex TrailingRuntimeIdRegex =
            new Regex(@"_[\-]?\d+$", RegexOptions.Compiled);

        private static string DebugFolderPath =>
            Path.Combine(MelonEnvironment.ModsDirectory, "AutoSwitch");

        private static string DebugLogPath =>
            Path.Combine(DebugFolderPath, "autoswitch-debug.log");

        private static readonly HashSet<string> AllowedSwitchModels = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Switch32xQSFP",
            "Switch4xQSXP16xSFP",
            "Switch4xSFP",
            "Switch16CU"
        };

        internal static readonly Dictionary<string, RegisteredSwitchInfo> RegisteredSwitches =
            new Dictionary<string, RegisteredSwitchInfo>(StringComparer.OrdinalIgnoreCase);

        internal static readonly Dictionary<int, RegisteredCableInfo> RegisteredCables =
            new Dictionary<int, RegisteredCableInfo>();

        private static readonly Dictionary<string, float> FabricIdToSpeed =
            new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, List<int>> FabricIdToBundleCableIds =
            new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, string> DeviceIdToFabricIdSnapshot =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, string> FabricIdToDomainSnapshot =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, FabricRuntimePlan> FabricPlansSnapshot =
            new Dictionary<string, FabricRuntimePlan>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedPatches =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedBundleSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedRemoteResolutionSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedDomainLinkSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedFanoutCandidateSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedPropagationSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedFabricPlanSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private float _nextScanTime;
        private string _lastSummary = string.Empty;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();

            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(
                DebugLogPath,
                "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "] AutoSwitch 2.19.0 debug log started." + Environment.NewLine
            );

            InstallNativePatches();

            MelonLogger.Msg("[AutoSwitch] v2.19.0 active. Fabric share planner + routing-domain propagation.");
        }

        public override void OnSceneWasLoaded(int buildIndex, string sceneName)
        {
            if (buildIndex == 0)
                return;

            _nextScanTime = 0f;
            _lastSummary = string.Empty;

            RegisteredSwitches.Clear();
            RegisteredCables.Clear();
            FabricIdToSpeed.Clear();
            FabricIdToBundleCableIds.Clear();
            DeviceIdToFabricIdSnapshot.Clear();
            FabricIdToDomainSnapshot.Clear();
            FabricPlansSnapshot.Clear();

            LoggedBundleSignatures.Clear();
            LoggedRemoteResolutionSignatures.Clear();
            LoggedDomainLinkSignatures.Clear();
            LoggedFanoutCandidateSignatures.Clear();
            LoggedPropagationSignatures.Clear();
            LoggedFabricPlanSignatures.Clear();

            SaveDataAutoLACP.ResetForScene();
            SaveDataAutoLACP.StartSafeBootstrap();

            LogToFile("Scene loaded: " + sceneName + " (" + buildIndex.ToString(CultureInfo.InvariantCulture) + ")");
        }

        public override void OnUpdate()
        {
            try
            {
                if (Time.time < _nextScanTime)
                    return;

                _nextScanTime = Time.time + ScanIntervalSeconds;
                RunLiveBundlePass();
            }
            catch (Exception ex)
            {
                MelonLogger.Error("[AutoSwitch] OnUpdate exception: " + ex);
                LogToFile("OnUpdate exception: " + ex);
                _nextScanTime = Time.time + 10.0f;
            }
        }

        private void RunLiveBundlePass()
        {
            RefreshRegisteredSwitchPositions();

            List<RegisteredSwitchInfo> liveSwitches = RegisteredSwitches.Values
                .Where(s => s != null && s.HasWorldPosition && IsAllowedSwitchModel(s.ModelName))
                .Where(IsLikelyPlacedSwitch)
                .OrderBy(s => s.WorldPosition.x)
                .ThenBy(s => s.WorldPosition.y)
                .ThenBy(s => s.WorldPosition.z)
                .ToList();

            List<List<RegisteredSwitchInfo>> fabrics = BuildRegistryFabrics(liveSwitches)
                .Where(f => f.Count >= 2)
                .OrderByDescending(f => f.Count)
                .ThenByDescending(f => f.Sum(x => x.EstimatedCapacityGbps))
                .ToList();

            Dictionary<string, string> deviceIdToFabricId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, HashSet<string>> fabricIdToMembers = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < fabrics.Count; i++)
            {
                string fabricId = "FABRIC-" + (i + 1).ToString("000", CultureInfo.InvariantCulture);
                HashSet<string> members = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (RegisteredSwitchInfo sw in fabrics[i])
                {
                    if (sw != null && !string.IsNullOrWhiteSpace(sw.DeviceName))
                    {
                        deviceIdToFabricId[sw.DeviceName] = fabricId;
                        members.Add(sw.DeviceName);
                    }
                }

                fabricIdToMembers[fabricId] = members;
            }

            DeviceIdToFabricIdSnapshot.Clear();
            foreach (KeyValuePair<string, string> kvp in deviceIdToFabricId)
                DeviceIdToFabricIdSnapshot[kvp.Key] = kvp.Value;

            HashSet<string> allManagedIds = new HashSet<string>(
                deviceIdToFabricId.Keys,
                StringComparer.OrdinalIgnoreCase);

            NetworkSaveData networkSaveData = GetNetworkSaveData();
            int saveCableCount = CountSaveCables(networkSaveData);
            HashSet<int> externalLacpCableIds = GetExternalLacpCableIds(networkSaveData);

            FabricIdToSpeed.Clear();
            FabricIdToBundleCableIds.Clear();
            FabricPlansSnapshot.Clear();

            Dictionary<string, HashSet<string>> fabricAdjacency =
                BuildFabricAdjacency(deviceIdToFabricId, externalLacpCableIds);

            Dictionary<string, string> fabricToDomain =
                BuildRoutingDomains(fabricAdjacency, fabricIdToMembers.Keys.ToList());

            FabricIdToDomainSnapshot.Clear();
            foreach (KeyValuePair<string, string> kvp in fabricToDomain)
                FabricIdToDomainSnapshot[kvp.Key] = kvp.Value;

            List<FabricRuntimePlan> fabricPlans = BuildFabricRuntimePlans(
                fabrics,
                fabricIdToMembers,
                deviceIdToFabricId,
                fabricToDomain,
                externalLacpCableIds);

            foreach (FabricRuntimePlan plan in fabricPlans)
                FabricPlansSnapshot[plan.FabricId] = plan;

            ApplyFabricTags(fabricPlans);

            List<BundleBuilder> allBundles = new List<BundleBuilder>();
            HashSet<int> globallyClaimedCableIds = new HashSet<int>();

            int safeNativeBundleCount = 0;
            int fanoutCandidateCount = 0;
            int internalPassThroughCount = 0;
            int interFabricPassThroughCount = 0;
            int propagatedEdgeCount = 0;
            int propagatedDownstreamCount = 0;

            HashSet<int> internalPassThroughCableIds = new HashSet<int>();
            HashSet<int> interFabricPassThroughCableIds = new HashSet<int>();
            HashSet<int> fanoutCandidateCableIds = new HashSet<int>();

            foreach (FabricRuntimePlan plan in fabricPlans.OrderBy(p => p.FabricId, StringComparer.OrdinalIgnoreCase))
            {
                FabricIdToSpeed[plan.FabricId] = plan.EstimatedCapacityGbps;

                Dictionary<string, RemoteEdgeAccumulator> accumulators =
                    BuildRemoteEdgeAccumulators(
                        plan,
                        deviceIdToFabricId,
                        fabricToDomain,
                        externalLacpCableIds,
                        ref internalPassThroughCount,
                        internalPassThroughCableIds,
                        ref interFabricPassThroughCount,
                        interFabricPassThroughCableIds,
                        ref fanoutCandidateCount,
                        fanoutCandidateCableIds,
                        ref propagatedEdgeCount,
                        ref propagatedDownstreamCount);

                List<int> fabricCableIds = new List<int>();

                foreach (RemoteEdgeAccumulator accumulator in accumulators.Values.OrderBy(a => a.RemoteDeviceId, StringComparer.OrdinalIgnoreCase))
                {
                    if (accumulator.TotalCableCount < 2)
                        continue;

                    bool hasPredictableNative = false;

                    foreach (LocalEdgeBucket bucket in accumulator.LocalBuckets.Values.OrderBy(b => b.OwnerLocalDeviceId, StringComparer.OrdinalIgnoreCase))
                    {
                        List<int> distinctIds = bucket.CableIds
                            .Distinct()
                            .OrderBy(x => x)
                            .ToList();

                        if (distinctIds.Count < 2)
                            continue;

                        BundleBuilder bundle = new BundleBuilder();
                        bundle.FabricId = plan.FabricId;
                        bundle.DomainId = plan.DomainId;
                        bundle.FabricEstimatedSpeedGbps = plan.EstimatedCapacityGbps;
                        bundle.OwnerLocalDeviceId = bucket.OwnerLocalDeviceId;
                        bundle.RemoteDeviceId = accumulator.RemoteDeviceId;

                        foreach (string localMember in bucket.LocalMembers.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                            bundle.LocalMemberIds.Add(localMember);

                        foreach (int cableId in distinctIds)
                        {
                            bundle.CableIds.Add(cableId);
                            fabricCableIds.Add(cableId);
                            globallyClaimedCableIds.Add(cableId);
                        }

                        bundle.IsIngressPreferredOwner =
                            string.Equals(bundle.OwnerLocalDeviceId, plan.PreferredIngressSwitchId, StringComparison.OrdinalIgnoreCase);

                        bundle.IsPredictedDownstreamExit = accumulator.IsDomainExit || accumulator.IsExternalDownstream;

                        allBundles.Add(bundle);
                        safeNativeBundleCount++;
                        hasPredictableNative = true;

                        string signature =
                            bundle.FabricId + "|" +
                            bundle.DomainId + "|" +
                            bundle.OwnerLocalDeviceId + "|" +
                            bundle.RemoteDeviceId + "|" +
                            string.Join(",", bundle.CableIds.OrderBy(x => x));

                        if (LoggedBundleSignatures.Add(signature))
                        {
                            LogToFile(
                                "SAFE NATIVE BUNDLE | fabricId=" + bundle.FabricId +
                                " | domainId=" + bundle.DomainId +
                                " | ownerLocal=" + bundle.OwnerLocalDeviceId +
                                " | ownerPreferredIngress=" + bundle.IsIngressPreferredOwner.ToString().ToLowerInvariant() +
                                " | localMembers=[" + string.Join(",", bundle.LocalMemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                                " | remote=" + bundle.RemoteDeviceId +
                                " | domainExit=" + bundle.IsPredictedDownstreamExit.ToString().ToLowerInvariant() +
                                " | cableCount=" + bundle.CableIds.Count.ToString(CultureInfo.InvariantCulture) +
                                " | cableIds=[" + string.Join(",", bundle.CableIds.OrderBy(x => x)) + "]"
                            );
                        }
                    }

                    if (!hasPredictableNative && accumulator.TotalCableCount >= 2)
                    {
                        LogFanoutCandidate(plan, accumulator);
                        foreach (int cableId in accumulator.AllCableIds)
                            fanoutCandidateCableIds.Add(cableId);
                    }
                }

                FabricIdToBundleCableIds[plan.FabricId] = fabricCableIds.Distinct().OrderBy(x => x).ToList();
            }

            SaveDataAutoLACP.UpdateDesiredState(allBundles, allManagedIds);

            int adjacencyPairs = fabrics.Sum(f => Math.Max(0, f.Count - 1));
            string summary =
                "SCAN SUMMARY | registeredSwitches=" + RegisteredSwitches.Count.ToString(CultureInfo.InvariantCulture) +
                " | registeredCables=" + RegisteredCables.Count.ToString(CultureInfo.InvariantCulture) +
                " | saveCables=" + saveCableCount.ToString(CultureInfo.InvariantCulture) +
                " | externalLacpCableIds=" + externalLacpCableIds.Count.ToString(CultureInfo.InvariantCulture) +
                " | liveSwitches=" + liveSwitches.Count.ToString(CultureInfo.InvariantCulture) +
                " | activeFabrics=" + fabrics.Count.ToString(CultureInfo.InvariantCulture) +
                " | routingDomains=" + fabricToDomain.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString(CultureInfo.InvariantCulture) +
                " | adjacentPairs=" + adjacencyPairs.ToString(CultureInfo.InvariantCulture) +
                " | safeNativeBundles=" + safeNativeBundleCount.ToString(CultureInfo.InvariantCulture) +
                " | fanoutCandidates=" + fanoutCandidateCount.ToString(CultureInfo.InvariantCulture) +
                " | propagatedEdges=" + propagatedEdgeCount.ToString(CultureInfo.InvariantCulture) +
                " | propagatedDownstream=" + propagatedDownstreamCount.ToString(CultureInfo.InvariantCulture) +
                " | internalPassThrough=" + internalPassThroughCount.ToString(CultureInfo.InvariantCulture) +
                " | interFabricPassThrough=" + interFabricPassThroughCount.ToString(CultureInfo.InvariantCulture) +
                " | bundleCount=" + allBundles.Count.ToString(CultureInfo.InvariantCulture) +
                " | managedIds=" + allManagedIds.Count.ToString(CultureInfo.InvariantCulture);

            if (!string.Equals(summary, _lastSummary, StringComparison.Ordinal))
            {
                _lastSummary = summary;
                MelonLogger.Msg("[AutoSwitch] " + summary);
                LogToFile(summary);

                foreach (string domainId in fabricToDomain.Values.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    List<string> domainFabrics = fabricToDomain
                        .Where(kvp => string.Equals(kvp.Value, domainId, StringComparison.OrdinalIgnoreCase))
                        .Select(kvp => kvp.Key)
                        .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    LogToFile("ROUTING DOMAIN | id=" + domainId + " | fabrics=[" + string.Join(",", domainFabrics) + "]");
                }

                if (internalPassThroughCableIds.Count > 0)
                    LogToFile("INTERNAL PASS-THROUGH CABLE IDS | [" + string.Join(",", internalPassThroughCableIds.OrderBy(x => x)) + "]");

                if (interFabricPassThroughCableIds.Count > 0)
                    LogToFile("INTER-FABRIC PASS-THROUGH CABLE IDS | [" + string.Join(",", interFabricPassThroughCableIds.OrderBy(x => x)) + "]");

                if (fanoutCandidateCableIds.Count > 0)
                    LogToFile("FANOUT CANDIDATE CABLE IDS | [" + string.Join(",", fanoutCandidateCableIds.OrderBy(x => x)) + "]");

                foreach (FabricRuntimePlan plan in fabricPlans.OrderBy(p => p.FabricId, StringComparer.OrdinalIgnoreCase))
                {
                    List<int> cableIds;
                    int tracked = FabricIdToBundleCableIds.TryGetValue(plan.FabricId, out cableIds)
                        ? cableIds.Distinct().Count()
                        : 0;

                    string members = string.Join(", ",
                        plan.Members
                            .OrderBy(x => x.DeviceName, StringComparer.OrdinalIgnoreCase)
                            .Select(x => x.DeviceName + "@(" +
                                         x.WorldPosition.x.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                                         x.WorldPosition.y.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                                         x.WorldPosition.z.ToString("0.###", CultureInfo.InvariantCulture) + ")[" +
                                         x.ModelName + "]"));

                    LogToFile(
                        "FABRIC | id=" + plan.FabricId +
                        " | domainId=" + plan.DomainId +
                        " | members=" + plan.Members.Count.ToString(CultureInfo.InvariantCulture) +
                        " | estCapacityGbps=" + plan.EstimatedCapacityGbps.ToString("0.##", CultureInfo.InvariantCulture) +
                        " | preferredIngress=" + plan.PreferredIngressSwitchId +
                        " | ingressScore=" + plan.PreferredIngressScore.ToString("0.##", CultureInfo.InvariantCulture) +
                        " | bundledCableIds=" + tracked.ToString(CultureInfo.InvariantCulture) +
                        " | switches=" + members
                    );

                    if (tracked > 0)
                    {
                        LogToFile(
                            "FABRIC BUNDLE CABLE MAP | fabricId=" + plan.FabricId +
                            " | cableIds=[" + string.Join(",", cableIds.Distinct().OrderBy(x => x)) + "]"
                        );
                    }

                    LogFabricPlan(plan);
                }
            }
        }

        private static List<FabricRuntimePlan> BuildFabricRuntimePlans(
            List<List<RegisteredSwitchInfo>> fabrics,
            Dictionary<string, HashSet<string>> fabricIdToMembers,
            Dictionary<string, string> deviceIdToFabricId,
            Dictionary<string, string> fabricToDomain,
            HashSet<int> externalLacpCableIds)
        {
            List<FabricRuntimePlan> plans = new List<FabricRuntimePlan>();

            for (int index = 0; index < fabrics.Count; index++)
            {
                string fabricId = "FABRIC-" + (index + 1).ToString("000", CultureInfo.InvariantCulture);
                List<RegisteredSwitchInfo> fabricMembers = fabrics[index];

                string domainId;
                if (!fabricToDomain.TryGetValue(fabricId, out domainId))
                    domainId = fabricId;

                FabricRuntimePlan plan = new FabricRuntimePlan();
                plan.FabricId = fabricId;
                plan.DomainId = domainId;
                plan.EstimatedCapacityGbps = fabricMembers.Sum(x => x.EstimatedCapacityGbps);

                foreach (RegisteredSwitchInfo member in fabricMembers)
                {
                    plan.Members.Add(member);
                    if (member != null && !string.IsNullOrWhiteSpace(member.DeviceName))
                        plan.MemberIds.Add(member.DeviceName);
                }

                foreach (RegisteredSwitchInfo member in fabricMembers)
                {
                    if (member == null || string.IsNullOrWhiteSpace(member.DeviceName))
                        continue;

                    SwitchTrafficProfile profile = BuildSwitchTrafficProfile(
                        member.DeviceName,
                        fabricId,
                        deviceIdToFabricId,
                        fabricToDomain,
                        externalLacpCableIds);

                    plan.SwitchProfiles[member.DeviceName] = profile;
                }

                SwitchTrafficProfile preferred = plan.SwitchProfiles.Values
                    .OrderByDescending(p => p.IngressScore)
                    .ThenByDescending(p => p.ExternalServerEdgeCount)
                    .ThenByDescending(p => p.ExternalUnknownEdgeCount)
                    .ThenByDescending(p => p.DomainExitEdgeCount)
                    .ThenBy(p => p.DeviceId, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault();

                plan.PreferredIngressSwitchId = preferred != null ? preferred.DeviceId : string.Empty;
                plan.PreferredIngressScore = preferred != null ? preferred.IngressScore : 0f;

                foreach (SwitchTrafficProfile profile in plan.SwitchProfiles.Values)
                {
                    if (profile == null)
                        continue;

                    profile.IsPreferredIngress = string.Equals(
                        profile.DeviceId,
                        plan.PreferredIngressSwitchId,
                        StringComparison.OrdinalIgnoreCase);
                }

                plans.Add(plan);
            }

            BuildDomainPropagationPlans(plans);

            return plans;
        }

        private static void BuildDomainPropagationPlans(List<FabricRuntimePlan> plans)
        {
            Dictionary<string, List<FabricRuntimePlan>> domainGroups = plans
                .GroupBy(p => p.DomainId, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderBy(p => p.FabricId, StringComparer.OrdinalIgnoreCase).ToList(),
                    StringComparer.OrdinalIgnoreCase);

            foreach (KeyValuePair<string, List<FabricRuntimePlan>> kvp in domainGroups)
            {
                string domainId = kvp.Key;
                List<FabricRuntimePlan> domainPlans = kvp.Value;

                float domainCapacity = domainPlans.Sum(p => p.EstimatedCapacityGbps);

                foreach (FabricRuntimePlan plan in domainPlans)
                {
                    plan.DomainEstimatedCapacityGbps = domainCapacity;
                    plan.DomainFabricIds = new List<string>(
                        domainPlans.Select(p => p.FabricId).OrderBy(x => x, StringComparer.OrdinalIgnoreCase));

                    foreach (FabricRuntimePlan peer in domainPlans)
                    {
                        if (string.Equals(peer.FabricId, plan.FabricId, StringComparison.OrdinalIgnoreCase))
                            continue;

                        DomainPropagationIntent intent = new DomainPropagationIntent();
                        intent.DomainId = domainId;
                        intent.FromFabricId = plan.FabricId;
                        intent.ToFabricId = peer.FabricId;
                        intent.FromPreferredIngress = plan.PreferredIngressSwitchId;
                        intent.ToPreferredIngress = peer.PreferredIngressSwitchId;
                        intent.EstimatedBandwidthGbps = Math.Min(plan.EstimatedCapacityGbps, peer.EstimatedCapacityGbps);
                        intent.IntentKind = "intra-domain-fabric-share";
                        plan.OutboundPropagationIntents.Add(intent);
                    }
                }
            }
        }

        private static SwitchTrafficProfile BuildSwitchTrafficProfile(
            string deviceId,
            string currentFabricId,
            Dictionary<string, string> deviceIdToFabricId,
            Dictionary<string, string> fabricToDomain,
            HashSet<int> externalLacpCableIds)
        {
            SwitchTrafficProfile profile = new SwitchTrafficProfile();
            profile.DeviceId = deviceId;

            string currentDomainId;
            if (!fabricToDomain.TryGetValue(currentFabricId, out currentDomainId))
                currentDomainId = currentFabricId;

            foreach (RegisteredCableInfo cable in RegisteredCables.Values.OrderBy(x => x.CableId))
            {
                if (cable == null)
                    continue;

                if (externalLacpCableIds.Contains(cable.CableId))
                    continue;

                string localDeviceId;
                string remoteDeviceId;
                if (!TryResolveRemoteDeviceForSpecificLocal(deviceId, cable, out localDeviceId, out remoteDeviceId))
                    continue;

                if (string.IsNullOrWhiteSpace(remoteDeviceId))
                    continue;

                profile.TouchedCableIds.Add(cable.CableId);

                bool remoteIsKnownSwitch = deviceIdToFabricId.ContainsKey(remoteDeviceId);
                if (remoteIsKnownSwitch)
                {
                    string remoteFabricId = deviceIdToFabricId[remoteDeviceId];
                    string remoteDomainId;
                    if (!fabricToDomain.TryGetValue(remoteFabricId, out remoteDomainId))
                        remoteDomainId = remoteFabricId;

                    if (string.Equals(remoteFabricId, currentFabricId, StringComparison.OrdinalIgnoreCase))
                    {
                        profile.InternalFabricEdgeCount++;
                        continue;
                    }

                    if (string.Equals(remoteDomainId, currentDomainId, StringComparison.OrdinalIgnoreCase))
                    {
                        profile.InterFabricSameDomainEdgeCount++;
                        profile.DomainPeerFabricIds.Add(remoteFabricId);
                        continue;
                    }

                    profile.DomainExitEdgeCount++;
                    continue;
                }

                if (remoteDeviceId.StartsWith("SERVERROOT:", StringComparison.OrdinalIgnoreCase) ||
                    remoteDeviceId.StartsWith("SERVERFAMILY:", StringComparison.OrdinalIgnoreCase))
                {
                    profile.ExternalServerEdgeCount++;
                    continue;
                }

                profile.ExternalUnknownEdgeCount++;
            }

            profile.IngressScore =
                (profile.ExternalServerEdgeCount * 10.0f) +
                (profile.ExternalUnknownEdgeCount * 6.0f) +
                (profile.DomainExitEdgeCount * 5.0f) +
                (profile.InterFabricSameDomainEdgeCount * 3.0f) +
                (profile.InternalFabricEdgeCount * 0.25f);

            return profile;
        }

        private static Dictionary<string, RemoteEdgeAccumulator> BuildRemoteEdgeAccumulators(
            FabricRuntimePlan plan,
            Dictionary<string, string> deviceIdToFabricId,
            Dictionary<string, string> fabricToDomain,
            HashSet<int> externalLacpCableIds,
            ref int internalPassThroughCount,
            HashSet<int> internalPassThroughCableIds,
            ref int interFabricPassThroughCount,
            HashSet<int> interFabricPassThroughCableIds,
            ref int fanoutCandidateCount,
            HashSet<int> fanoutCandidateCableIds,
            ref int propagatedEdgeCount,
            ref int propagatedDownstreamCount)
        {
            Dictionary<string, RemoteEdgeAccumulator> result =
                new Dictionary<string, RemoteEdgeAccumulator>(StringComparer.OrdinalIgnoreCase);

            foreach (RegisteredCableInfo cable in RegisteredCables.Values.OrderBy(x => x.CableId))
            {
                if (cable == null)
                    continue;

                if (externalLacpCableIds.Contains(cable.CableId))
                    continue;

                string localDeviceId;
                string remoteDeviceId;
                if (!TryResolveRemoteDeviceFromLiveCable(cable, plan.MemberIds, out localDeviceId, out remoteDeviceId))
                    continue;

                if (string.IsNullOrWhiteSpace(localDeviceId) || string.IsNullOrWhiteSpace(remoteDeviceId))
                    continue;

                bool remoteIsKnownFabricMember = false;
                bool sameFabric = false;
                bool sameDomain = false;
                bool remoteIsDomainExit = false;

                string remoteFabricId;
                if (deviceIdToFabricId.TryGetValue(remoteDeviceId, out remoteFabricId))
                {
                    remoteIsKnownFabricMember = true;
                    sameFabric = string.Equals(remoteFabricId, plan.FabricId, StringComparison.OrdinalIgnoreCase);

                    string remoteDomainId;
                    if (!fabricToDomain.TryGetValue(remoteFabricId, out remoteDomainId))
                        remoteDomainId = remoteFabricId;

                    sameDomain = string.Equals(remoteDomainId, plan.DomainId, StringComparison.OrdinalIgnoreCase);
                    remoteIsDomainExit = !sameDomain;
                }

                if (remoteIsKnownFabricMember && sameFabric)
                {
                    internalPassThroughCount++;
                    internalPassThroughCableIds.Add(cable.CableId);
                    continue;
                }

                if (remoteIsKnownFabricMember && sameDomain)
                {
                    interFabricPassThroughCount++;
                    interFabricPassThroughCableIds.Add(cable.CableId);

                    propagatedEdgeCount++;

                    DomainPropagationIntent intent = new DomainPropagationIntent();
                    intent.DomainId = plan.DomainId;
                    intent.FromFabricId = plan.FabricId;
                    intent.ToFabricId = remoteFabricId;
                    intent.FromPreferredIngress = plan.PreferredIngressSwitchId;
                    intent.ToPreferredIngress = string.Empty;
                    intent.EstimatedBandwidthGbps = plan.EstimatedCapacityGbps;
                    intent.IntentKind = "inter-fabric-link";
                    intent.CableIds.Add(cable.CableId);
                    plan.OutboundPropagationIntents.Add(intent);

                    continue;
                }

                string key = remoteDeviceId;
                RemoteEdgeAccumulator accumulator;
                if (!result.TryGetValue(key, out accumulator))
                {
                    accumulator = new RemoteEdgeAccumulator();
                    accumulator.FabricId = plan.FabricId;
                    accumulator.DomainId = plan.DomainId;
                    accumulator.RemoteDeviceId = remoteDeviceId;
                    accumulator.IsKnownSwitchRemote = remoteIsKnownFabricMember;
                    accumulator.IsDomainExit = remoteIsDomainExit;
                    accumulator.IsExternalDownstream = !remoteIsKnownFabricMember;
                    result[key] = accumulator;
                }

                LocalEdgeBucket bucket;
                if (!accumulator.LocalBuckets.TryGetValue(localDeviceId, out bucket))
                {
                    bucket = new LocalEdgeBucket();
                    bucket.OwnerLocalDeviceId = localDeviceId;
                    accumulator.LocalBuckets[localDeviceId] = bucket;
                }

                bucket.LocalMembers.Add(localDeviceId);
                bucket.CableIds.Add(cable.CableId);
                accumulator.AllCableIds.Add(cable.CableId);

                if (accumulator.IsDomainExit || accumulator.IsExternalDownstream)
                {
                    propagatedDownstreamCount++;
                }
            }

            foreach (RemoteEdgeAccumulator accumulator in result.Values.OrderBy(x => x.RemoteDeviceId, StringComparer.OrdinalIgnoreCase))
            {
                if (accumulator.TotalCableCount >= 2 && accumulator.LocalBuckets.Count > 1)
                {
                    fanoutCandidateCount++;
                    foreach (int id in accumulator.AllCableIds)
                        fanoutCandidateCableIds.Add(id);
                }

                if (accumulator.IsDomainExit || accumulator.IsExternalDownstream)
                {
                    DomainPropagationIntent intent = new DomainPropagationIntent();
                    intent.DomainId = plan.DomainId;
                    intent.FromFabricId = plan.FabricId;
                    intent.ToFabricId = accumulator.IsKnownSwitchRemote ? "REMOTE-DOMAIN" : "DOWNSTREAM";
                    intent.FromPreferredIngress = plan.PreferredIngressSwitchId;
                    intent.ToPreferredIngress = accumulator.RemoteDeviceId;
                    intent.EstimatedBandwidthGbps = plan.EstimatedCapacityGbps;
                    intent.IntentKind = accumulator.IsDomainExit ? "domain-exit" : "downstream-exit";

                    foreach (int cableId in accumulator.AllCableIds.OrderBy(x => x))
                        intent.CableIds.Add(cableId);

                    plan.OutboundPropagationIntents.Add(intent);
                }
            }

            return result;
        }

        private static void ApplyFabricTags(List<FabricRuntimePlan> plans)
        {
            foreach (FabricRuntimePlan plan in plans)
            {
                int memberCount = plan.Members.Count;

                for (int i = 0; i < memberCount; i++)
                {
                    RegisteredSwitchInfo member = plan.Members[i];
                    if (member == null || member.GameObject == null)
                        continue;

                    try
                    {
                        FabricGroupTag tag = member.GameObject.GetComponent<FabricGroupTag>();
                        if (tag == null)
                            tag = member.GameObject.AddComponent<FabricGroupTag>();

                        tag.FabricId = plan.FabricId;
                        tag.DomainId = plan.DomainId;
                        tag.MemberIndex = i;
                        tag.MemberCount = memberCount;
                        tag.AggregatedBandwidth = plan.EstimatedCapacityGbps;
                        tag.DomainAggregatedBandwidth = plan.DomainEstimatedCapacityGbps;
                        tag.AggregatedPortCount = memberCount;
                        tag.IsPreferredIngress = string.Equals(
                            member.DeviceName,
                            plan.PreferredIngressSwitchId,
                            StringComparison.OrdinalIgnoreCase);
                        tag.PreferredIngressSwitchId = plan.PreferredIngressSwitchId;
                    }
                    catch (Exception ex)
                    {
                        LogToFile("ApplyFabricTags failed for " + member.DeviceName + " | " + ex.Message);
                    }
                }
            }
        }

        private static Dictionary<string, HashSet<string>> BuildFabricAdjacency(
            Dictionary<string, string> deviceIdToFabricId,
            HashSet<int> externalLacpCableIds)
        {
            Dictionary<string, HashSet<string>> adjacency =
                new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (string fabricId in deviceIdToFabricId.Values.Distinct(StringComparer.OrdinalIgnoreCase))
                adjacency[fabricId] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (RegisteredCableInfo cable in RegisteredCables.Values.OrderBy(x => x.CableId))
            {
                if (cable == null)
                    continue;

                if (externalLacpCableIds.Contains(cable.CableId))
                    continue;

                string deviceA = PreserveSwitchIdentity(cable.DeviceA);
                string deviceB = PreserveSwitchIdentity(cable.DeviceB);

                if (string.IsNullOrWhiteSpace(deviceA) || string.IsNullOrWhiteSpace(deviceB))
                    continue;

                string fabricA;
                string fabricB;

                if (!deviceIdToFabricId.TryGetValue(deviceA, out fabricA))
                    continue;

                if (!deviceIdToFabricId.TryGetValue(deviceB, out fabricB))
                    continue;

                if (string.Equals(fabricA, fabricB, StringComparison.OrdinalIgnoreCase))
                    continue;

                adjacency[fabricA].Add(fabricB);
                adjacency[fabricB].Add(fabricA);

                string signature = cable.CableId.ToString(CultureInfo.InvariantCulture) + "|" + fabricA + "|" + fabricB;
                if (LoggedDomainLinkSignatures.Add(signature))
                {
                    LogToFile(
                        "DOMAIN LINK | cableId=" + cable.CableId.ToString(CultureInfo.InvariantCulture) +
                        " | fabricA=" + fabricA +
                        " | fabricB=" + fabricB +
                        " | deviceA=" + deviceA +
                        " | deviceB=" + deviceB
                    );
                }
            }

            return adjacency;
        }

        private static Dictionary<string, string> BuildRoutingDomains(
            Dictionary<string, HashSet<string>> fabricAdjacency,
            List<string> allFabricIds)
        {
            Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int domainIndex = 1;

            foreach (string start in allFabricIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                if (visited.Contains(start))
                    continue;

                string domainId = "DOMAIN-" + domainIndex.ToString("000", CultureInfo.InvariantCulture);
                domainIndex++;

                Queue<string> queue = new Queue<string>();
                queue.Enqueue(start);
                visited.Add(start);

                while (queue.Count > 0)
                {
                    string current = queue.Dequeue();
                    result[current] = domainId;

                    HashSet<string> neighbors;
                    if (!fabricAdjacency.TryGetValue(current, out neighbors))
                        continue;

                    foreach (string neighbor in neighbors)
                    {
                        if (visited.Add(neighbor))
                            queue.Enqueue(neighbor);
                    }
                }
            }

            return result;
        }

        private static void LogFanoutCandidate(FabricRuntimePlan plan, RemoteEdgeAccumulator accumulator)
        {
            string signature =
                plan.FabricId + "|" +
                plan.DomainId + "|" +
                accumulator.RemoteDeviceId + "|" +
                string.Join(",", accumulator.LocalBuckets.Keys.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "|" +
                string.Join(",", accumulator.AllCableIds.OrderBy(x => x));

            if (!LoggedFanoutCandidateSignatures.Add(signature))
                return;

            LogToFile(
                "FANOUT CANDIDATE | fabricId=" + plan.FabricId +
                " | domainId=" + plan.DomainId +
                " | preferredIngress=" + plan.PreferredIngressSwitchId +
                " | remote=" + accumulator.RemoteDeviceId +
                " | localMembers=[" + string.Join(",", accumulator.LocalBuckets.Keys.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                " | cableIds=[" + string.Join(",", accumulator.AllCableIds.OrderBy(x => x)) + "]"
            );
        }

        private static void LogFabricPlan(FabricRuntimePlan plan)
        {
            string signature =
                plan.FabricId + "|" +
                plan.DomainId + "|" +
                plan.PreferredIngressSwitchId + "|" +
                string.Join(",", plan.DomainFabricIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase));

            if (LoggedFabricPlanSignatures.Add(signature))
            {
                LogToFile(
                    "FABRIC SHARE PLAN | fabricId=" + plan.FabricId +
                    " | domainId=" + plan.DomainId +
                    " | preferredIngress=" + plan.PreferredIngressSwitchId +
                    " | ingressScore=" + plan.PreferredIngressScore.ToString("0.##", CultureInfo.InvariantCulture) +
                    " | domainCapacityGbps=" + plan.DomainEstimatedCapacityGbps.ToString("0.##", CultureInfo.InvariantCulture) +
                    " | memberIds=[" + string.Join(",", plan.MemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                    " | domainFabrics=[" + string.Join(",", plan.DomainFabricIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]"
                );
            }

            foreach (SwitchTrafficProfile profile in plan.SwitchProfiles.Values.OrderByDescending(p => p.IngressScore).ThenBy(p => p.DeviceId, StringComparer.OrdinalIgnoreCase))
            {
                LogToFile(
                    "FABRIC MEMBER SCORE | fabricId=" + plan.FabricId +
                    " | device=" + profile.DeviceId +
                    " | preferredIngress=" + profile.IsPreferredIngress.ToString().ToLowerInvariant() +
                    " | ingressScore=" + profile.IngressScore.ToString("0.##", CultureInfo.InvariantCulture) +
                    " | serverEdges=" + profile.ExternalServerEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | unknownEdges=" + profile.ExternalUnknownEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | domainExitEdges=" + profile.DomainExitEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | interFabricSameDomain=" + profile.InterFabricSameDomainEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | internalFabric=" + profile.InternalFabricEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | touchedCables=[" + string.Join(",", profile.TouchedCableIds.OrderBy(x => x)) + "]"
                );
            }

            foreach (DomainPropagationIntent intent in plan.OutboundPropagationIntents
                .OrderBy(i => i.IntentKind, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.ToFabricId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.ToPreferredIngress, StringComparer.OrdinalIgnoreCase))
            {
                string sig =
                    intent.DomainId + "|" +
                    intent.FromFabricId + "|" +
                    intent.ToFabricId + "|" +
                    intent.IntentKind + "|" +
                    string.Join(",", intent.CableIds.OrderBy(x => x));

                if (!LoggedPropagationSignatures.Add(sig))
                    continue;

                LogToFile(
                    "DOMAIN PROPAGATION | domainId=" + intent.DomainId +
                    " | fromFabric=" + intent.FromFabricId +
                    " | toFabric=" + intent.ToFabricId +
                    " | kind=" + intent.IntentKind +
                    " | fromIngress=" + intent.FromPreferredIngress +
                    " | toIngress=" + intent.ToPreferredIngress +
                    " | estGbps=" + intent.EstimatedBandwidthGbps.ToString("0.##", CultureInfo.InvariantCulture) +
                    " | cableIds=[" + string.Join(",", intent.CableIds.OrderBy(x => x)) + "]"
                );
            }
        }

        private void InstallNativePatches()
        {
            try
            {
                Type networkMapType = AccessTools.TypeByName("Il2Cpp.NetworkMap");
                if (networkMapType != null)
                {
                    PatchIfFound(networkMapType, "RegisterSwitch", nameof(NetworkMap_RegisterSwitch_Postfix));
                    PatchIfFound(networkMapType, "RegisterCableConnection", nameof(NetworkMap_RegisterCableConnection_Postfix));
                }
            }
            catch (Exception ex)
            {
                LogToFile("InstallNativePatches failed: " + ex);
            }
        }

        private void PatchIfFound(Type type, string methodName, string postfixName)
        {
            try
            {
                MethodInfo target = AccessTools.Method(type, methodName);
                if (target == null)
                {
                    LogToFile("PATCH MISS | " + type.FullName + "." + methodName);
                    return;
                }

                MethodInfo postfixMethod = typeof(AutoSwitchMod).GetMethod(postfixName, BindingFlags.NonPublic | BindingFlags.Static);
                if (postfixMethod == null)
                {
                    LogToFile("PATCH MISS POSTFIX | " + postfixName);
                    return;
                }

                HarmonyInstance.Patch(target, postfix: new HarmonyMethod(postfixMethod));
                LogPatch("Patched " + type.FullName + "." + methodName);
            }
            catch (Exception ex)
            {
                LogToFile("PATCH FAIL | " + type.FullName + "." + methodName + " | " + ex);
            }
        }

        private static void LogPatch(string message)
        {
            if (LoggedPatches.Add(message))
                LogToFile(message);
        }

        private static void NetworkMap_RegisterSwitch_Postfix(object __0)
        {
            try
            {
                if (__0 == null)
                    return;

                RegisteredSwitchInfo info = BuildRegisteredSwitchInfo(__0);
                if (info == null || string.IsNullOrWhiteSpace(info.DeviceName))
                    return;

                RegisteredSwitches[info.DeviceName] = info;
            }
            catch (Exception ex)
            {
                LogToFile("NetworkMap_RegisterSwitch_Postfix failed: " + ex);
            }
        }

        private static void NetworkMap_RegisterCableConnection_Postfix(
            int __0,
            Vector3 __1,
            Vector3 __2,
            object __3,
            object __4,
            string __5,
            string __6,
            int __7,
            int __8,
            string __9,
            string __10)
        {
            try
            {
                RegisteredCableInfo info = new RegisteredCableInfo();
                info.CableId = __0;
                info.DeviceA = (__5 ?? string.Empty).Trim();
                info.DeviceB = (__6 ?? string.Empty).Trim();
                info.ExtraA = (__9 ?? string.Empty).Trim();
                info.ExtraB = (__10 ?? string.Empty).Trim();
                info.PortA = __7;
                info.PortB = __8;
                info.EndpointObjectA = __3;
                info.EndpointObjectB = __4;
                info.ServerRootKeyA = ResolveServerRootKey(__3, __9, __5);
                info.ServerRootKeyB = ResolveServerRootKey(__4, __10, __6);
                info.EndpointSwitchIdA = ResolveEndpointSwitchId(__3, __5, __9);
                info.EndpointSwitchIdB = ResolveEndpointSwitchId(__4, __6, __10);

                RegisteredCables[__0] = info;
            }
            catch (Exception ex)
            {
                LogToFile("NetworkMap_RegisterCableConnection_Postfix failed: " + ex);
            }
        }

        internal static NetworkSaveData GetNetworkSaveData()
        {
            try
            {
                MainGameManager mgm = MainGameManager.instance;
                if (mgm == null)
                    return null;

                object[] roots =
                {
                    TryGetMemberValue(mgm, "currentSaveData"),
                    TryGetMemberValue(mgm, "saveData"),
                    TryGetMemberValue(mgm, "loadedSaveData"),
                    TryGetMemberValue(mgm, "currentSave"),
                };

                foreach (object root in roots)
                {
                    NetworkSaveData data = ExtractNetworkSaveData(root);
                    if (data != null)
                        return data;
                }
            }
            catch (Exception ex)
            {
                LogToFile("GetNetworkSaveData failed: " + ex);
            }

            return null;
        }

        private static NetworkSaveData ExtractNetworkSaveData(object root)
        {
            if (root == null)
                return null;

            NetworkSaveData direct = root as NetworkSaveData;
            if (direct != null)
                return direct;

            object[] nested =
            {
                TryGetMemberValue(root, "networkSaveData"),
                TryGetMemberValue(root, "networkData"),
                TryGetMemberValue(root, "networkSave"),
            };

            foreach (object candidate in nested)
            {
                NetworkSaveData cast = candidate as NetworkSaveData;
                if (cast != null)
                    return cast;
            }

            return null;
        }

        private static int CountSaveCables(NetworkSaveData networkSaveData)
        {
            try
            {
                if (networkSaveData == null || networkSaveData.cables == null)
                    return 0;

                int count = 0;
                foreach (CableSaveData cable in networkSaveData.cables)
                {
                    if (cable != null)
                        count++;
                }

                return count;
            }
            catch (Exception ex)
            {
                LogToFile("CountSaveCables failed: " + ex);
                return 0;
            }
        }

        private static HashSet<int> GetExternalLacpCableIds(NetworkSaveData networkSaveData)
        {
            HashSet<int> result = new HashSet<int>();

            try
            {
                if (networkSaveData == null || networkSaveData.lacpGroups == null)
                    return result;

                HashSet<int> owned = SaveDataAutoLACP.GetOwnedGroupIdsSnapshot();

                foreach (LACPGroupSaveData group in networkSaveData.lacpGroups)
                {
                    if (group == null)
                        continue;

                    if (owned.Contains(group.groupId))
                        continue;

                    if (group.cableIds == null)
                        continue;

                    foreach (int cableId in group.cableIds)
                        result.Add(cableId);
                }
            }
            catch (Exception ex)
            {
                LogToFile("GetExternalLacpCableIds failed: " + ex);
            }

            return result;
        }

        private static bool TryResolveRemoteDeviceFromLiveCable(
            RegisteredCableInfo cable,
            HashSet<string> localIds,
            out string localDeviceId,
            out string remoteDeviceId)
        {
            localDeviceId = null;
            remoteDeviceId = null;

            if (cable == null)
                return false;

            bool aIsLocal = !string.IsNullOrWhiteSpace(cable.DeviceA) && localIds.Contains(cable.DeviceA);
            bool bIsLocal = !string.IsNullOrWhiteSpace(cable.DeviceB) && localIds.Contains(cable.DeviceB);

            if (aIsLocal && bIsLocal)
                return false;

            if (aIsLocal)
            {
                localDeviceId = cable.DeviceA;
                remoteDeviceId = BuildRemoteIdentity(
                    cable.CableId,
                    localDeviceId,
                    cable.DeviceB,
                    cable.ExtraB,
                    cable.ServerRootKeyB,
                    cable.EndpointSwitchIdB,
                    cable.PortB);
                return !string.IsNullOrWhiteSpace(remoteDeviceId);
            }

            if (bIsLocal)
            {
                localDeviceId = cable.DeviceB;
                remoteDeviceId = BuildRemoteIdentity(
                    cable.CableId,
                    localDeviceId,
                    cable.DeviceA,
                    cable.ExtraA,
                    cable.ServerRootKeyA,
                    cable.EndpointSwitchIdA,
                    cable.PortA);
                return !string.IsNullOrWhiteSpace(remoteDeviceId);
            }

            return false;
        }

        private static bool TryResolveRemoteDeviceForSpecificLocal(
            string specificLocalDeviceId,
            RegisteredCableInfo cable,
            out string localDeviceId,
            out string remoteDeviceId)
        {
            localDeviceId = null;
            remoteDeviceId = null;

            if (cable == null || string.IsNullOrWhiteSpace(specificLocalDeviceId))
                return false;

            if (string.Equals(cable.DeviceA, specificLocalDeviceId, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(cable.DeviceB, specificLocalDeviceId, StringComparison.OrdinalIgnoreCase))
            {
                localDeviceId = cable.DeviceA;
                remoteDeviceId = BuildRemoteIdentity(
                    cable.CableId,
                    localDeviceId,
                    cable.DeviceB,
                    cable.ExtraB,
                    cable.ServerRootKeyB,
                    cable.EndpointSwitchIdB,
                    cable.PortB);
                return !string.IsNullOrWhiteSpace(remoteDeviceId);
            }

            if (string.Equals(cable.DeviceB, specificLocalDeviceId, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(cable.DeviceA, specificLocalDeviceId, StringComparison.OrdinalIgnoreCase))
            {
                localDeviceId = cable.DeviceB;
                remoteDeviceId = BuildRemoteIdentity(
                    cable.CableId,
                    localDeviceId,
                    cable.DeviceA,
                    cable.ExtraA,
                    cable.ServerRootKeyA,
                    cable.EndpointSwitchIdA,
                    cable.PortA);
                return !string.IsNullOrWhiteSpace(remoteDeviceId);
            }

            return false;
        }

        private static string BuildRemoteIdentity(
            int cableId,
            string localDeviceId,
            string deviceName,
            string extraName,
            string serverRootKey,
            string endpointSwitchId,
            int port)
        {
            string preservedDevice = PreserveSwitchIdentity(deviceName);
            string preservedEndpointSwitch = PreserveSwitchIdentity(endpointSwitchId);
            string normalizedExtra = NormalizeRuntimeIdentity(extraName);

            if (!string.IsNullOrWhiteSpace(preservedDevice))
            {
                LogRemoteResolution(cableId, localDeviceId, "deviceId", preservedDevice, deviceName, extraName, endpointSwitchId, serverRootKey, port);
                return preservedDevice;
            }

            if (!string.IsNullOrWhiteSpace(preservedEndpointSwitch))
            {
                LogRemoteResolution(cableId, localDeviceId, "endpointSwitch", preservedEndpointSwitch, deviceName, extraName, endpointSwitchId, serverRootKey, port);
                return preservedEndpointSwitch;
            }

            if (!string.IsNullOrWhiteSpace(serverRootKey))
            {
                LogRemoteResolution(cableId, localDeviceId, "serverRoot", serverRootKey, deviceName, extraName, endpointSwitchId, serverRootKey, port);
                return serverRootKey;
            }

            if (!string.IsNullOrWhiteSpace(normalizedExtra))
            {
                string value = normalizedExtra + "#P" + port.ToString(CultureInfo.InvariantCulture);
                LogRemoteResolution(cableId, localDeviceId, "fallbackExtra", value, deviceName, extraName, endpointSwitchId, serverRootKey, port);
                return value;
            }

            return string.Empty;
        }

        private static void LogRemoteResolution(
            int cableId,
            string localDeviceId,
            string kind,
            string chosen,
            string deviceName,
            string extraName,
            string endpointSwitchId,
            string serverRootKey,
            int port)
        {
            string signature =
                cableId.ToString(CultureInfo.InvariantCulture) + "|" +
                localDeviceId + "|" +
                kind + "|" +
                chosen;

            if (!LoggedRemoteResolutionSignatures.Add(signature))
                return;

            LogToFile(
                "REMOTE RESOLVE | cableId=" + cableId.ToString(CultureInfo.InvariantCulture) +
                " | local=" + localDeviceId +
                " | kind=" + kind +
                " | chosen=" + chosen +
                " | rawDevice=" + (deviceName ?? string.Empty) +
                " | rawExtra=" + (extraName ?? string.Empty) +
                " | endpointSwitchId=" + (endpointSwitchId ?? string.Empty) +
                " | serverRoot=" + (serverRootKey ?? string.Empty) +
                " | port=" + port.ToString(CultureInfo.InvariantCulture)
            );
        }

        private static string ResolveEndpointSwitchId(object endpointObj, string deviceName, string extraName)
        {
            try
            {
                object[] candidates =
                {
                    endpointObj,
                    TryGetMemberValue(endpointObj, "connectedSwitch"),
                    TryGetMemberValue(endpointObj, "networkSwitch"),
                    TryGetMemberValue(endpointObj, "switchRef"),
                    TryGetMemberValue(endpointObj, "parentSwitch"),
                    TryGetMemberValue(endpointObj, "switchData")
                };

                foreach (object candidate in candidates)
                {
                    string discovered = DiscoverBestDeviceName(candidate);
                    if (!string.IsNullOrWhiteSpace(discovered) && IsSwitchLikeIdentity(discovered))
                        return discovered;

                    GameObject go = DiscoverGameObject(candidate);
                    string fromGo = ResolveSwitchIdFromGameObject(go);
                    if (!string.IsNullOrWhiteSpace(fromGo))
                        return fromGo;
                }

                GameObject endpointGo = DiscoverGameObject(endpointObj);
                string fromEndpointGo = ResolveSwitchIdFromGameObject(endpointGo);
                if (!string.IsNullOrWhiteSpace(fromEndpointGo))
                    return fromEndpointGo;

                string raw = !string.IsNullOrWhiteSpace(deviceName) ? deviceName : extraName;
                if (IsSwitchLikeIdentity(raw))
                    return raw;
            }
            catch (Exception ex)
            {
                LogToFile("ResolveEndpointSwitchId failed: " + ex.Message);
            }

            return string.Empty;
        }

        private static string ResolveSwitchIdFromGameObject(GameObject go)
        {
            if (go == null)
                return string.Empty;

            try
            {
                Transform current = go.transform;
                int depth = 0;

                while (current != null && depth < 20)
                {
                    string discovered = DiscoverBestDeviceName(current.gameObject);
                    if (!string.IsNullOrWhiteSpace(discovered) && IsSwitchLikeIdentity(discovered))
                        return discovered;

                    string currentName = NormalizeCloneName(current.name);
                    if (AllowedSwitchModels.Contains(currentName))
                    {
                        string modelNamed = DiscoverBestDeviceName(current.gameObject);
                        if (!string.IsNullOrWhiteSpace(modelNamed) && IsSwitchLikeIdentity(modelNamed))
                            return modelNamed;
                    }

                    current = current.parent;
                    depth++;
                }
            }
            catch { }

            return string.Empty;
        }

        private static string PreserveSwitchIdentity(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            string trimmed = raw.Trim();

            if (LooksLikeSwitchIdentity(trimmed))
                return trimmed;

            return string.Empty;
        }

        private static bool IsSwitchLikeIdentity(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string normalized = NormalizeCloneName(value);

            return normalized.StartsWith("Switch32xQSFP", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("Switch4xSFP", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("Switch16CU", StringComparison.OrdinalIgnoreCase);
        }

        private static bool LooksLikeSwitchIdentity(string raw)
        {
            return IsSwitchLikeIdentity(raw);
        }

        internal static bool GroupTouchesManagedIds(LACPGroupSaveData group, HashSet<string> managedIds)
        {
            if (group == null)
                return false;

            return (!string.IsNullOrWhiteSpace(group.deviceA) && managedIds.Contains(group.deviceA)) ||
                   (!string.IsNullOrWhiteSpace(group.deviceB) && managedIds.Contains(group.deviceB));
        }

        private static RegisteredSwitchInfo BuildRegisteredSwitchInfo(object networkSwitchObj)
        {
            if (networkSwitchObj == null)
                return null;

            string deviceName = DiscoverBestDeviceName(networkSwitchObj);
            string modelName = DiscoverBestModelName(networkSwitchObj);

            GameObject go = DiscoverGameObject(networkSwitchObj);
            Vector3 pos = go != null ? go.transform.position : Vector3.zero;

            RegisteredSwitchInfo info = new RegisteredSwitchInfo();
            info.NetworkSwitchObject = networkSwitchObj;
            info.GameObject = go;
            info.DeviceName = deviceName;
            info.ModelName = modelName;
            info.WorldPosition = pos;
            info.HasWorldPosition = go != null;
            info.EstimatedCapacityGbps = EstimateCapacityFromModel(modelName);

            return info;
        }

        private void RefreshRegisteredSwitchPositions()
        {
            foreach (RegisteredSwitchInfo info in RegisteredSwitches.Values)
            {
                if (info == null)
                    continue;

                if (info.GameObject == null && info.NetworkSwitchObject != null)
                    info.GameObject = DiscoverGameObject(info.NetworkSwitchObject);

                if (info.GameObject != null)
                {
                    info.WorldPosition = info.GameObject.transform.position;
                    info.HasWorldPosition = true;
                }

                if (string.IsNullOrWhiteSpace(info.ModelName))
                    info.ModelName = DiscoverBestModelName(info.NetworkSwitchObject);

                info.EstimatedCapacityGbps = EstimateCapacityFromModel(info.ModelName);
            }
        }

        private static bool IsAllowedSwitchModel(string modelName)
        {
            if (string.IsNullOrWhiteSpace(modelName))
                return false;

            return AllowedSwitchModels.Contains(NormalizeCloneName(modelName));
        }

        private static bool IsLikelyPlacedSwitch(RegisteredSwitchInfo info)
        {
            if (info == null || !info.HasWorldPosition)
                return false;

            float xzRadius = Mathf.Sqrt((info.WorldPosition.x * info.WorldPosition.x) + (info.WorldPosition.z * info.WorldPosition.z));
            if (xzRadius <= OriginXZRejectRadius)
                return false;

            if (info.WorldPosition.y < MinUsefulWorldY || info.WorldPosition.y > MaxUsefulWorldY)
                return false;

            return true;
        }

        private static float EstimateCapacityFromModel(string modelName)
        {
            string normalized = NormalizeCloneName(modelName);

            if (normalized.Equals("Switch32xQSFP", StringComparison.OrdinalIgnoreCase))
                return 32f * 40f;

            if (normalized.Equals("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase))
                return (4f * 40f) + (16f * 10f);

            if (normalized.Equals("Switch4xSFP", StringComparison.OrdinalIgnoreCase))
                return 4f * 10f;

            if (normalized.Equals("Switch16CU", StringComparison.OrdinalIgnoreCase))
                return 16f * 1f;

            return 0f;
        }

        private static List<List<RegisteredSwitchInfo>> BuildRegistryFabrics(List<RegisteredSwitchInfo> switches)
        {
            List<List<RegisteredSwitchInfo>> groups = new List<List<RegisteredSwitchInfo>>();
            HashSet<RegisteredSwitchInfo> remaining = new HashSet<RegisteredSwitchInfo>(switches);

            while (remaining.Count > 0)
            {
                RegisteredSwitchInfo seed = remaining.First();
                remaining.Remove(seed);

                List<RegisteredSwitchInfo> group = new List<RegisteredSwitchInfo>();
                Queue<RegisteredSwitchInfo> queue = new Queue<RegisteredSwitchInfo>();
                queue.Enqueue(seed);

                while (queue.Count > 0)
                {
                    RegisteredSwitchInfo current = queue.Dequeue();
                    group.Add(current);

                    List<RegisteredSwitchInfo> neighbors = remaining
                        .Where(other => AreAdjacent(current, other))
                        .ToList();

                    foreach (RegisteredSwitchInfo neighbor in neighbors)
                    {
                        remaining.Remove(neighbor);
                        queue.Enqueue(neighbor);
                    }
                }

                groups.Add(group);
            }

            return groups;
        }

        private static bool AreAdjacent(RegisteredSwitchInfo a, RegisteredSwitchInfo b)
        {
            if (a == null || b == null || !a.HasWorldPosition || !b.HasWorldPosition)
                return false;

            float dx = Mathf.Abs(a.WorldPosition.x - b.WorldPosition.x);
            float dz = Mathf.Abs(a.WorldPosition.z - b.WorldPosition.z);

            if (dx > SameFabricXZTolerance || dz > SameFabricXZTolerance)
                return false;

            float actualStep = Mathf.Abs(a.WorldPosition.y - b.WorldPosition.y);
            return Mathf.Abs(actualStep - FallbackExpectedStep) <= AdjacentYTolerance;
        }

        private static string ResolveServerRootKey(object endpointObj, string extraName, string deviceName)
        {
            GameObject endpointGo = DiscoverGameObject(endpointObj);
            if (endpointGo != null)
            {
                Transform current = endpointGo.transform;
                int depth = 0;

                while (current != null && depth < 16)
                {
                    string n = NormalizeCloneName(current.name);
                    if (n.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                        return "SERVERROOT:" + n + "#" + current.gameObject.GetInstanceID().ToString(CultureInfo.InvariantCulture);

                    current = current.parent;
                    depth++;
                }
            }

            string raw = !string.IsNullOrWhiteSpace(extraName) ? extraName : deviceName;
            if (!string.IsNullOrWhiteSpace(raw) && raw.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                return "SERVERFAMILY:" + NormalizeRuntimeIdentity(raw);

            return string.Empty;
        }

        private static GameObject DiscoverGameObject(object obj)
        {
            if (obj == null)
                return null;

            try
            {
                GameObject go = obj as GameObject;
                if (go != null)
                    return go;
            }
            catch { }

            try
            {
                Component comp = obj as Component;
                if (comp != null)
                    return comp.gameObject;
            }
            catch { }

            try
            {
                object go = TryGetMemberValue(obj, "gameObject");
                GameObject gameObject = go as GameObject;
                if (gameObject != null)
                    return gameObject;
            }
            catch { }

            try
            {
                object transform = TryGetMemberValue(obj, "transform");
                Transform t = transform as Transform;
                if (t != null)
                    return t.gameObject;
            }
            catch { }

            return null;
        }

        private static string DiscoverBestDeviceName(object obj)
        {
            if (obj == null)
                return string.Empty;

            string[] candidates =
            {
                "deviceName", "DeviceName", "switchId", "SwitchId", "networkId", "NetworkId", "id", "Id"
            };

            foreach (string name in candidates)
            {
                object val = TryGetMemberValue(obj, name);
                string s = val as string;
                if (!string.IsNullOrWhiteSpace(s))
                    return s.Trim();
            }

            Type type;
            try
            {
                type = obj.GetType();
            }
            catch
            {
                return string.Empty;
            }

            foreach (FieldInfo f in SafeFields(type))
            {
                if (f.FieldType != typeof(string))
                    continue;

                string lower = f.Name.ToLowerInvariant();
                if (!(lower.Contains("device") || lower.Contains("switch") || lower.Contains("network") || lower == "id"))
                    continue;

                try
                {
                    string s = f.GetValue(obj) as string;
                    if (!string.IsNullOrWhiteSpace(s))
                        return s.Trim();
                }
                catch { }
            }

            foreach (PropertyInfo p in SafeProperties(type))
            {
                if (p.PropertyType != typeof(string) || !p.CanRead)
                    continue;

                string lower = p.Name.ToLowerInvariant();
                if (!(lower.Contains("device") || lower.Contains("switch") || lower.Contains("network") || lower == "id"))
                    continue;

                try
                {
                    string s = p.GetValue(obj, null) as string;
                    if (!string.IsNullOrWhiteSpace(s))
                        return s.Trim();
                }
                catch { }
            }

            return string.Empty;
        }

        private static string DiscoverBestModelName(object obj)
        {
            GameObject go = DiscoverGameObject(obj);
            if (go != null)
            {
                string normalized = NormalizeCloneName(go.name);
                if (AllowedSwitchModels.Contains(normalized))
                    return normalized;
            }

            foreach (string name in AllowedSwitchModels)
            {
                if (go != null && go.name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                    return name;
            }

            return go != null ? NormalizeCloneName(go.name) : string.Empty;
        }

        private static IEnumerable<FieldInfo> SafeFields(Type t)
        {
            try { return t.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); }
            catch { return new FieldInfo[0]; }
        }

        private static IEnumerable<PropertyInfo> SafeProperties(Type t)
        {
            try { return t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); }
            catch { return new PropertyInfo[0]; }
        }

        private static object TryGetMemberValue(object target, string memberName)
        {
            if (target == null || string.IsNullOrWhiteSpace(memberName))
                return null;

            try
            {
                Type t = target.GetType();

                PropertyInfo prop = t.GetProperty(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (prop != null && prop.CanRead)
                    return prop.GetValue(target, null);

                FieldInfo field = t.GetField(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (field != null)
                    return field.GetValue(target);
            }
            catch { }

            return null;
        }

        private static string NormalizeCloneName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return string.Empty;

            string result = name.Trim();
            if (result.EndsWith("(Clone)", StringComparison.OrdinalIgnoreCase))
                result = result.Substring(0, result.Length - "(Clone)".Length).Trim();

            return result;
        }

        private static string NormalizeRuntimeIdentity(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            return TrailingRuntimeIdRegex.Replace(raw.Trim(), "");
        }

        internal static void LogToFile(string message)
        {
            try
            {
                Directory.CreateDirectory(DebugFolderPath);
                File.AppendAllText(
                    DebugLogPath,
                    "[" + DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) + "] " + message + Environment.NewLine
                );
            }
            catch { }
        }
    }

    internal static class SaveDataAutoLACP
    {
        private static bool _bootstrapStarted;
        private static bool _regroupQueued;
        private static float _lastRunRealtime;
        private static string _lastAppliedSignature = string.Empty;
        private static bool _hasAppliedAtLeastOnce;

        private const float InitialDelaySeconds = 8.0f;
        private const float MinDelaySeconds = 0.35f;
        private const float MinGapBetweenRunsSeconds = 0.50f;

        private static readonly object _stateLock = new object();
        private static List<BundleBuilder> _desiredBundles = new List<BundleBuilder>();
        private static HashSet<string> _desiredManagedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<int> _ownedGroupIds = new HashSet<int>();

        internal static void StartSafeBootstrap()
        {
            if (_bootstrapStarted)
                return;

            _bootstrapStarted = true;
            MelonLogger.Msg("[AutoSwitch] Live Auto-LACP bootstrap armed.");
            MelonCoroutines.Start(BootstrapRoutine());
        }

        internal static void ResetForScene()
        {
            _bootstrapStarted = false;
            _regroupQueued = false;
            _lastRunRealtime = 0f;
            _lastAppliedSignature = string.Empty;
            _hasAppliedAtLeastOnce = false;

            lock (_stateLock)
            {
                _desiredBundles = new List<BundleBuilder>();
                _desiredManagedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            _ownedGroupIds.Clear();
        }

        internal static HashSet<int> GetOwnedGroupIdsSnapshot()
        {
            return new HashSet<int>(_ownedGroupIds);
        }

        private static IEnumerator BootstrapRoutine()
        {
            yield return new WaitForSeconds(InitialDelaySeconds);

            if (_hasAppliedAtLeastOnce)
            {
                AutoSwitchMod.LogToFile("AUTO LACP | bootstrap skipped because desired state was already applied.");
                yield break;
            }

            QueueRegroup("safe bootstrap");
        }

        internal static void UpdateDesiredState(List<BundleBuilder> bundles, HashSet<string> managedIds)
        {
            if (bundles == null)
                bundles = new List<BundleBuilder>();

            if (managedIds == null)
                managedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            string signature = BuildDesiredSignature(bundles, managedIds);

            lock (_stateLock)
            {
                _desiredBundles = bundles
                    .Select(CloneBundle)
                    .OrderBy(b => b.FabricId, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(b => b.OwnerLocalDeviceId, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(b => b.RemoteDeviceId, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                _desiredManagedIds = new HashSet<string>(managedIds, StringComparer.OrdinalIgnoreCase);
            }

            if (!string.Equals(signature, _lastAppliedSignature, StringComparison.Ordinal))
            {
                _lastAppliedSignature = signature;
                QueueRegroup("desired live bundles changed");
            }
        }

        internal static void QueueRegroup(string reason)
        {
            if (_regroupQueued)
                return;

            _regroupQueued = true;
            MelonLogger.Msg("[AutoSwitch] Live Auto-LACP queued: " + reason);
            MelonCoroutines.Start(RunQueuedRegroup());
        }

        private static IEnumerator RunQueuedRegroup()
        {
            yield return new WaitForSeconds(MinDelaySeconds);

            float now = Time.realtimeSinceStartup;
            if (now - _lastRunRealtime < MinGapBetweenRunsSeconds)
            {
                _regroupQueued = false;
                yield return new WaitForSeconds(MinGapBetweenRunsSeconds);
                QueueRegroup("debounced");
                yield break;
            }

            _regroupQueued = false;
            _lastRunRealtime = Time.realtimeSinceStartup;

            try
            {
                RebuildAutoLacpGroups();
                _hasAppliedAtLeastOnce = true;
            }
            catch (Exception ex)
            {
                MelonLogger.Warning("[AutoSwitch] Live Auto-LACP regroup failed: " + ex);
                AutoSwitchMod.LogToFile("Live Auto-LACP regroup failed: " + ex);
            }
        }

        internal static void RebuildAutoLacpGroups()
        {
            List<BundleBuilder> desiredBundles;
            HashSet<string> desiredManagedIds;

            lock (_stateLock)
            {
                desiredBundles = _desiredBundles.Select(CloneBundle).ToList();
                desiredManagedIds = new HashSet<string>(_desiredManagedIds, StringComparer.OrdinalIgnoreCase);
            }

            NetworkMap networkMap = NetworkMap.instance;
            NetworkSaveData networkSaveData = AutoSwitchMod.GetNetworkSaveData();

            if (networkMap == null)
            {
                AutoSwitchMod.LogToFile("AUTO LACP | skipped, NetworkMap.instance is null.");
                return;
            }

            List<int> removedGroupIds = new List<int>();

            foreach (int ownedId in _ownedGroupIds.OrderBy(x => x))
            {
                try
                {
                    networkMap.RemoveLACPGroup(ownedId);
                    removedGroupIds.Add(ownedId);
                }
                catch (Exception ex)
                {
                    AutoSwitchMod.LogToFile("AUTO LACP | Remove owned group failed for " +
                                            ownedId.ToString(CultureInfo.InvariantCulture) + " | " + ex.Message);
                }
            }

            _ownedGroupIds.Clear();

            if (networkSaveData != null && networkSaveData.lacpGroups != null)
            {
                foreach (LACPGroupSaveData existing in networkSaveData.lacpGroups)
                {
                    if (AutoSwitchMod.GroupTouchesManagedIds(existing, desiredManagedIds))
                    {
                        try
                        {
                            networkMap.RemoveLACPGroup(existing.groupId);
                            removedGroupIds.Add(existing.groupId);
                        }
                        catch (Exception ex)
                        {
                            AutoSwitchMod.LogToFile("AUTO LACP | Remove save group failed for " +
                                                    existing.groupId.ToString(CultureInfo.InvariantCulture) + " | " + ex.Message);
                        }
                    }
                }
            }

            List<LACPGroupSaveData> rebuiltSaveGroups = new List<LACPGroupSaveData>();

            if (networkSaveData != null && networkSaveData.lacpGroups != null)
            {
                foreach (LACPGroupSaveData existing in networkSaveData.lacpGroups)
                {
                    if (!AutoSwitchMod.GroupTouchesManagedIds(existing, desiredManagedIds))
                        rebuiltSaveGroups.Add(existing);
                }
            }

            foreach (BundleBuilder bundle in desiredBundles)
            {
                List<int> distinctIds = bundle.CableIds.Distinct().OrderBy(x => x).ToList();
                if (distinctIds.Count < 2)
                    continue;

                if (bundle.LocalMemberIds.Count != 1)
                {
                    AutoSwitchMod.LogToFile(
                        "AUTO LACP | skipped multi-local bundle | fabricId=" + bundle.FabricId +
                        " | domainId=" + bundle.DomainId +
                        " | ownerLocal=" + bundle.OwnerLocalDeviceId +
                        " | localMembers=[" + string.Join(",", bundle.LocalMemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                        " | deviceB=" + bundle.RemoteDeviceId +
                        " | cableIds=[" + string.Join(",", distinctIds) + "]"
                    );
                    continue;
                }

                try
                {
                    Il2CppSystem.Collections.Generic.List<int> il2cppCableIds = new Il2CppSystem.Collections.Generic.List<int>();
                    foreach (int cableId in distinctIds)
                        il2cppCableIds.Add(cableId);

                    int groupId = networkMap.CreateLACPGroup(bundle.OwnerLocalDeviceId, bundle.RemoteDeviceId, il2cppCableIds);
                    _ownedGroupIds.Add(groupId);

                    if (networkSaveData != null)
                    {
                        try
                        {
                            LACPGroupSaveData saveGroup = new LACPGroupSaveData();
                            saveGroup.groupId = groupId;
                            saveGroup.deviceA = bundle.OwnerLocalDeviceId;
                            saveGroup.deviceB = bundle.RemoteDeviceId;
                            saveGroup.cableIds = il2cppCableIds;
                            rebuiltSaveGroups.Add(saveGroup);
                        }
                        catch (Exception ex)
                        {
                            AutoSwitchMod.LogToFile(
                                "AUTO LACP | save group construct failed for fabricId=" + bundle.FabricId +
                                " | ownerLocal=" + bundle.OwnerLocalDeviceId +
                                " | remote=" + bundle.RemoteDeviceId +
                                " | " + ex.Message
                            );
                        }
                    }

                    AutoSwitchMod.LogToFile(
                        "AUTO LACP | created groupId=" + groupId.ToString(CultureInfo.InvariantCulture) +
                        " | fabricId=" + bundle.FabricId +
                        " | domainId=" + bundle.DomainId +
                        " | ownerLocal=" + bundle.OwnerLocalDeviceId +
                        " | ownerPreferredIngress=" + bundle.IsIngressPreferredOwner.ToString().ToLowerInvariant() +
                        " | localMembers=[" + string.Join(",", bundle.LocalMemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                        " | deviceB=" + bundle.RemoteDeviceId +
                        " | downstreamExit=" + bundle.IsPredictedDownstreamExit.ToString().ToLowerInvariant() +
                        " | cableIds=[" + string.Join(",", distinctIds) + "]"
                    );
                }
                catch (Exception ex)
                {
                    AutoSwitchMod.LogToFile(
                        "AUTO LACP | CreateLACPGroup failed | fabricId=" + bundle.FabricId +
                        " | domainId=" + bundle.DomainId +
                        " | ownerLocal=" + bundle.OwnerLocalDeviceId +
                        " | localMembers=[" + string.Join(",", bundle.LocalMemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                        " | deviceB=" + bundle.RemoteDeviceId +
                        " | cableIds=[" + string.Join(",", distinctIds) + "]" +
                        " | " + ex
                    );
                }
            }

            if (networkSaveData != null)
            {
                try
                {
                    Il2CppSystem.Collections.Generic.List<LACPGroupSaveData> saveList =
                        new Il2CppSystem.Collections.Generic.List<LACPGroupSaveData>();

                    foreach (LACPGroupSaveData group in rebuiltSaveGroups)
                        saveList.Add(group);

                    networkSaveData.lacpGroups = saveList;
                }
                catch (Exception ex)
                {
                    AutoSwitchMod.LogToFile("AUTO LACP | failed updating networkSaveData.lacpGroups | " + ex);
                }
            }

            AutoSwitchMod.LogToFile(
                "AUTO LACP | regroup complete | removed=" +
                removedGroupIds.Distinct().Count().ToString(CultureInfo.InvariantCulture) +
                " | desiredBundles=" + desiredBundles.Count.ToString(CultureInfo.InvariantCulture) +
                " | ownedNow=" + _ownedGroupIds.Count.ToString(CultureInfo.InvariantCulture)
            );
        }

        private static string BuildDesiredSignature(List<BundleBuilder> bundles, HashSet<string> managedIds)
        {
            List<string> parts = new List<string>();

            foreach (BundleBuilder bundle in bundles
                .OrderBy(b => b.FabricId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(b => b.OwnerLocalDeviceId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(b => b.RemoteDeviceId, StringComparer.OrdinalIgnoreCase))
            {
                parts.Add(
                    bundle.FabricId + "|" +
                    bundle.DomainId + "|" +
                    bundle.OwnerLocalDeviceId + "|" +
                    bundle.RemoteDeviceId + "|" +
                    bundle.IsIngressPreferredOwner.ToString().ToLowerInvariant() + "|" +
                    bundle.IsPredictedDownstreamExit.ToString().ToLowerInvariant() + "|" +
                    string.Join(",", bundle.LocalMemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "|" +
                    string.Join(",", bundle.CableIds.Distinct().OrderBy(x => x))
                );
            }

            parts.Add("managed:" + string.Join(",", managedIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)));

            return string.Join(" || ", parts);
        }

        private static BundleBuilder CloneBundle(BundleBuilder source)
        {
            BundleBuilder clone = new BundleBuilder();
            clone.FabricId = source.FabricId;
            clone.DomainId = source.DomainId;
            clone.FabricEstimatedSpeedGbps = source.FabricEstimatedSpeedGbps;
            clone.OwnerLocalDeviceId = source.OwnerLocalDeviceId;
            clone.RemoteDeviceId = source.RemoteDeviceId;
            clone.IsIngressPreferredOwner = source.IsIngressPreferredOwner;
            clone.IsPredictedDownstreamExit = source.IsPredictedDownstreamExit;

            foreach (string id in source.LocalMemberIds)
                clone.LocalMemberIds.Add(id);

            foreach (int cableId in source.CableIds)
                clone.CableIds.Add(cableId);

            return clone;
        }
    }

    public sealed class FabricGroupTag : MonoBehaviour
    {
        public FabricGroupTag(IntPtr ptr) : base(ptr) { }

        public string FabricId { get; set; }
        public string DomainId { get; set; }
        public int MemberIndex { get; set; }
        public int MemberCount { get; set; }
        public float AggregatedBandwidth { get; set; }
        public float DomainAggregatedBandwidth { get; set; }
        public int AggregatedPortCount { get; set; }
        public bool IsPreferredIngress { get; set; }
        public string PreferredIngressSwitchId { get; set; }
    }

    internal sealed class RegisteredSwitchInfo
    {
        public object NetworkSwitchObject;
        public GameObject GameObject;
        public string DeviceName;
        public string ModelName;
        public Vector3 WorldPosition;
        public bool HasWorldPosition;
        public float EstimatedCapacityGbps;
    }

    internal sealed class RegisteredCableInfo
    {
        public int CableId;
        public string DeviceA;
        public string DeviceB;
        public string ExtraA;
        public string ExtraB;
        public int PortA;
        public int PortB;

        public object EndpointObjectA;
        public object EndpointObjectB;

        public string ServerRootKeyA;
        public string ServerRootKeyB;

        public string EndpointSwitchIdA;
        public string EndpointSwitchIdB;
    }

    internal sealed class BundleBuilder
    {
        public string FabricId;
        public string DomainId;
        public float FabricEstimatedSpeedGbps;
        public string OwnerLocalDeviceId;
        public string RemoteDeviceId;
        public bool IsIngressPreferredOwner;
        public bool IsPredictedDownstreamExit;
        public readonly List<int> CableIds = new List<int>();
        public readonly HashSet<string> LocalMemberIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    internal sealed class LocalEdgeBucket
    {
        public string OwnerLocalDeviceId;
        public readonly List<int> CableIds = new List<int>();
        public readonly HashSet<string> LocalMembers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    internal sealed class RemoteEdgeAccumulator
    {
        public string FabricId;
        public string DomainId;
        public string RemoteDeviceId;
        public bool IsKnownSwitchRemote;
        public bool IsDomainExit;
        public bool IsExternalDownstream;

        public readonly Dictionary<string, LocalEdgeBucket> LocalBuckets =
            new Dictionary<string, LocalEdgeBucket>(StringComparer.OrdinalIgnoreCase);

        public readonly HashSet<int> AllCableIds = new HashSet<int>();

        public int TotalCableCount
        {
            get { return AllCableIds.Count; }
        }
    }

    internal sealed class FabricRuntimePlan
    {
        public string FabricId;
        public string DomainId;
        public float EstimatedCapacityGbps;
        public float DomainEstimatedCapacityGbps;
        public string PreferredIngressSwitchId;
        public float PreferredIngressScore;

        public readonly List<RegisteredSwitchInfo> Members = new List<RegisteredSwitchInfo>();
        public readonly HashSet<string> MemberIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, SwitchTrafficProfile> SwitchProfiles =
            new Dictionary<string, SwitchTrafficProfile>(StringComparer.OrdinalIgnoreCase);

        public List<string> DomainFabricIds = new List<string>();
        public readonly List<DomainPropagationIntent> OutboundPropagationIntents =
            new List<DomainPropagationIntent>();
    }

    internal sealed class SwitchTrafficProfile
    {
        public string DeviceId;
        public int InternalFabricEdgeCount;
        public int InterFabricSameDomainEdgeCount;
        public int DomainExitEdgeCount;
        public int ExternalServerEdgeCount;
        public int ExternalUnknownEdgeCount;
        public float IngressScore;
        public bool IsPreferredIngress;

        public readonly HashSet<int> TouchedCableIds = new HashSet<int>();
        public readonly HashSet<string> DomainPeerFabricIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    internal sealed class DomainPropagationIntent
    {
        public string DomainId;
        public string FromFabricId;
        public string ToFabricId;
        public string FromPreferredIngress;
        public string ToPreferredIngress;
        public float EstimatedBandwidthGbps;
        public string IntentKind;
        public readonly List<int> CableIds = new List<int>();
    }
}