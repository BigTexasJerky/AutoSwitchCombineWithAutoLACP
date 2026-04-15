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

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "1.0.0", "Big Texas Jerky")]
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

        private const bool DebugLoggingEnabled = false;

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

        private static readonly HashSet<int> SyntheticRegisteredCableIds =
            new HashSet<int>();

        private static readonly HashSet<string> LoggedSyntheticRegistrationSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, float> FabricIdToSpeed =
            new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, List<int>> FabricIdToBundleCableIds =
            new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, HashSet<string>> FabricIdToMemberIdsSnapshot =
            new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, string> SwitchIdToFabricIdSnapshot =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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
        internal static string _lastFabricMembershipSignature = string.Empty;
        internal static bool _fabricChurnCooldownPending = false;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();

            if (DebugLoggingEnabled)
            {
                Directory.CreateDirectory(DebugFolderPath);
                File.WriteAllText(
                    DebugLogPath,
                    "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "] AutoSwitch 1.0.0 debug log started." + Environment.NewLine
                );
            }

            InstallNativePatches();

            MelonLogger.Msg("[AutoSwitch] ╔══════════════════════════════════════════╗");
            MelonLogger.Msg("[AutoSwitch] ║   Auto Switch initialized            ║");
            MelonLogger.Msg("[AutoSwitch] ║   The packets learned line dancing ║");
            MelonLogger.Msg("[AutoSwitch] ╚══════════════════════════════════════════╝");
        }

        public override void OnSceneWasLoaded(int buildIndex, string sceneName)
        {
            if (buildIndex == 0)
                return;

            _nextScanTime = 0f;
            _lastSummary = string.Empty;
            AutoSwitchMod._lastFabricMembershipSignature = string.Empty;
            AutoSwitchMod._fabricChurnCooldownPending = false;

            RegisteredSwitches.Clear();
            RegisteredCables.Clear();
            FabricIdToSpeed.Clear();
            FabricIdToBundleCableIds.Clear();
            InjectedVirtualEdgeSignatures.Clear();
            VirtualEdgeSignatureToCableId.Clear();
            LoggedGraphTypeSignatures.Clear();
            LoggedFabricControllerSignatures.Clear();
            LoggedCableTupleTypes.Clear();
            LoggedSyntheticRegistrationSignatures.Clear();
            SyntheticRegisteredCableIds.Clear();
            NextVirtualCableId = -2000000;
            FabricIdToMemberIdsSnapshot.Clear();
            SwitchIdToFabricIdSnapshot.Clear();

            LoggedBundleSignatures.Clear();
            LoggedRemoteResolutionSignatures.Clear();
            LoggedDomainLinkSignatures.Clear();
            LoggedFanoutCandidateSignatures.Clear();
            LoggedPropagationSignatures.Clear();
            LoggedFabricPlanSignatures.Clear();

            SaveDataAutoLACP.ResetForScene();
            SaveDataAutoLACP.StartSafeBootstrap();
            SaveDataAutoLACP.QueueSceneWakePulse();

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

            FabricIdToMemberIdsSnapshot.Clear();
            foreach (KeyValuePair<string, HashSet<string>> kvp in fabricIdToMembers)
                FabricIdToMemberIdsSnapshot[kvp.Key] = new HashSet<string>(kvp.Value, StringComparer.OrdinalIgnoreCase);

            SwitchIdToFabricIdSnapshot.Clear();
            foreach (KeyValuePair<string, string> kvp in deviceIdToFabricId)
                SwitchIdToFabricIdSnapshot[kvp.Key] = kvp.Value;

            string fabricMembershipSignature = string.Join(
                "|",
                fabricIdToMembers
                    .OrderBy(kvp => kvp.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(kvp => kvp.Key + "=" + string.Join(",", kvp.Value.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))));

            if (!string.IsNullOrWhiteSpace(_lastFabricMembershipSignature) &&
                !string.Equals(_lastFabricMembershipSignature, fabricMembershipSignature, StringComparison.Ordinal))
            {
                if (!_fabricChurnCooldownPending)
                {
                    _fabricChurnCooldownPending = true;
                    _lastFabricMembershipSignature = fabricMembershipSignature;
                    LogToFile("FABRIC CHURN COOLDOWN | reason=membership-changed | fabrics=" + fabrics.Count.ToString(CultureInfo.InvariantCulture) + " | liveSwitches=" + liveSwitches.Count.ToString(CultureInfo.InvariantCulture));
                    SaveDataAutoLACP.QueueRegroup("fabric churn cooldown");
                    return;
                }

                _fabricChurnCooldownPending = false;
            }
            else
            {
                _fabricChurnCooldownPending = false;
            }

            _lastFabricMembershipSignature = fabricMembershipSignature;

            HashSet<string> allManagedIds = new HashSet<string>(
                deviceIdToFabricId.Keys,
                StringComparer.OrdinalIgnoreCase);

            NetworkSaveData networkSaveData = GetNetworkSaveData();
            int saveCableCount = CountSaveCables(networkSaveData);
            HashSet<int> externalLacpCableIds = GetExternalLacpCableIds(networkSaveData);

            FabricIdToSpeed.Clear();
            FabricIdToBundleCableIds.Clear();
            InjectedVirtualEdgeSignatures.Clear();
            VirtualEdgeSignatureToCableId.Clear();
            LoggedGraphTypeSignatures.Clear();
            LoggedFabricControllerSignatures.Clear();
            LoggedCableTupleTypes.Clear();
            LoggedSyntheticRegistrationSignatures.Clear();
            SyntheticRegisteredCableIds.Clear();
            NextVirtualCableId = -2000000;

            Dictionary<string, HashSet<string>> fabricAdjacency =
                BuildFabricAdjacency(deviceIdToFabricId, externalLacpCableIds);

            Dictionary<string, string> fabricToDomain =
                BuildRoutingDomains(fabricAdjacency, fabricIdToMembers.Keys.ToList());

            List<FabricRuntimePlan> fabricPlans = BuildFabricRuntimePlans(
                fabrics,
                fabricIdToMembers,
                deviceIdToFabricId,
                fabricToDomain,
                externalLacpCableIds);

            ApplyFabricTags(fabricPlans);
            ApplyVirtualFabricEdgesToNetworkMap(fabricPlans);
            RehydrateFabricPlansFromRegisteredCables(
                fabricPlans,
                deviceIdToFabricId,
                fabricToDomain,
                externalLacpCableIds);

            List<BundleBuilder> allBundles = new List<BundleBuilder>();

            int safeNativeBundleCount = 0;
            int fanoutCandidateCount = 0;
            int internalPassThroughCount = 0;
            int interFabricPassThroughCount = 0;
            int syntheticShareLinks = 0;
            int syntheticDomainFeeds = 0;

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
                        interFabricPassThroughCableIds);

                List<int> fabricCableIds = new List<int>();

                foreach (RemoteEdgeAccumulator accumulator in accumulators.Values.OrderBy(a => a.RemoteDeviceId, StringComparer.OrdinalIgnoreCase))
                {
                    List<int> distinctIds = accumulator.AllCableIds.Distinct().OrderBy(x => x).ToList();
                    if (distinctIds.Count < 2)
                        continue;

                    if (accumulator.LocalBuckets.Count > 1)
                    {
                        fanoutCandidateCount++;
                        foreach (int cableId in accumulator.AllCableIds)
                            fanoutCandidateCableIds.Add(cableId);

                        LogFanoutCandidate(plan, accumulator);
                    }

                    if (accumulator.IsManagedRemote)
                    {
                        string managedSkipSignature =
                            plan.FabricId + "|" +
                            plan.DomainId + "|" +
                            accumulator.RemoteDeviceId + "|" +
                            string.Join(",", distinctIds.OrderBy(x => x));

                        if (LoggedBundleSignatures.Add("SKIP|" + managedSkipSignature))
                        {
                            LogToFile(
                                "SAFE NATIVE BUNDLE SKIPPED | fabricId=" + plan.FabricId +
                                " | domainId=" + plan.DomainId +
                                " | remote=" + accumulator.RemoteDeviceId +
                                " | reason=managed-remote-cross-fabric" +
                                " | localMembers=[" + string.Join(",", accumulator.LocalBuckets.Keys.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                                " | cableCount=" + distinctIds.Count.ToString(CultureInfo.InvariantCulture) +
                                " | cableIds=[" + string.Join(",", distinctIds.OrderBy(x => x)) + "]"
                            );
                        }

                        continue;
                    }

                    BundleBuilder bundle = new BundleBuilder();
                    bundle.FabricId = plan.FabricId;
                    bundle.DomainId = plan.DomainId;
                    bundle.FabricEstimatedSpeedGbps = plan.EstimatedCapacityGbps;
                    bundle.OwnerLocalDeviceId = ChooseBundleOwner(plan, accumulator);
                    bundle.RemoteDeviceId = accumulator.RemoteDeviceId;
                    bundle.IsIngressPreferredOwner =
                        string.Equals(bundle.OwnerLocalDeviceId, plan.IngressAnchorSwitchId, StringComparison.OrdinalIgnoreCase);
                    bundle.IsPredictedDownstreamExit = accumulator.IsDownstreamExit || accumulator.IsDomainExit;

                    foreach (string localMember in accumulator.LocalBuckets.Keys.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                        bundle.LocalMemberIds.Add(localMember);

                    foreach (int cableId in distinctIds)
                    {
                        bundle.CableIds.Add(cableId);
                        fabricCableIds.Add(cableId);
                    }

                    allBundles.Add(bundle);
                    safeNativeBundleCount++;

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
                            " | ownerIsAnchor=" + bundle.IsIngressPreferredOwner.ToString().ToLowerInvariant() +
                            " | localMembers=[" + string.Join(",", bundle.LocalMemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                            " | remote=" + bundle.RemoteDeviceId +
                            " | downstreamExit=" + bundle.IsPredictedDownstreamExit.ToString().ToLowerInvariant() +
                            " | cableCount=" + bundle.CableIds.Count.ToString(CultureInfo.InvariantCulture) +
                            " | cableIds=[" + string.Join(",", bundle.CableIds.OrderBy(x => x)) + "]"
                        );
                    }
                }

                FabricIdToBundleCableIds[plan.FabricId] = fabricCableIds.Distinct().OrderBy(x => x).ToList();

                syntheticShareLinks += plan.SyntheticInternalShareLinks.Count;
                syntheticDomainFeeds += plan.DomainPropagationIntents.Count(i =>
                    string.Equals(i.IntentKind, "anchor-to-domain-uplink", StringComparison.OrdinalIgnoreCase));
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
                " | syntheticShareLinks=" + syntheticShareLinks.ToString(CultureInfo.InvariantCulture) +
                " | syntheticDomainFeeds=" + syntheticDomainFeeds.ToString(CultureInfo.InvariantCulture) +
                " | internalPassThrough=" + internalPassThroughCount.ToString(CultureInfo.InvariantCulture) +
                " | interFabricPassThrough=" + interFabricPassThroughCount.ToString(CultureInfo.InvariantCulture) +
                " | bundleCount=" + allBundles.Count.ToString(CultureInfo.InvariantCulture) +
                " | managedIds=" + allManagedIds.Count.ToString(CultureInfo.InvariantCulture);

            if (!string.Equals(summary, _lastSummary, StringComparison.Ordinal))
            {
                _lastSummary = summary;
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
                        " | anchor=" + plan.IngressAnchorSwitchId +
                        " | customerBaseScore=" + plan.AnchorCustomerBaseScore.ToString("0.##", CultureInfo.InvariantCulture) +
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
                List<RegisteredSwitchInfo> members = fabrics[index];

                string domainId;
                if (!fabricToDomain.TryGetValue(fabricId, out domainId))
                    domainId = fabricId;

                FabricRuntimePlan plan = new FabricRuntimePlan();
                plan.FabricId = fabricId;
                plan.DomainId = domainId;
                plan.EstimatedCapacityGbps = members.Sum(x => x.EstimatedCapacityGbps);

                foreach (RegisteredSwitchInfo member in members)
                {
                    plan.Members.Add(member);
                    if (member != null && !string.IsNullOrWhiteSpace(member.DeviceName))
                        plan.MemberIds.Add(member.DeviceName);
                }

                foreach (RegisteredSwitchInfo member in members)
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

                ChooseFabricAnchor(plan);
                ChooseInterFabricUplink(plan);
                BuildSyntheticInternalSharePlan(plan);
                plans.Add(plan);
            }

            BuildDomainPropagationPlans(plans);
            return plans;
        }


        private static bool IsLowTierSwitchId(string deviceId)
        {
            if (string.IsNullOrWhiteSpace(deviceId))
                return false;

            return deviceId.StartsWith("Switch16CU", StringComparison.OrdinalIgnoreCase) ||
                   deviceId.StartsWith("Switch4xSFP", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCoreSwitchId(string deviceId)
        {
            if (string.IsNullOrWhiteSpace(deviceId))
                return false;

            return deviceId.StartsWith("Switch32xQSFP", StringComparison.OrdinalIgnoreCase) ||
                   deviceId.StartsWith("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase);
        }

        private static void ChooseFabricAnchor(FabricRuntimePlan plan)
        {
            SwitchTrafficProfile best = plan.SwitchProfiles.Values
                .Where(p => p != null && IsCoreSwitchId(p.DeviceId))
                .OrderByDescending(p => p.CustomerBaseScore)
                .ThenByDescending(p => p.IngressScore)
                .ThenByDescending(p => p.ExternalServerEdgeCount)
                .ThenByDescending(p => p.ExternalUnknownEdgeCount)
                .ThenByDescending(p => p.DomainExitEdgeCount)
                .ThenBy(p => p.DeviceId, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();

            if (best == null)
            {
                best = plan.SwitchProfiles.Values
                    .Where(p => p != null)
                    .OrderBy(p => IsLowTierSwitchId(p.DeviceId) ? 1 : 0)
                    .ThenByDescending(p => p.CustomerBaseScore)
                    .ThenByDescending(p => p.IngressScore)
                    .ThenBy(p => p.DeviceId, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault();
            }

            plan.IngressAnchorSwitchId = best != null ? best.DeviceId : string.Empty;
            plan.AnchorCustomerBaseScore = best != null ? best.CustomerBaseScore : 0f;

            foreach (SwitchTrafficProfile profile in plan.SwitchProfiles.Values)
                profile.IsAnchor = string.Equals(profile.DeviceId, plan.IngressAnchorSwitchId, StringComparison.OrdinalIgnoreCase);
        }

        private static void ChooseInterFabricUplink(FabricRuntimePlan plan)
        {
            SwitchTrafficProfile best = plan.SwitchProfiles.Values
                .OrderByDescending(p => p.DomainExitEdgeCount)
                .ThenByDescending(p => p.InterFabricSameDomainEdgeCount)
                .ThenByDescending(p => p.IngressScore)
                .ThenBy(p => p.DeviceId, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();

            plan.PreferredDomainUplinkSwitchId =
                best != null && (best.DomainExitEdgeCount > 0 || best.InterFabricSameDomainEdgeCount > 0)
                    ? best.DeviceId
                    : string.Empty;
        }

        private static void BuildSyntheticInternalSharePlan(FabricRuntimePlan plan)
        {
            plan.SyntheticInternalShareLinks.Clear();

            if (string.IsNullOrWhiteSpace(plan.IngressAnchorSwitchId))
                return;

            foreach (string memberId in plan.MemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                if (string.Equals(memberId, plan.IngressAnchorSwitchId, StringComparison.OrdinalIgnoreCase))
                    continue;

                SyntheticShareLink link = new SyntheticShareLink();
                link.FabricId = plan.FabricId;
                link.DomainId = plan.DomainId;
                link.FromSwitchId = plan.IngressAnchorSwitchId;
                link.ToSwitchId = memberId;
                link.Reason = "anchor-share";
                plan.SyntheticInternalShareLinks.Add(link);
            }

            if (!string.IsNullOrWhiteSpace(plan.PreferredDomainUplinkSwitchId) &&
                !string.Equals(plan.PreferredDomainUplinkSwitchId, plan.IngressAnchorSwitchId, StringComparison.OrdinalIgnoreCase))
            {
                SyntheticShareLink uplinkLink = new SyntheticShareLink();
                uplinkLink.FabricId = plan.FabricId;
                uplinkLink.DomainId = plan.DomainId;
                uplinkLink.FromSwitchId = plan.IngressAnchorSwitchId;
                uplinkLink.ToSwitchId = plan.PreferredDomainUplinkSwitchId;
                uplinkLink.Reason = "anchor-to-uplink";
                plan.SyntheticInternalShareLinks.Add(uplinkLink);
            }
        }

        private static void BuildDomainPropagationPlans(List<FabricRuntimePlan> plans)
        {
            Dictionary<string, List<FabricRuntimePlan>> grouped = plans
                .GroupBy(p => p.DomainId, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderBy(p => p.FabricId, StringComparer.OrdinalIgnoreCase).ToList(),
                    StringComparer.OrdinalIgnoreCase);

            foreach (KeyValuePair<string, List<FabricRuntimePlan>> kvp in grouped)
            {
                string domainId = kvp.Key;
                List<FabricRuntimePlan> domainPlans = kvp.Value;
                float domainCapacity = domainPlans.Sum(p => p.EstimatedCapacityGbps);

                foreach (FabricRuntimePlan plan in domainPlans)
                {
                    plan.DomainEstimatedCapacityGbps = domainCapacity;
                    plan.DomainFabricIds = domainPlans.Select(p => p.FabricId).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();

                    foreach (FabricRuntimePlan peer in domainPlans)
                    {
                        if (string.Equals(peer.FabricId, plan.FabricId, StringComparison.OrdinalIgnoreCase))
                            continue;

                        DomainPropagationIntent intent = new DomainPropagationIntent();
                        intent.DomainId = domainId;
                        intent.FromFabricId = plan.FabricId;
                        intent.ToFabricId = peer.FabricId;
                        intent.FromSwitchId = !string.IsNullOrWhiteSpace(plan.PreferredDomainUplinkSwitchId)
                            ? plan.PreferredDomainUplinkSwitchId
                            : plan.IngressAnchorSwitchId;
                        intent.ToSwitchId = peer.IngressAnchorSwitchId;
                        intent.EstimatedBandwidthGbps = Math.Min(plan.EstimatedCapacityGbps, peer.EstimatedCapacityGbps);
                        intent.IntentKind = "anchor-to-domain-uplink";
                        plan.DomainPropagationIntents.Add(intent);
                    }
                }
            }
        }

        private static bool IsRoutingOnlySyntheticEdge(RegisteredCableInfo cable)
        {
            if (cable == null)
                return false;

            string extraA = cable.ExtraA ?? string.Empty;
            string extraB = cable.ExtraB ?? string.Empty;
            return extraA.IndexOf("domain-propagation-", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   extraB.IndexOf("domain-propagation-", StringComparison.OrdinalIgnoreCase) >= 0;
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

                if (cable.CableId < 0 || SyntheticRegisteredCableIds.Contains(cable.CableId))
                    continue;

                if (externalLacpCableIds.Contains(cable.CableId))
                    continue;

                if (IsRoutingOnlySyntheticEdge(cable))
                    continue;

                string localDeviceId;
                string remoteDeviceId;
                if (!TryResolveRemoteDeviceForSpecificLocal(deviceId, cable, out localDeviceId, out remoteDeviceId))
                    continue;

                if (string.IsNullOrWhiteSpace(remoteDeviceId))
                    continue;

                profile.TouchedCableIds.Add(cable.CableId);

                string remoteFabricId;
                if (deviceIdToFabricId.TryGetValue(remoteDeviceId, out remoteFabricId))
                {
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
                        continue;
                    }

                    profile.DomainExitEdgeCount++;
                    continue;
                }

                if (remoteDeviceId.StartsWith("SERVERROOT:", StringComparison.OrdinalIgnoreCase) ||
                    remoteDeviceId.StartsWith("SERVERFAMILY:", StringComparison.OrdinalIgnoreCase))
                {
                    profile.ExternalServerEdgeCount++;
                }
                else
                {
                    profile.ExternalUnknownEdgeCount++;
                }
            }

            profile.IngressScore =
                (profile.ExternalServerEdgeCount * 10.0f) +
                (profile.ExternalUnknownEdgeCount * 6.0f) +
                (profile.DomainExitEdgeCount * 5.0f) +
                (profile.InternalFabricEdgeCount * 0.25f);

            profile.CustomerBaseScore =
                (profile.ExternalServerEdgeCount * 12.0f) +
                (profile.ExternalUnknownEdgeCount * 9.0f) +
                (profile.DomainExitEdgeCount * 2.0f);

            return profile;
        }

        private static string ChooseBundleOwner(FabricRuntimePlan plan, RemoteEdgeAccumulator accumulator)
        {
            if (plan == null)
                return string.Empty;

            LocalEdgeBucket biggest = null;
            if (accumulator != null)
            {
                biggest = accumulator.LocalBuckets.Values
                    .OrderByDescending(b => b.CableIds.Distinct().Count())
                    .ThenBy(b => b.OwnerLocalDeviceId, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault();

                if (biggest != null &&
                    !string.IsNullOrWhiteSpace(biggest.OwnerLocalDeviceId) &&
                    biggest.CableIds.Distinct().Count() >= 2)
                {
                    return biggest.OwnerLocalDeviceId;
                }

                if (accumulator.LocalBuckets.Count == 1 &&
                    biggest != null &&
                    !string.IsNullOrWhiteSpace(biggest.OwnerLocalDeviceId))
                {
                    return biggest.OwnerLocalDeviceId;
                }
            }

            if (accumulator != null && accumulator.IsManagedRemote && !accumulator.IsDomainExit)
            {
                if (!string.IsNullOrWhiteSpace(plan.PreferredDomainUplinkSwitchId) &&
                    plan.MemberIds.Contains(plan.PreferredDomainUplinkSwitchId))
                {
                    return plan.PreferredDomainUplinkSwitchId;
                }
            }

            if (!string.IsNullOrWhiteSpace(plan.IngressAnchorSwitchId) &&
                plan.MemberIds.Contains(plan.IngressAnchorSwitchId))
            {
                return plan.IngressAnchorSwitchId;
            }

            if (biggest != null && !string.IsNullOrWhiteSpace(biggest.OwnerLocalDeviceId))
                return biggest.OwnerLocalDeviceId;

            return plan.MemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).FirstOrDefault() ?? string.Empty;
        }

        private static Dictionary<string, RemoteEdgeAccumulator> BuildRemoteEdgeAccumulators(
            FabricRuntimePlan plan,
            Dictionary<string, string> deviceIdToFabricId,
            Dictionary<string, string> fabricToDomain,
            HashSet<int> externalLacpCableIds,
            ref int internalPassThroughCount,
            HashSet<int> internalPassThroughCableIds,
            ref int interFabricPassThroughCount,
            HashSet<int> interFabricPassThroughCableIds)
        {
            Dictionary<string, RemoteEdgeAccumulator> result =
                new Dictionary<string, RemoteEdgeAccumulator>(StringComparer.OrdinalIgnoreCase);

            foreach (RegisteredCableInfo cable in RegisteredCables.Values.OrderBy(x => x.CableId))
            {
                if (cable == null)
                    continue;

                if (cable.CableId < 0 || SyntheticRegisteredCableIds.Contains(cable.CableId))
                    continue;

                if (externalLacpCableIds.Contains(cable.CableId))
                    continue;

                if (IsRoutingOnlySyntheticEdge(cable))
                    continue;

                string localDeviceId;
                string remoteDeviceId;
                if (!TryResolveRemoteDeviceFromLiveCable(cable, plan.MemberIds, out localDeviceId, out remoteDeviceId))
                    continue;

                if (string.IsNullOrWhiteSpace(localDeviceId) || string.IsNullOrWhiteSpace(remoteDeviceId))
                    continue;

                bool remoteIsManagedSwitch = false;
                bool sameFabric = false;
                bool sameDomain = false;
                bool isDomainExit = false;

                string remoteFabricId;
                if (deviceIdToFabricId.TryGetValue(remoteDeviceId, out remoteFabricId))
                {
                    remoteIsManagedSwitch = true;
                    sameFabric = string.Equals(remoteFabricId, plan.FabricId, StringComparison.OrdinalIgnoreCase);

                    string remoteDomainId;
                    if (!fabricToDomain.TryGetValue(remoteFabricId, out remoteDomainId))
                        remoteDomainId = remoteFabricId;

                    sameDomain = string.Equals(remoteDomainId, plan.DomainId, StringComparison.OrdinalIgnoreCase);
                    isDomainExit = !sameDomain;
                }

                if (remoteIsManagedSwitch && sameFabric)
                {
                    internalPassThroughCount++;
                    internalPassThroughCableIds.Add(cable.CableId);
                    continue;
                }

                if (remoteIsManagedSwitch && sameDomain)
                {
                    interFabricPassThroughCount++;
                    interFabricPassThroughCableIds.Add(cable.CableId);
                    continue;
                }

                RemoteEdgeAccumulator accumulator;
                if (!result.TryGetValue(remoteDeviceId, out accumulator))
                {
                    accumulator = new RemoteEdgeAccumulator();
                    accumulator.FabricId = plan.FabricId;
                    accumulator.DomainId = plan.DomainId;
                    accumulator.RemoteDeviceId = remoteDeviceId;
                    accumulator.IsDomainExit = isDomainExit;
                    accumulator.IsDownstreamExit = !remoteIsManagedSwitch;
                    accumulator.IsManagedRemote = remoteIsManagedSwitch;
                    result[remoteDeviceId] = accumulator;
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
                        tag.IsAnchor = string.Equals(member.DeviceName, plan.IngressAnchorSwitchId, StringComparison.OrdinalIgnoreCase);
                        tag.AnchorSwitchId = plan.IngressAnchorSwitchId;
                        tag.UplinkSwitchId = plan.PreferredDomainUplinkSwitchId;
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

                if (cable.CableId < 0 || SyntheticRegisteredCableIds.Contains(cable.CableId))
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
            int domainIndex = 1;

            foreach (string fabricId in allFabricIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                string domainId = "DOMAIN-" + domainIndex.ToString("000", CultureInfo.InvariantCulture);
                domainIndex++;
                result[fabricId] = domainId;
            }

            foreach (KeyValuePair<string, HashSet<string>> kvp in fabricAdjacency.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
            {
                foreach (string neighbor in kvp.Value.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    LogToFile(
                        "DOMAIN LINK SKIPPED | fabricA=" + kvp.Key +
                        " | fabricB=" + neighbor +
                        " | reason=insufficient-physical-bridge");
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
                " | anchor=" + plan.IngressAnchorSwitchId +
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
                plan.IngressAnchorSwitchId + "|" +
                plan.PreferredDomainUplinkSwitchId + "|" +
                string.Join(",", plan.DomainFabricIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase));

            if (LoggedFabricPlanSignatures.Add(signature))
            {
                LogToFile(
                    "FABRIC SHARE PLAN | fabricId=" + plan.FabricId +
                    " | domainId=" + plan.DomainId +
                    " | anchor=" + plan.IngressAnchorSwitchId +
                    " | anchorCustomerBaseScore=" + plan.AnchorCustomerBaseScore.ToString("0.##", CultureInfo.InvariantCulture) +
                    " | uplink=" + plan.PreferredDomainUplinkSwitchId +
                    " | domainCapacityGbps=" + plan.DomainEstimatedCapacityGbps.ToString("0.##", CultureInfo.InvariantCulture) +
                    " | memberIds=[" + string.Join(",", plan.MemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                    " | domainFabrics=[" + string.Join(",", plan.DomainFabricIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]"
                );
            }

            foreach (SwitchTrafficProfile profile in plan.SwitchProfiles.Values
                .OrderByDescending(p => p.CustomerBaseScore)
                .ThenByDescending(p => p.IngressScore)
                .ThenBy(p => p.DeviceId, StringComparer.OrdinalIgnoreCase))
            {
                LogToFile(
                    "FABRIC MEMBER SCORE | fabricId=" + plan.FabricId +
                    " | device=" + profile.DeviceId +
                    " | isAnchor=" + profile.IsAnchor.ToString().ToLowerInvariant() +
                    " | customerBaseScore=" + profile.CustomerBaseScore.ToString("0.##", CultureInfo.InvariantCulture) +
                    " | ingressScore=" + profile.IngressScore.ToString("0.##", CultureInfo.InvariantCulture) +
                    " | serverEdges=" + profile.ExternalServerEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | unknownEdges=" + profile.ExternalUnknownEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | domainExitEdges=" + profile.DomainExitEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | interFabricSameDomain=" + profile.InterFabricSameDomainEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | internalFabric=" + profile.InternalFabricEdgeCount.ToString(CultureInfo.InvariantCulture) +
                    " | touchedCables=[" + string.Join(",", profile.TouchedCableIds.OrderBy(x => x)) + "]"
                );
            }

            foreach (SyntheticShareLink link in plan.SyntheticInternalShareLinks
                .OrderBy(l => l.FromSwitchId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(l => l.ToSwitchId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(l => l.Reason, StringComparer.OrdinalIgnoreCase))
            {
                string sig =
                    link.FabricId + "|" + link.DomainId + "|" + link.FromSwitchId + "|" + link.ToSwitchId + "|" + link.Reason;

                if (!LoggedPropagationSignatures.Add("S|" + sig))
                    continue;

                LogToFile(
                    "SYNTHETIC FABRIC SHARE | fabricId=" + link.FabricId +
                    " | domainId=" + link.DomainId +
                    " | from=" + link.FromSwitchId +
                    " | to=" + link.ToSwitchId +
                    " | reason=" + link.Reason
                );
            }

            foreach (DomainPropagationIntent intent in plan.DomainPropagationIntents
                .OrderBy(i => i.IntentKind, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.ToFabricId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.ToSwitchId, StringComparer.OrdinalIgnoreCase))
            {
                string sig =
                    intent.DomainId + "|" +
                    intent.FromFabricId + "|" +
                    intent.ToFabricId + "|" +
                    intent.IntentKind + "|" +
                    intent.FromSwitchId + "|" +
                    intent.ToSwitchId;

                if (!LoggedPropagationSignatures.Add("D|" + sig))
                    continue;

                LogToFile(
                    "DOMAIN PROPAGATION | domainId=" + intent.DomainId +
                    " | fromFabric=" + intent.FromFabricId +
                    " | toFabric=" + intent.ToFabricId +
                    " | kind=" + intent.IntentKind +
                    " | fromSwitch=" + intent.FromSwitchId +
                    " | toSwitch=" + intent.ToSwitchId +
                    " | estGbps=" + intent.EstimatedBandwidthGbps.ToString("0.##", CultureInfo.InvariantCulture)
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
                    PatchIfFound(networkMapType, "FindAllRoutes", nameof(NetworkMap_FindAllRoutes_Postfix));
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

        private static void NetworkMap_FindAllRoutes_Postfix(
            object __instance,
            string __0,
            string __1,
            ref Il2CppSystem.Collections.Generic.List<Il2CppSystem.Collections.Generic.List<string>> __result)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(__0) || string.IsNullOrWhiteSpace(__1))
                    return;

                Dictionary<string, HashSet<string>> graph = BuildAugmentedGraph(__instance);
                if (graph.Count == 0)
                    return;

                List<string> route = FindAugmentedShortestPath(graph, __0, __1);
                if (route == null || route.Count < 2)
                    return;

                if (__result == null)
                    __result = new Il2CppSystem.Collections.Generic.List<Il2CppSystem.Collections.Generic.List<string>>();

                string newSignature = string.Join(">", route);
                foreach (Il2CppSystem.Collections.Generic.List<string> existing in __result)
                {
                    if (existing == null)
                        continue;

                    List<string> existingManaged = new List<string>();
                    foreach (string hop in existing)
                        existingManaged.Add(hop);

                    if (string.Equals(string.Join(">", existingManaged), newSignature, StringComparison.OrdinalIgnoreCase))
                        return;
                }

                Il2CppSystem.Collections.Generic.List<string> il2cppRoute = new Il2CppSystem.Collections.Generic.List<string>();
                foreach (string hop in route)
                    il2cppRoute.Add(hop);

                __result.Add(il2cppRoute);

                LogToFile(
                    "VIRTUAL ROUTE | start=" + __0 +
                    " | target=" + __1 +
                    " | hops=" + route.Count.ToString(CultureInfo.InvariantCulture) +
                    " | path=[" + string.Join(" -> ", route) + "]"
                );
            }
            catch (Exception ex)
            {
                LogToFile("NetworkMap_FindAllRoutes_Postfix failed: " + ex);
            }
        }


        private static readonly HashSet<string> InjectedVirtualEdgeSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, int> VirtualEdgeSignatureToCableId =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedGraphTypeSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedFabricControllerSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedCableTupleTypes =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static int NextVirtualCableId = -2000000;

        private static void ApplyVirtualFabricEdgesToNetworkMap(List<FabricRuntimePlan> plans)
        {
            try
            {
                NetworkMap networkMap = NetworkMap.instance;
                if (networkMap == null || plans == null)
                    return;

                LogGraphRuntimeTypes(networkMap);

                foreach (FabricRuntimePlan plan in plans.OrderBy(p => p.FabricId, StringComparer.OrdinalIgnoreCase))
                {
                    RegisterFabricController(plan);
                    List<Tuple<string, string, string>> desiredEdges = BuildFabricSpineEdges(plan);

                    foreach (Tuple<string, string, string> edge in desiredEdges)
                        EnsureVirtualConnection(networkMap, edge.Item1, edge.Item2, plan.FabricId, plan.DomainId, edge.Item3);

                    foreach (SyntheticShareLink shareLink in plan.SyntheticInternalShareLinks)
                    {
                        EnsureVirtualConnection(
                            networkMap,
                            shareLink.FromSwitchId,
                            shareLink.ToSwitchId,
                            plan.FabricId,
                            plan.DomainId,
                            "synthetic-share-" + (shareLink.Reason ?? "anchor-share"));
                    }
                }

                foreach (FabricRuntimePlan plan in plans.OrderBy(p => p.FabricId, StringComparer.OrdinalIgnoreCase))
                {
                    foreach (DomainPropagationIntent intent in plan.DomainPropagationIntents)
                    {
                        string fromSwitch = intent.FromSwitchId;
                        string toSwitch = intent.ToSwitchId;

                        if (string.IsNullOrWhiteSpace(fromSwitch))
                            fromSwitch = !string.IsNullOrWhiteSpace(plan.PreferredDomainUplinkSwitchId)
                                ? plan.PreferredDomainUplinkSwitchId
                                : plan.IngressAnchorSwitchId;

                        FabricRuntimePlan targetPlan = plans.FirstOrDefault(p =>
                            string.Equals(p.FabricId, intent.ToFabricId, StringComparison.OrdinalIgnoreCase));

                        if (string.IsNullOrWhiteSpace(toSwitch) && targetPlan != null)
                            toSwitch = !string.IsNullOrWhiteSpace(targetPlan.IngressAnchorSwitchId)
                                ? targetPlan.IngressAnchorSwitchId
                                : targetPlan.PreferredDomainUplinkSwitchId;

                        EnsureVirtualConnection(
                            networkMap,
                            fromSwitch,
                            toSwitch,
                            plan.FabricId,
                            plan.DomainId,
                            "domain-propagation-" + (intent.IntentKind ?? "inter-fabric"));
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile("ApplyVirtualFabricEdgesToNetworkMap failed: " + ex);
            }
        }

        private static List<Tuple<string, string, string>> BuildFabricSpineEdges(FabricRuntimePlan plan)
        {
            List<Tuple<string, string, string>> result = new List<Tuple<string, string, string>>();
            if (plan == null || plan.MemberIds.Count == 0)
                return result;

            string anchor = plan.IngressAnchorSwitchId ?? string.Empty;
            string uplink = plan.PreferredDomainUplinkSwitchId ?? string.Empty;

            List<string> members = plan.MemberIds
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddEdge(string a, string b, string reason)
            {
                if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b))
                    return;

                if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase))
                    return;

                string key = string.Compare(a, b, StringComparison.OrdinalIgnoreCase) < 0
                    ? a + "|" + b
                    : b + "|" + a;

                if (seen.Add(key))
                    result.Add(Tuple.Create(a, b, reason));
            }

            List<SwitchTrafficProfile> profiles = plan.SwitchProfiles.Values
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.DeviceId))
                .Where(p => p.DeviceId.StartsWith("Switch32xQSFP", StringComparison.OrdinalIgnoreCase) ||
                            p.DeviceId.StartsWith("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(p => p.ExternalUnknownEdgeCount + p.ExternalServerEdgeCount)
                .ThenByDescending(p => p.IngressScore)
                .ThenByDescending(p => p.InterFabricSameDomainEdgeCount)
                .ThenBy(p => p.DeviceId, StringComparer.OrdinalIgnoreCase)
                .ToList();

            List<string> hotLeaves = profiles
                .Where(p => !string.Equals(p.DeviceId, anchor, StringComparison.OrdinalIgnoreCase))
                .Where(p => p.ExternalUnknownEdgeCount > 0 || p.ExternalServerEdgeCount > 0)
                .Select(p => p.DeviceId)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(4)
                .ToList();

            // Anchor to uplink is the main trunk.
            if (!string.IsNullOrWhiteSpace(anchor) &&
                !string.IsNullOrWhiteSpace(uplink) &&
                !string.Equals(anchor, uplink, StringComparison.OrdinalIgnoreCase))
            {
                AddEdge(anchor, uplink, "anchor-uplink");
            }

            // Anchor to hot leaves with real external edges.
            foreach (string member in hotLeaves)
                AddEdge(anchor, member, "anchor-hotleaf");

            // Uplink to hot leaves helps handoff into the neighboring fabric.
            if (!string.IsNullOrWhiteSpace(uplink))
            {
                foreach (string member in hotLeaves)
                {
                    if (!string.Equals(member, uplink, StringComparison.OrdinalIgnoreCase))
                        AddEdge(uplink, member, "uplink-hotleaf");
                }
            }

            // Add one nearest-neighbor chain in sorted order to mimic a physical stack spine.
            List<string> qsfpBackbone = members
                .Where(id =>
                    id.StartsWith("Switch32xQSFP", StringComparison.OrdinalIgnoreCase) ||
                    id.StartsWith("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase))
                .OrderBy(id => id, StringComparer.OrdinalIgnoreCase)
                .ToList();

            for (int i = 0; i < qsfpBackbone.Count - 1; i++)
                AddEdge(qsfpBackbone[i], qsfpBackbone[i + 1], "stack-neighbor");

            // Anything still floating without a usable internal path gets a direct anchor feed.
            // This is especially important for 16CU / 4xSFP members that sit inside the fabric stack
            // but would otherwise never get any synthetic cable touches or ingress weight.
            HashSet<string> alreadyConnected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Tuple<string, string, string> edge in result)
            {
                alreadyConnected.Add(edge.Item1);
                alreadyConnected.Add(edge.Item2);
            }

            List<string> dormantMembers = members
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Where(id => !string.Equals(id, anchor, StringComparison.OrdinalIgnoreCase))
                .Where(id => !alreadyConnected.Contains(id))
                .OrderBy(id => id, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (string dormant in dormantMembers)
                AddEdge(anchor, dormant, "anchor-dormant");

            // Also give low-visibility members a direct feed even if they only touch one weak synthetic link.
            if (IsCoreSwitchId(anchor))
            {
                foreach (SwitchTrafficProfile profile in plan.SwitchProfiles.Values
                    .Where(p => p != null && !string.IsNullOrWhiteSpace(p.DeviceId))
                    .Where(p => !string.Equals(p.DeviceId, anchor, StringComparison.OrdinalIgnoreCase))
                    .Where(p => !IsLowTierSwitchId(p.DeviceId))
                    .Where(p => p.InternalFabricEdgeCount <= 1)
                    .Where(p => p.ExternalUnknownEdgeCount == 0)
                    .Where(p => p.ExternalServerEdgeCount == 0)
                    .OrderBy(p => p.DeviceId, StringComparer.OrdinalIgnoreCase))
                {
                    AddEdge(anchor, profile.DeviceId, "anchor-share-weak");
                }
            }

            // If the anchor somehow had no edges yet, give it a minimal lifeline to the first two members.
            if (!string.IsNullOrWhiteSpace(anchor) && result.Count == 0)
            {
                foreach (string member in members.Where(x => !string.Equals(x, anchor, StringComparison.OrdinalIgnoreCase)).Take(2))
                    AddEdge(anchor, member, "anchor-minimal");
            }

            return result;
        }

        private static void RegisterFabricController(FabricRuntimePlan plan)
        {
            if (plan == null)
                return;

            List<string> hotMembers = plan.SwitchProfiles.Values
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.DeviceId))
                .Where(p => p.DeviceId.StartsWith("Switch32xQSFP", StringComparison.OrdinalIgnoreCase) ||
                            p.DeviceId.StartsWith("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(p => p.ExternalUnknownEdgeCount + p.ExternalServerEdgeCount)
                .ThenByDescending(p => p.IngressScore)
                .ThenBy(p => p.DeviceId, StringComparer.OrdinalIgnoreCase)
                .Select(p => p.DeviceId)
                .Take(4)
                .ToList();

            string signature =
                plan.FabricId + "|" +
                plan.DomainId + "|" +
                plan.IngressAnchorSwitchId + "|" +
                plan.PreferredDomainUplinkSwitchId + "|" +
                string.Join(",", hotMembers);

            if (LoggedFabricControllerSignatures.Add(signature))
            {
                LogToFile(
                    "FABRIC CONTROLLER | fabricId=" + plan.FabricId +
                    " | domainId=" + plan.DomainId +
                    " | anchor=" + plan.IngressAnchorSwitchId +
                    " | uplink=" + plan.PreferredDomainUplinkSwitchId +
                    " | hotMembers=[" + string.Join(",", hotMembers) + "]" +
                    " | memberCount=" + plan.MemberIds.Count.ToString(CultureInfo.InvariantCulture)
                );
            }
        }


        private static void RegisterSyntheticCableInfo(
            int virtualCableId,
            string a,
            string b,
            string fabricId,
            string domainId,
            string reason)
        {
            try
            {
                RegisteredCableInfo info = new RegisteredCableInfo
                {
                    CableId = virtualCableId,
                    DeviceA = a ?? string.Empty,
                    DeviceB = b ?? string.Empty,
                    ExtraA = reason ?? string.Empty,
                    ExtraB = string.Empty,
                    PortA = -1,
                    PortB = -1,
                    EndpointObjectA = null,
                    EndpointObjectB = null,
                    ServerRootKeyA = string.Empty,
                    ServerRootKeyB = string.Empty,
                    EndpointSwitchIdA = a ?? string.Empty,
                    EndpointSwitchIdB = b ?? string.Empty
                };

                RegisteredCables[virtualCableId] = info;
                SyntheticRegisteredCableIds.Add(virtualCableId);

                string signature =
                    virtualCableId.ToString(CultureInfo.InvariantCulture) + "|" +
                    info.DeviceA + "|" + info.DeviceB + "|" +
                    (fabricId ?? string.Empty) + "|" + (domainId ?? string.Empty);

                if (LoggedSyntheticRegistrationSignatures.Add(signature))
                {
                    LogToFile(
                        "SYNTHETIC CABLE REGISTERED | fabricId=" + (fabricId ?? string.Empty) +
                        " | domainId=" + (domainId ?? string.Empty) +
                        " | cableId=" + virtualCableId.ToString(CultureInfo.InvariantCulture) +
                        " | a=" + info.DeviceA +
                        " | b=" + info.DeviceB +
                        " | reason=" + (reason ?? string.Empty));
                }
            }
            catch (Exception ex)
            {
                LogToFile("SYNTHETIC CABLE REGISTER FAILED | cableId=" + virtualCableId.ToString(CultureInfo.InvariantCulture) + " | ex=" + ex.Message);
            }
        }

        private static void RehydrateFabricPlansFromRegisteredCables(
            List<FabricRuntimePlan> existingPlans,
            Dictionary<string, string> deviceIdToFabricId,
            Dictionary<string, string> fabricToDomain,
            HashSet<int> externalLacpCableIds)
        {
            if (existingPlans == null)
                return;

            foreach (FabricRuntimePlan plan in existingPlans)
            {
                if (plan == null)
                    continue;

                plan.SwitchProfiles.Clear();

                foreach (RegisteredSwitchInfo member in plan.Members)
                {
                    if (member == null || string.IsNullOrWhiteSpace(member.DeviceName))
                        continue;

                    SwitchTrafficProfile profile = BuildSwitchTrafficProfile(
                        member.DeviceName,
                        plan.FabricId,
                        deviceIdToFabricId,
                        fabricToDomain,
                        externalLacpCableIds);

                    plan.SwitchProfiles[member.DeviceName] = profile;

                    LogToFile(
                        "REHYDRATED PROFILE | fabricId=" + plan.FabricId +
                        " | domainId=" + plan.DomainId +
                        " | device=" + profile.DeviceId +
                        " | touchedCables=[" + string.Join(",", profile.TouchedCableIds.OrderBy(x => x)) + "]" +
                        " | internalFabric=" + profile.InternalFabricEdgeCount.ToString(CultureInfo.InvariantCulture) +
                        " | interFabricSameDomain=" + profile.InterFabricSameDomainEdgeCount.ToString(CultureInfo.InvariantCulture) +
                        " | unknownEdges=" + profile.ExternalUnknownEdgeCount.ToString(CultureInfo.InvariantCulture) +
                        " | serverEdges=" + profile.ExternalServerEdgeCount.ToString(CultureInfo.InvariantCulture));
                }

                ChooseFabricAnchor(plan);
                ChooseInterFabricUplink(plan);
                BuildSyntheticInternalSharePlan(plan);
            }

            BuildDomainPropagationPlans(existingPlans);
        }

        private static void LogGraphRuntimeTypes(NetworkMap networkMap)
        {
            try
            {
                object rawAdjacency = TryGetMemberValue(networkMap, "adjacencyList");
                object rawSwitchConnections = TryGetMemberValue(networkMap, "switchConnections");
                object rawCableConnections = TryGetMemberValue(networkMap, "cableConnections");

                string adjacencyType = rawAdjacency != null ? rawAdjacency.GetType().FullName : "null";
                string switchConnectionsType = rawSwitchConnections != null ? rawSwitchConnections.GetType().FullName : "null";
                string cableConnectionsType = rawCableConnections != null ? rawCableConnections.GetType().FullName : "null";

                string signature = adjacencyType + " | " + switchConnectionsType + " | " + cableConnectionsType;
                if (LoggedGraphTypeSignatures.Add(signature))
                {
                    LogToFile(
                        "GRAPH TYPE | adjacencyList=" + adjacencyType +
                        " | switchConnections=" + switchConnectionsType +
                        " | cableConnections=" + cableConnectionsType
                    );
                }
            }
            catch (Exception ex)
            {
                LogToFile("LogGraphRuntimeTypes failed: " + ex.Message);
            }
        }

        private static void EnsureVirtualConnection(
            NetworkMap networkMap,
            string a,
            string b,
            string fabricId,
            string domainId,
            string reason)
        {
            if (networkMap == null || string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b))
                return;

            string signature = string.Compare(a, b, StringComparison.OrdinalIgnoreCase) < 0
                ? a + "|" + b
                : b + "|" + a;

            if (InjectedVirtualEdgeSignatures.Contains(signature))
                return;

            bool hadAdjacencyBefore = HasDirectAdjacency(networkMap, a, b);
            if (hadAdjacencyBefore)
            {
                InjectedVirtualEdgeSignatures.Add(signature);
                return;
            }

            bool connectedViaMethod = false;
            bool adjacencyForced = false;
            bool switchConnectionsForced = false;
            bool deviceConnectionsForced = false;
            bool cableMapForced = false;
            int virtualCableId = 0;
            bool routingOnly = !string.IsNullOrWhiteSpace(reason) &&
                reason.StartsWith("domain-propagation-", StringComparison.OrdinalIgnoreCase);

            try
            {
                MethodInfo connectMethod = networkMap.GetType().GetMethod(
                    "Connect",
                    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                    null,
                    new[] { typeof(string), typeof(string) },
                    null);

                if (connectMethod != null)
                {
                    connectMethod.Invoke(networkMap, new object[] { a, b });
                    connectedViaMethod = true;
                }

                adjacencyForced = ForceAdjacencyLink(networkMap, a, b) || hadAdjacencyBefore;

                if (!routingOnly)
                {
                    switchConnectionsForced = ForceSwitchConnectionsLink(networkMap, a, b);
                    deviceConnectionsForced = ForceDeviceConnectionsLink(networkMap, a, b);
                    cableMapForced = ForceCableConnectionMap(networkMap, a, b, signature, out virtualCableId);
                    if (cableMapForced && virtualCableId != 0)
                        RegisterSyntheticCableInfo(virtualCableId, a, b, fabricId, domainId, reason);
                }

                InjectedVirtualEdgeSignatures.Add(signature);

                LogToFile(
                    "VIRTUAL EDGE | fabricId=" + fabricId +
                    " | domainId=" + domainId +
                    " | a=" + a +
                    " | b=" + b +
                    " | reason=" + reason +
                    " | connectMethod=" + connectedViaMethod.ToString().ToLowerInvariant() +
                    " | adjacency=" + adjacencyForced.ToString().ToLowerInvariant() +
                    " | switchConnections=" + switchConnectionsForced.ToString().ToLowerInvariant() +
                    " | deviceConnections=" + deviceConnectionsForced.ToString().ToLowerInvariant() +
                    " | cableMap=" + cableMapForced.ToString().ToLowerInvariant() +
                    " | virtualCableId=" + virtualCableId.ToString(CultureInfo.InvariantCulture) +
                    " | cableBacked=" + cableMapForced.ToString().ToLowerInvariant()
                );
            }
            catch (Exception ex)
            {
                LogToFile(
                    "VIRTUAL EDGE FAIL | fabricId=" + fabricId +
                    " | domainId=" + domainId +
                    " | a=" + a +
                    " | b=" + b +
                    " | reason=" + reason +
                    " | " + ex.Message
                );
            }
        }

        private static bool ForceAdjacencyLink(NetworkMap networkMap, string a, string b)
        {
            try
            {
                object rawAdjacency = TryGetMemberValue(networkMap, "adjacencyList");
                if (rawAdjacency == null)
                    return false;

                return ForceDictionaryCollectionLink(rawAdjacency, a, b, "adjacency");
            }
            catch
            {
                return false;
            }
        }


        private static bool ForceSwitchConnectionsLink(NetworkMap networkMap, string a, string b)
        {
            try
            {
                object rawSwitchConnections = TryGetMemberValue(networkMap, "switchConnections");
                if (rawSwitchConnections == null)
                    return false;

                return ForceDictionaryCollectionLink(rawSwitchConnections, a, b, "switchConnections");
            }
            catch
            {
                return false;
            }
        }


        private static bool ForceDeviceConnectionsLink(NetworkMap networkMap, string a, string b)
        {
            try
            {
                object rawDevices = TryGetMemberValue(networkMap, "devices");
                if (rawDevices == null)
                    return false;

                Type dictType = rawDevices.GetType();
                MethodInfo tryGetValueMethod = dictType.GetMethod("TryGetValue");
                if (tryGetValueMethod == null)
                    return false;

                object[] argsA = new object[] { a, null };
                bool hasA = (bool)tryGetValueMethod.Invoke(rawDevices, argsA);
                object deviceA = hasA ? argsA[1] : null;

                object[] argsB = new object[] { b, null };
                bool hasB = (bool)tryGetValueMethod.Invoke(rawDevices, argsB);
                object deviceB = hasB ? argsB[1] : null;

                if (deviceA == null || deviceB == null)
                    return false;

                bool changed = false;
                changed |= ForceSingleDeviceConnection(deviceA, deviceB);
                changed |= ForceSingleDeviceConnection(deviceB, deviceA);
                return changed;
            }
            catch
            {
                return false;
            }
        }

        private static bool ForceSingleDeviceConnection(object sourceDevice, object targetDevice)
        {
            try
            {
                object rawConnections = TryGetMemberValue(sourceDevice, "Connections");
                if (rawConnections == null)
                    return false;

                MethodInfo addMethod = rawConnections.GetType().GetMethod("Add");
                MethodInfo containsMethod = rawConnections.GetType().GetMethod("Contains");
                if (addMethod == null)
                    return false;

                bool already = false;
                if (containsMethod != null)
                    already = (bool)containsMethod.Invoke(rawConnections, new object[] { targetDevice });

                if (!already)
                {
                    addMethod.Invoke(rawConnections, new object[] { targetDevice });
                    return true;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool ForceCableConnectionMap(NetworkMap networkMap, string a, string b, string signature, out int virtualCableId)
        {
            virtualCableId = 0;

            try
            {
                object rawCableConnections = TryGetMemberValue(networkMap, "cableConnections");
                if (rawCableConnections == null)
                    return false;

                int existingId;
                if (!VirtualEdgeSignatureToCableId.TryGetValue(signature, out existingId))
                {
                    existingId = NextVirtualCableId--;
                    VirtualEdgeSignatureToCableId[signature] = existingId;
                }

                virtualCableId = existingId;

                object tupleValue = CreateIl2CppCableTuple(rawCableConnections, a, b);
                if (tupleValue == null)
                    return false;

                bool setOk = TryDictionarySet(rawCableConnections, existingId, tupleValue);
                if (!setOk)
                    return false;

                object verify;
                if (TryDictionaryGet(rawCableConnections, existingId, out verify) && verify != null)
                    return true;

                return false;
            }
            catch
            {
                return false;
            }
        }


        private static bool ForceDictionaryCollectionLink(object dictionaryObject, string a, string b, string label)
        {
            bool changed = false;

            object collectionA = GetOrCreateDictionaryValueCollection(dictionaryObject, a);
            if (collectionA != null)
                changed |= TryCollectionAdd(collectionA, b);

            object collectionB = GetOrCreateDictionaryValueCollection(dictionaryObject, b);
            if (collectionB != null)
                changed |= TryCollectionAdd(collectionB, a);

            return changed || CollectionContainsValue(collectionA, b) || CollectionContainsValue(collectionB, a);
        }

        private static object GetOrCreateDictionaryValueCollection(object dictionaryObject, object key)
        {
            object existing;
            if (TryDictionaryGet(dictionaryObject, key, out existing) && existing != null)
                return existing;

            Type dictType = dictionaryObject.GetType();
            Type[] genericArgs = dictType.IsGenericType ? dictType.GetGenericArguments() : Type.EmptyTypes;
            if (genericArgs.Length < 2)
                return null;

            Type valueType = genericArgs[1];
            object created = null;

            try
            {
                created = Activator.CreateInstance(valueType);
            }
            catch
            {
                try
                {
                    MethodInfo ctor = valueType.GetMethod(".ctor", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (ctor != null)
                        created = ctor.Invoke(null, null);
                }
                catch { }
            }

            if (created == null)
                return null;

            if (TryDictionarySet(dictionaryObject, key, created))
                return created;

            object afterSet;
            if (TryDictionaryGet(dictionaryObject, key, out afterSet))
                return afterSet;

            return null;
        }

        private static bool TryDictionaryGet(object dictionaryObject, object key, out object value)
        {
            value = null;
            if (dictionaryObject == null)
                return false;

            try
            {
                MethodInfo tryGetValue = dictionaryObject.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(m => m.Name == "TryGetValue" && m.GetParameters().Length == 2);

                if (tryGetValue != null)
                {
                    object[] args = new object[] { key, null };
                    bool ok = (bool)tryGetValue.Invoke(dictionaryObject, args);
                    if (ok)
                        value = args[1];
                    return ok;
                }
            }
            catch { }

            try
            {
                PropertyInfo indexer = dictionaryObject.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(p => p.GetIndexParameters().Length == 1);
                if (indexer != null)
                {
                    value = indexer.GetValue(dictionaryObject, new object[] { key });
                    return value != null;
                }
            }
            catch { }

            return false;
        }

        private static bool TryDictionarySet(object dictionaryObject, object key, object value)
        {
            if (dictionaryObject == null)
                return false;

            try
            {
                PropertyInfo indexer = dictionaryObject.GetType().GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(p => p.GetIndexParameters().Length == 1 && p.CanWrite);
                if (indexer != null)
                {
                    indexer.SetValue(dictionaryObject, value, new object[] { key });
                    return true;
                }
            }
            catch { }

            try
            {
                MethodInfo addMethod = dictionaryObject.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(m => m.Name == "Add" && m.GetParameters().Length == 2);
                if (addMethod != null)
                {
                    addMethod.Invoke(dictionaryObject, new object[] { key, value });
                    return true;
                }
            }
            catch { }

            return false;
        }

        private static bool TryCollectionAdd(object collectionObject, object value)
        {
            if (collectionObject == null)
                return false;

            if (CollectionContainsValue(collectionObject, value))
                return false;

            try
            {
                MethodInfo addMethod = collectionObject.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(m => m.Name == "Add" && m.GetParameters().Length == 1);
                if (addMethod != null)
                {
                    addMethod.Invoke(collectionObject, new object[] { value });
                    return true;
                }
            }
            catch { }

            return false;
        }

        private static bool CollectionContainsValue(object collectionObject, object value)
        {
            if (collectionObject == null)
                return false;

            try
            {
                MethodInfo containsMethod = collectionObject.GetType().GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .FirstOrDefault(m => m.Name == "Contains" && m.GetParameters().Length == 1);
                if (containsMethod != null)
                    return (bool)containsMethod.Invoke(collectionObject, new object[] { value });
            }
            catch { }

            try
            {
                System.Collections.IEnumerable enumerable = collectionObject as System.Collections.IEnumerable;
                if (enumerable != null)
                {
                    foreach (object item in enumerable)
                    {
                        if (item != null && value != null && string.Equals(item.ToString(), value.ToString(), StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }
            catch { }

            return false;
        }


        private static object CreateIl2CppCableTuple(object rawCableConnections, string a, string b)
        {
            try
            {
                Type dictType = rawCableConnections.GetType();
                Type[] genericArgs = dictType.IsGenericType ? dictType.GetGenericArguments() : Type.EmptyTypes;
                if (genericArgs.Length < 2)
                    return null;

                Type valueType = genericArgs[1];
                string sig = valueType.FullName ?? valueType.Name;
                if (LoggedCableTupleTypes.Add(sig))
                    LogToFile("CABLE TUPLE TYPE | " + sig);

                try
                {
                    object tuple = Activator.CreateInstance(valueType, new object[] { a, b });
                    if (tuple != null)
                        return tuple;
                }
                catch { }

                try
                {
                    object tuple = Activator.CreateInstance(valueType);
                    if (tuple == null)
                        return null;

                    FieldInfo f1 = valueType.GetField("Item1", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    FieldInfo f2 = valueType.GetField("Item2", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

                    if (f1 != null) f1.SetValue(tuple, a);
                    if (f2 != null) f2.SetValue(tuple, b);

                    return tuple;
                }
                catch { }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static bool HasDirectAdjacency(NetworkMap networkMap, string a, string b)
        {
            try
            {
                object rawAdjacency = TryGetMemberValue(networkMap, "adjacencyList");
                if (rawAdjacency == null)
                    return false;

                object neighborsA;
                if (TryDictionaryGet(rawAdjacency, a, out neighborsA) && CollectionContainsValue(neighborsA, b))
                    return true;

                object neighborsB;
                if (TryDictionaryGet(rawAdjacency, b, out neighborsB) && CollectionContainsValue(neighborsB, a))
                    return true;
            }
            catch { }

            return false;
        }

        private static Dictionary<string, HashSet<string>> BuildAugmentedGraph(object networkMapInstance)
        {
            Dictionary<string, HashSet<string>> graph =
                new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            try
            {
                object rawAdjacency = TryGetMemberValue(networkMapInstance, "adjacencyList");
                Dictionary<string, List<string>> typedAdjacency = rawAdjacency as Dictionary<string, List<string>>;

                if (typedAdjacency != null)
                {
                    foreach (KeyValuePair<string, List<string>> kvp in typedAdjacency)
                    {
                        EnsureNode(graph, kvp.Key);

                        if (kvp.Value == null)
                            continue;

                        foreach (string neighbor in kvp.Value)
                            AddUndirectedEdge(graph, kvp.Key, neighbor);
                    }
                }

                object rawCableConnections = TryGetMemberValue(networkMapInstance, "cableConnections");
                if (rawCableConnections != null)
                {
                    PropertyInfo valuesProp = rawCableConnections.GetType().GetProperty("Values");
                    if (valuesProp != null)
                    {
                        IEnumerable values = valuesProp.GetValue(rawCableConnections, null) as IEnumerable;
                        if (values != null)
                        {
                            foreach (object entry in values)
                            {
                                if (entry == null)
                                    continue;

                                string startDevice = ReadTupleString(entry, "startDevice", "Item1");
                                string endDevice = ReadTupleString(entry, "endDevice", "Item2");

                                if (!string.IsNullOrWhiteSpace(startDevice) && !string.IsNullOrWhiteSpace(endDevice))
                                    AddUndirectedEdge(graph, startDevice, endDevice);
                            }
                        }
                    }
                }

                foreach (KeyValuePair<string, HashSet<string>> fabric in FabricIdToMemberIdsSnapshot)
                {
                    List<string> members = fabric.Value
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    for (int i = 0; i < members.Count; i++)
                    {
                        EnsureNode(graph, members[i]);

                        for (int j = i + 1; j < members.Count; j++)
                            AddUndirectedEdge(graph, members[i], members[j]);
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile("BuildAugmentedGraph failed: " + ex);
            }

            return graph;
        }

        private static List<string> FindAugmentedShortestPath(
            Dictionary<string, HashSet<string>> graph,
            string start,
            string target)
        {
            if (!graph.ContainsKey(start) || !graph.ContainsKey(target))
                return null;

            Queue<string> queue = new Queue<string>();
            HashSet<string> visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> previous = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            queue.Enqueue(start);
            visited.Add(start);

            while (queue.Count > 0)
            {
                string current = queue.Dequeue();
                if (string.Equals(current, target, StringComparison.OrdinalIgnoreCase))
                    break;

                HashSet<string> neighbors;
                if (!graph.TryGetValue(current, out neighbors) || neighbors == null)
                    continue;

                foreach (string neighbor in neighbors.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(neighbor))
                        continue;

                    if (!visited.Add(neighbor))
                        continue;

                    previous[neighbor] = current;
                    queue.Enqueue(neighbor);
                }
            }

            if (!visited.Contains(target))
                return null;

            List<string> path = new List<string>();
            string step = target;

            while (!string.IsNullOrWhiteSpace(step))
            {
                path.Add(step);

                string prior;
                if (!previous.TryGetValue(step, out prior))
                    break;

                step = prior;
            }

            path.Reverse();
            return path;
        }

        private static void EnsureNode(Dictionary<string, HashSet<string>> graph, string node)
        {
            if (string.IsNullOrWhiteSpace(node))
                return;

            if (!graph.ContainsKey(node))
                graph[node] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        private static void AddUndirectedEdge(Dictionary<string, HashSet<string>> graph, string a, string b)
        {
            if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b))
                return;

            EnsureNode(graph, a);
            EnsureNode(graph, b);
            graph[a].Add(b);
            graph[b].Add(a);
        }

        private static string ReadTupleString(object tupleLike, params string[] memberNames)
        {
            if (tupleLike == null || memberNames == null)
                return string.Empty;

            Type t = tupleLike.GetType();

            foreach (string name in memberNames)
            {
                try
                {
                    FieldInfo field = t.GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (field != null)
                    {
                        string value = field.GetValue(tupleLike) as string;
                        if (!string.IsNullOrWhiteSpace(value))
                            return value;
                    }

                    PropertyInfo prop = t.GetProperty(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                    if (prop != null && prop.CanRead)
                    {
                        string value = prop.GetValue(tupleLike, null) as string;
                        if (!string.IsNullOrWhiteSpace(value))
                            return value;
                    }
                }
                catch { }
            }

            return string.Empty;
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

        private static bool EndsWithRuntimeSuffix(string deviceId, string suffixDigits)
        {
            if (string.IsNullOrWhiteSpace(deviceId) || string.IsNullOrWhiteSpace(suffixDigits))
                return false;

            return deviceId.EndsWith("_-" + suffixDigits, StringComparison.OrdinalIgnoreCase) ||
                   deviceId.EndsWith("_" + suffixDigits, StringComparison.OrdinalIgnoreCase);
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

        internal static string NormalizeCloneName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return string.Empty;

            string result = name.Trim();
            if (result.EndsWith("(Clone)", StringComparison.OrdinalIgnoreCase))
                result = result.Substring(0, result.Length - "(Clone)".Length).Trim();

            return result;
        }

        internal static string NormalizeRuntimeIdentity(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            return TrailingRuntimeIdRegex.Replace(raw.Trim(), "");
        }

        internal static void LogToFile(string message)
        {
            if (!DebugLoggingEnabled)
                return;

            try
            {
                if (!ShouldWriteDebugLine(message))
                    return;

                Directory.CreateDirectory(DebugFolderPath);
                File.AppendAllText(
                    DebugLogPath,
                    "[" + DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) + "] " + message + Environment.NewLine
                );
            }
            catch { }
        }

        private static bool ShouldWriteDebugLine(string message)
        {
            if (!DebugLoggingEnabled || string.IsNullOrWhiteSpace(message))
                return false;

            string[] noisyPrefixes = new[]
            {
                "GRAPH TYPE |",
                "CABLE TUPLE TYPE |",
                "REMOTE RESOLVE |",
                "REHYDRATED PROFILE |",
                "SYNTHETIC CABLE REGISTERED |",
                "VIRTUAL EDGE |",
                "FABRIC CONTROLLER |",
                "FABRIC MEMBER SCORE |",
                "SYNTHETIC FABRIC SHARE |",
                "FABRIC SHARE PLAN |",
                "SCAN SUMMARY |",
                "ROUTING DOMAIN |",
                "FABRIC |",
                "FABRIC BUNDLE CABLE MAP |",
                "AUTO LACP | skipped",
                "AUTO LACP | regroup complete",
                "SWITCH WAKE |"
            };

            foreach (string prefix in noisyPrefixes)
            {
                if (message.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            return true;
        }
    }

    internal static class SaveDataAutoLACP
    {
        private static bool _bootstrapStarted;
        private static bool _regroupQueued;
        private static float _lastRunRealtime;
        private static string _lastAppliedSignature = string.Empty;
        private static bool _hasAppliedAtLeastOnce;
        private static bool _serverWakeQueued;
        private static float _lastServerWakeRealtime;
        private static readonly Dictionary<string, System.Reflection.MethodInfo> CachedWakeMethods =
            new Dictionary<string, System.Reflection.MethodInfo>(StringComparer.OrdinalIgnoreCase);
        private static string _lastFollowupRegroupSignature = string.Empty;

        private const float InitialDelaySeconds = 8.0f;
        private const float AdditionalSceneWakeDelaySeconds = 12.0f;
        private const float AdditionalSceneWakeDelaySecondsSecond = 20.0f;
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
            MelonCoroutines.Start(BootstrapRoutine());
        }

        internal static void ResetForScene()
        {
            _bootstrapStarted = false;
            _regroupQueued = false;
            _lastRunRealtime = 0f;
            _lastAppliedSignature = string.Empty;
            _hasAppliedAtLeastOnce = false;
            _serverWakeQueued = false;
            _lastServerWakeRealtime = 0f;
            CachedWakeMethods.Clear();
            _lastFollowupRegroupSignature = string.Empty;
            AutoSwitchMod._lastFabricMembershipSignature = string.Empty;
            AutoSwitchMod._fabricChurnCooldownPending = false;

            lock (_stateLock)
            {
                _desiredBundles = new List<BundleBuilder>();
                _desiredManagedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            _ownedGroupIds.Clear();
        }

        internal static void QueueSceneWakePulse()
        {
            QueueSwitchWakePulse("scene bootstrap");
            MelonCoroutines.Start(RunDelayedSceneWakePulse("scene delayed anchor", AdditionalSceneWakeDelaySeconds));
            MelonCoroutines.Start(RunDelayedSceneWakePulse("scene delayed random power", AdditionalSceneWakeDelaySecondsSecond));
        }

        private static IEnumerator RunDelayedSceneWakePulse(string reason, float delaySeconds)
        {
            if (delaySeconds > 0f)
                yield return new WaitForSeconds(delaySeconds);

            QueueSwitchWakePulse(reason);
        }

        private static void QueueFollowupRegroupIfNeeded()
        {
            string signature = _lastAppliedSignature ?? string.Empty;
            if (string.IsNullOrWhiteSpace(signature))
                return;

            if (string.Equals(_lastFollowupRegroupSignature, signature, StringComparison.Ordinal))
                return;

            _lastFollowupRegroupSignature = signature;
            MelonCoroutines.Start(RunFollowupRegroup(signature));
        }

        private static IEnumerator RunFollowupRegroup(string signature)
        {
            yield return new WaitForSeconds(2.25f);

            if (!string.Equals(_lastAppliedSignature ?? string.Empty, signature, StringComparison.Ordinal))
                yield break;

            QueueRegroup("followup confirm");
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
                QueueSwitchWakePulse("lacp regroup");
                QueueFollowupRegroupIfNeeded();
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
                        " | ownerIsAnchor=" + bundle.IsIngressPreferredOwner.ToString().ToLowerInvariant() +
                        " | localMembers=[" + string.Join(",", bundle.LocalMemberIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "]" +
                        " | pooledLocalCount=" + bundle.LocalMemberIds.Count.ToString(CultureInfo.InvariantCulture) +
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
                        " | pooledLocalCount=" + bundle.LocalMemberIds.Count.ToString(CultureInfo.InvariantCulture) +
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

        private static void QueueSwitchWakePulse(string reason)
        {
            if (_serverWakeQueued)
                return;

            _serverWakeQueued = true;
            MelonCoroutines.Start(RunSwitchWakePulse(reason));
        }

        private static string GetPreferredWakeSwitchId()
        {
            lock (_stateLock)
            {
                return _desiredBundles
                    .Select(b => b != null ? (b.OwnerLocalDeviceId ?? string.Empty) : string.Empty)
                    .FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? string.Empty;
            }
        }

        private static bool RootMatchesPreferredSwitchId(GameObject root, string preferredSwitchId)
        {
            if (root == null || string.IsNullOrWhiteSpace(preferredSwitchId))
                return false;

            string preferredRaw = preferredSwitchId.Trim();
            string preferredBase = AutoSwitchMod.NormalizeRuntimeIdentity(preferredRaw);

            bool Matches(string candidate)
            {
                if (string.IsNullOrWhiteSpace(candidate))
                    return false;

                string raw = candidate.Trim();
                string clone = AutoSwitchMod.NormalizeCloneName(raw);
                string baseId = AutoSwitchMod.NormalizeRuntimeIdentity(clone);

                return string.Equals(raw, preferredRaw, StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(clone, preferredRaw, StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(baseId, preferredRaw, StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(raw, preferredBase, StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(clone, preferredBase, StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(baseId, preferredBase, StringComparison.OrdinalIgnoreCase) ||
                       raw.IndexOf(preferredRaw, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       clone.IndexOf(preferredRaw, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       raw.IndexOf(preferredBase, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       clone.IndexOf(preferredBase, StringComparison.OrdinalIgnoreCase) >= 0;
            }

            if (Matches(root.name))
                return true;

            try
            {
                FabricGroupTag directTag = root.GetComponent<FabricGroupTag>();
                if (directTag != null)
                {
                    if (Matches(directTag.AnchorSwitchId))
                        return true;

                    if (directTag.IsAnchor && !string.IsNullOrWhiteSpace(preferredRaw))
                        return true;
                }
            }
            catch { }

            try
            {
                foreach (Component component in root.GetComponentsInChildren<Component>(true))
                {
                    if (component == null)
                        continue;

                    if (Matches(component.name))
                        return true;

                    try
                    {
                        FabricGroupTag tag = component as FabricGroupTag;
                        if (tag != null)
                        {
                            if (Matches(tag.AnchorSwitchId))
                                return true;

                            if (tag.IsAnchor && !string.IsNullOrWhiteSpace(preferredRaw))
                                return true;
                        }
                    }
                    catch { }

                    try
                    {
                        Type componentType = component.GetType();
                        if (componentType != null &&
                            (componentType.Name.IndexOf("Switch", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             componentType.FullName.IndexOf("Switch", StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            MethodInfo getSwitchId = componentType.GetMethod("GetSwitchId", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, Type.EmptyTypes, null);
                            if (getSwitchId != null)
                            {
                                object value = getSwitchId.Invoke(component, null);
                                if (Matches(Convert.ToString(value, CultureInfo.InvariantCulture)))
                                    return true;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return false;
        }

        private static IEnumerator RunSwitchWakePulse(string reason)
        {
            const int maxAttempts = 10;
            const float retryDelaySeconds = 0.75f;

            yield return new WaitForSeconds(0.60f);
            _serverWakeQueued = false;

            float now = Time.realtimeSinceStartup;
            if (now - _lastServerWakeRealtime < 1.0f)
                yield break;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                Dictionary<int, GameObject> switchRoots = new Dictionary<int, GameObject>();
                int touched = 0;
                int behaviourToggles = 0;
                int objectPulses = 0;
                int cablePulses = 0;
                int networkCalls = 0;
                int switchMethodCalls = 0;
                int powerToggleCalls = 0;

                try
                {
                    foreach (Component component in UnityEngine.Object.FindObjectsOfType<Component>())
                    {
                        if (component == null || component.gameObject == null)
                            continue;

                        Type type = component.GetType();
                        if (type == null)
                            continue;

                        string typeName = type.FullName ?? type.Name ?? string.Empty;
                        if (typeName.IndexOf("Switch", StringComparison.OrdinalIgnoreCase) < 0)
                            continue;

                        GameObject go = component.gameObject;
                        int id = go.GetInstanceID();
                        if (!switchRoots.ContainsKey(id))
                            switchRoots[id] = go;
                    }

                    if (switchRoots.Count > 0)
                    {
                        _lastServerWakeRealtime = Time.realtimeSinceStartup;
                        string preferredWakeSwitchId = GetPreferredWakeSwitchId();
                        bool anchorOnly = reason.IndexOf("delayed anchor", StringComparison.OrdinalIgnoreCase) >= 0;

                        IEnumerable<GameObject> orderedRoots = switchRoots.Values
                            .Where(r => r != null)
                            .OrderByDescending(r => RootMatchesPreferredSwitchId(r, preferredWakeSwitchId) ? 1 : 0)
                            .ThenBy(r => r.name ?? string.Empty, StringComparer.OrdinalIgnoreCase);

                        bool randomPowerOnly = reason.IndexOf("random power", StringComparison.OrdinalIgnoreCase) >= 0;

                        if (anchorOnly && !string.IsNullOrWhiteSpace(preferredWakeSwitchId))
                            orderedRoots = orderedRoots.Where(r => RootMatchesPreferredSwitchId(r, preferredWakeSwitchId)).Take(1).ToList();
                        else if (randomPowerOnly)
                        {
                            List<GameObject> eligibleRoots = orderedRoots
                                .Where(r => r != null)
                                .Where(r => !string.IsNullOrWhiteSpace(r.name))
                                .Where(r => r.name.IndexOf("Switch", StringComparison.OrdinalIgnoreCase) >= 0)
                                .ToList();

                            if (eligibleRoots.Count > 0)
                            {
                                int seed = Environment.TickCount ^ eligibleRoots.Count;
                                System.Random rng = new System.Random(seed);
                                GameObject chosen = eligibleRoots[rng.Next(eligibleRoots.Count)];
                                orderedRoots = new List<GameObject> { chosen };
                            }
                        }

                        foreach (GameObject root in orderedRoots)
                        {
                            if (root == null)
                                continue;

                            touched++;

                            try
                            {
                                if (root.activeSelf)
                                {
                                    root.SetActive(false);
                                    root.SetActive(true);
                                }
                                else
                                {
                                    root.SetActive(true);
                                }

                                objectPulses++;
                            }
                            catch { }

                            try
                            {
                                switchMethodCalls += InvokeSwitchWakeCandidates(root);
                            }
                            catch { }

                            try
                            {
                                bool randomPowerOnlyCurrent = reason.IndexOf("random power", StringComparison.OrdinalIgnoreCase) >= 0;
                                if (randomPowerOnlyCurrent)
                                    powerToggleCalls += TrySwitchPowerMicroToggle(root);
                                else
                                    powerToggleCalls += InvokeNonDestructiveSwitchOnCandidates(root);
                            }
                            catch { }
                        }

                        // Public-release safety: only the dedicated random-power experiment below uses
                        // a real switch power interaction, and only on one randomly selected switch.

                        foreach (Behaviour behaviour in UnityEngine.Object.FindObjectsOfType<Behaviour>(true))
                        {
                            if (behaviour == null)
                                continue;

                            Type behaviourType = behaviour.GetType();
                            if (behaviourType == null)
                                continue;

                            string behaviourTypeName = behaviourType.FullName ?? behaviourType.Name ?? string.Empty;
                            if (behaviourTypeName.IndexOf("Switch", StringComparison.OrdinalIgnoreCase) < 0 &&
                                behaviourTypeName.IndexOf("Power", StringComparison.OrdinalIgnoreCase) < 0)
                                continue;

                            try
                            {
                                bool original = behaviour.enabled;
                                behaviour.enabled = !original;
                                behaviour.enabled = original;
                                behaviourToggles++;
                            }
                            catch { }
                        }

                        AutoSwitchMod.LogToFile(
                            "SWITCH WAKE | reason=" + reason +
                            " | preferred=" + preferredWakeSwitchId +
                            " | anchorOnly=" + anchorOnly.ToString().ToLowerInvariant() +
                            " | touched=" + touched.ToString(CultureInfo.InvariantCulture) +
                            " | behaviourToggles=" + behaviourToggles.ToString(CultureInfo.InvariantCulture) +
                            " | objectPulses=" + objectPulses.ToString(CultureInfo.InvariantCulture) +
                            " | cablePulses=" + cablePulses.ToString(CultureInfo.InvariantCulture) +
                            " | networkCalls=" + networkCalls.ToString(CultureInfo.InvariantCulture) +
                            " | switchMethodCalls=" + switchMethodCalls.ToString(CultureInfo.InvariantCulture) +
                            " | powerToggleCalls=" + powerToggleCalls.ToString(CultureInfo.InvariantCulture) +
                            " | attempts=" + attempt.ToString(CultureInfo.InvariantCulture));
                        yield break;
                    }
                }
                catch (Exception ex)
                {
                    AutoSwitchMod.LogToFile(
                        "SWITCH WAKE | failed | reason=" + reason +
                        " | attempt=" + attempt.ToString(CultureInfo.InvariantCulture) +
                        " | " + ex.Message);
                    yield break;
                }

                if (attempt < maxAttempts)
                {
                    yield return new WaitForSeconds(retryDelaySeconds);
                    continue;
                }

                string preferredWakeSwitchIdFinal = GetPreferredWakeSwitchId();
                bool anchorOnlyFinal = reason.IndexOf("delayed anchor", StringComparison.OrdinalIgnoreCase) >= 0;
                AutoSwitchMod.LogToFile(
                    "SWITCH WAKE | reason=" + reason +
                    " | preferred=" + preferredWakeSwitchIdFinal +
                    " | anchorOnly=" + anchorOnlyFinal.ToString().ToLowerInvariant() +
                    " | touched=0 | behaviourToggles=0 | objectPulses=0 | cablePulses=0 | networkCalls=0 | switchMethodCalls=0 | powerToggleCalls=0 | attempts=" + attempt.ToString(CultureInfo.InvariantCulture));
                yield break;
            }
        }

        private static int InvokeNetworkWakeCandidates(NetworkMap networkMap)
        {
            return 0;
        }

        private static int TryInvokeZeroArgMethod(Component component, string methodName)
        {
            if (component == null || string.IsNullOrWhiteSpace(methodName))
                return 0;

            return TryInvokeCachedWakeMethod(component, methodName);
        }

        private static int InvokeNonDestructiveSwitchOnCandidates(GameObject root)
        {
            if (root == null)
                return 0;

            int invoked = 0;
            Component[] components;
            try
            {
                components = root.GetComponentsInChildren<Component>(true);
            }
            catch
            {
                return 0;
            }

            if (components == null)
                return 0;

            foreach (Component component in components)
            {
                if (component == null)
                    continue;

                Type type = component.GetType();
                if (type == null)
                    continue;

                string typeName = type.FullName ?? type.Name ?? string.Empty;
                if (typeName.IndexOf("NetworkSwitch", StringComparison.OrdinalIgnoreCase) < 0 &&
                    typeName.IndexOf("Switch", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                invoked += TryInvokeZeroArgMethod(component, "TurnOnCommonFunction");
                invoked += TryInvokeZeroArgMethod(component, "ReconnectCables");
                invoked += TryInvokeZeroArgMethod(component, "UpdateScreenUI");
                invoked += TryInvokeZeroArgMethod(component, "RefreshStatus");
                invoked += TryInvokeZeroArgMethod(component, "CheckPower");
            }

            return invoked;
        }

        private static GameObject ChooseSacrificialSwitchRoot(IEnumerable<GameObject> roots)
        {
            if (roots == null)
                return null;

            foreach (GameObject root in roots)
            {
                if (root == null)
                    continue;

                string name = root.name ?? string.Empty;
                if (name.IndexOf("Switch32xQSFP", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Switch", StringComparison.OrdinalIgnoreCase) >= 0)
                    return root;
            }

            return roots.FirstOrDefault(r => r != null);
        }

        private static readonly string[] SwitchPowerOffMethodNameHints =
        {
            "TurnOff",
            "PowerOff",
            "SetPowerOff",
            "DisablePower",
            "Shutdown",
            "SwitchOff"
        };

        private static readonly string[] SwitchPowerOnMethodNameHints =
        {
            "TurnOn",
            "PowerOn",
            "SetPowerOn",
            "EnablePower",
            "BootUp",
            "SwitchOn"
        };

        private static readonly string[] SwitchPowerToggleMethodNameHints =
        {
            "SetPower",
            "SetPowered",
            "SetIsPowered",
            "SetPowerState",
            "SetSwitchOn",
            "SetOn",
            "SetEnabled",
            "SetActive",
            "TogglePower"
        };

        private static readonly string[] SwitchPowerStateMemberNameHints =
        {
            "isPowered",
            "powered",
            "powerOn",
            "isOn",
            "switchOn",
            "turnedOn",
            "enabled",
            "active"
        };

        private static readonly string[] ButtonPowerMethodNameHints =
        {
            "Press",
            "Click",
            "Interact",
            "Use",
            "Activate",
            "Toggle",
            "Trigger",
            "Submit"
        };

        private static int TrySwitchPowerMicroToggle(GameObject root)
        {
            if (root == null)
                return 0;

            int invoked = 0;
            Component[] components;
            try
            {
                components = root.GetComponentsInChildren<Component>(true);
            }
            catch
            {
                return 0;
            }

            if (components == null)
                return 0;

            foreach (Component component in components)
            {
                if (component == null)
                    continue;

                Type type = component.GetType();
                if (type == null)
                    continue;

                string typeName = type.FullName ?? type.Name ?? string.Empty;
                if (typeName.IndexOf("NetworkSwitch", StringComparison.OrdinalIgnoreCase) < 0 &&
                    typeName.IndexOf("Switch", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                try
                {
                    MethodInfo powerButton = type.GetMethod(
                        "PowerButton",
                        BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                        null,
                        new[] { typeof(bool) },
                        null);

                    if (powerButton != null)
                    {
                        powerButton.Invoke(component, new object[] { false });
                        powerButton.Invoke(component, new object[] { false });
                        return 2;
                    }
                }
                catch { }

                try
                {
                    invoked += TryInvokeZeroArgMethod(component, "TurnOffCommonFunctions");
                    invoked += TryInvokeZeroArgMethod(component, "TurnOnCommonFunction");
                    if (invoked > 0)
                        return invoked;
                }
                catch { }
            }

            return invoked;
        }



        private static int TryInvokeButtonPowerInteraction(GameObject root)
        {
            if (root == null)
                return 0;

            int invoked = 0;
            Transform[] transforms;
            try
            {
                transforms = root.GetComponentsInChildren<Transform>(true);
            }
            catch
            {
                return 0;
            }

            if (transforms == null)
                return 0;

            foreach (Transform transform in transforms)
            {
                if (transform == null || transform.gameObject == null)
                    continue;

                string name = transform.gameObject.name ?? string.Empty;
                if (name.IndexOf("Buttonpower_", StringComparison.OrdinalIgnoreCase) < 0 &&
                    name.IndexOf("ButtonPower_", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                try
                {
                    var uiButton = transform.gameObject.GetComponent<UnityEngine.UI.Button>();
                    if (uiButton != null)
                    {
                        uiButton.onClick?.Invoke();
                        invoked++;
                        return invoked;
                    }
                }
                catch { }

                Component[] buttonComponents;
                try
                {
                    buttonComponents = transform.gameObject.GetComponents<Component>();
                }
                catch
                {
                    continue;
                }

                if (buttonComponents == null)
                    continue;

                foreach (Component component in buttonComponents)
                {
                    if (component == null)
                        continue;

                    foreach (string methodName in ButtonPowerMethodNameHints)
                    {
                        if (TryInvokeCachedWakeMethod(component, methodName) > 0)
                        {
                            invoked++;
                            return invoked;
                        }
                    }
                }
            }

            return invoked;
        }

        private static readonly Dictionary<string, MethodInfo> CachedWakeBooleanMethods = new Dictionary<string, MethodInfo>();

        private static int TryInvokeCachedWakeBooleanMethod(Component component, string methodName, bool state)
        {
            if (component == null || string.IsNullOrWhiteSpace(methodName))
                return 0;

            Type type = component.GetType();
            if (type == null)
                return 0;

            string cacheKey = type.FullName + "::" + methodName + "(bool)";
            if (!CachedWakeBooleanMethods.TryGetValue(cacheKey, out MethodInfo method))
            {
                method = type.GetMethod(methodName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, new[] { typeof(bool) }, null);
                if (method == null)
                {
                    method = type.GetMethod(methodName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, new[] { typeof(Boolean) }, null);
                }
                CachedWakeBooleanMethods[cacheKey] = method;
            }

            if (method == null)
                return 0;

            try
            {
                method.Invoke(component, new object[] { state });
                return 1;
            }
            catch
            {
                return 0;
            }
        }

        private static bool TryToggleBooleanLikeMember(Component component, string memberName)
        {
            if (component == null || string.IsNullOrWhiteSpace(memberName))
                return false;

            Type type = component.GetType();
            if (type == null)
                return false;

            const BindingFlags flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            try
            {
                PropertyInfo prop = type.GetProperty(memberName, flags);
                if (prop != null && prop.CanRead && prop.CanWrite)
                {
                    Type pt = prop.PropertyType;
                    if (pt == typeof(bool) || pt == typeof(Boolean))
                    {
                        bool original = (bool)prop.GetValue(component, null);
                        prop.SetValue(component, !original, null);
                        prop.SetValue(component, original, null);
                        return true;
                    }
                    if (pt == typeof(int) || pt == typeof(Int32))
                    {
                        int original = (int)prop.GetValue(component, null);
                        int alt = original == 0 ? 1 : 0;
                        prop.SetValue(component, alt, null);
                        prop.SetValue(component, original, null);
                        return true;
                    }
                }
            }
            catch { }

            try
            {
                FieldInfo field = type.GetField(memberName, flags);
                if (field != null)
                {
                    Type ft = field.FieldType;
                    if (ft == typeof(bool) || ft == typeof(Boolean))
                    {
                        bool original = (bool)field.GetValue(component);
                        field.SetValue(component, !original);
                        field.SetValue(component, original);
                        return true;
                    }
                    if (ft == typeof(int) || ft == typeof(Int32))
                    {
                        int original = (int)field.GetValue(component);
                        int alt = original == 0 ? 1 : 0;
                        field.SetValue(component, alt);
                        field.SetValue(component, original);
                        return true;
                    }
                }
            }
            catch { }

            return false;
        }
        private static readonly string[] SwitchWakeMethodNameHints =
        {
            "CheckConnections",
            "CheckConnection",
            "RefreshConnections",
            "RefreshConnection",
            "RefreshPorts",
            "RefreshPort",
            "UpdatePorts",
            "UpdatePort",
            "UpdateConnections",
            "UpdateConnection",
            "RebuildPorts",
            "RebuildPort",
            "RebuildConnections",
            "RebuildConnection",
            "DirtyPorts",
            "DirtyPort",
            "MarkDirty",
            "SetDirty",
            "RefreshPower",
            "CheckPower",
            "UpdatePower",
            "Recalculate",
            "RefreshStatus",
            "UpdateStatus"
        };

        private static int InvokeSwitchWakeCandidates(GameObject root)
        {
            if (root == null)
                return 0;

            int invoked = 0;
            Component[] components;
            try
            {
                components = root.GetComponentsInChildren<Component>(true);
            }
            catch
            {
                return 0;
            }

            if (components == null)
                return 0;

            foreach (Component component in components)
            {
                if (component == null)
                    continue;

                Type type = component.GetType();
                if (type == null)
                    continue;

                string typeName = type.FullName ?? type.Name ?? string.Empty;
                if (typeName.IndexOf("Switch", StringComparison.OrdinalIgnoreCase) < 0 &&
                    typeName.IndexOf("Port", StringComparison.OrdinalIgnoreCase) < 0 &&
                    typeName.IndexOf("Socket", StringComparison.OrdinalIgnoreCase) < 0 &&
                    typeName.IndexOf("Cable", StringComparison.OrdinalIgnoreCase) < 0 &&
                    typeName.IndexOf("Power", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                foreach (string methodName in SwitchWakeMethodNameHints)
                    invoked += TryInvokeCachedWakeMethod(component, methodName);
            }

            return invoked;
        }

        private static int TryInvokeCachedWakeMethod(object target, string methodName)
        {
            if (target == null || string.IsNullOrWhiteSpace(methodName))
                return 0;

            string cacheKey = target.GetType().FullName + "::" + methodName;
            System.Reflection.MethodInfo method;
            if (!CachedWakeMethods.TryGetValue(cacheKey, out method))
            {
                method = target.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                    .FirstOrDefault(m =>
                        string.Equals(m.Name, methodName, StringComparison.OrdinalIgnoreCase) &&
                        m.GetParameters().Length == 0 &&
                        m.ReturnType != typeof(IEnumerator));
                CachedWakeMethods[cacheKey] = method;
            }

            if (method == null)
                return 0;

            try
            {
                method.Invoke(target, null);
                return 1;
            }
            catch
            {
                return 0;
            }
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
        public bool IsAnchor { get; set; }
        public string AnchorSwitchId { get; set; }
        public string UplinkSwitchId { get; set; }
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
        public bool IsDomainExit;
        public bool IsDownstreamExit;
        public bool IsManagedRemote;

        public readonly Dictionary<string, LocalEdgeBucket> LocalBuckets =
            new Dictionary<string, LocalEdgeBucket>(StringComparer.OrdinalIgnoreCase);

        public readonly HashSet<int> AllCableIds = new HashSet<int>();

        public int TotalCableCount => AllCableIds.Count;
    }

    internal sealed class FabricRuntimePlan
    {
        public string FabricId;
        public string DomainId;
        public float EstimatedCapacityGbps;
        public float DomainEstimatedCapacityGbps;
        public string IngressAnchorSwitchId;
        public float AnchorCustomerBaseScore;
        public string PreferredDomainUplinkSwitchId;

        public readonly List<RegisteredSwitchInfo> Members = new List<RegisteredSwitchInfo>();
        public readonly HashSet<string> MemberIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, SwitchTrafficProfile> SwitchProfiles =
            new Dictionary<string, SwitchTrafficProfile>(StringComparer.OrdinalIgnoreCase);

        public List<string> DomainFabricIds = new List<string>();
        public readonly List<SyntheticShareLink> SyntheticInternalShareLinks = new List<SyntheticShareLink>();
        public readonly List<DomainPropagationIntent> DomainPropagationIntents = new List<DomainPropagationIntent>();
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
        public float CustomerBaseScore;
        public bool IsAnchor;
        public readonly HashSet<int> TouchedCableIds = new HashSet<int>();
    }

    internal sealed class SyntheticShareLink
    {
        public string FabricId;
        public string DomainId;
        public string FromSwitchId;
        public string ToSwitchId;
        public string Reason;
    }

    internal sealed class DomainPropagationIntent
    {
        public string DomainId;
        public string FromFabricId;
        public string ToFabricId;
        public string FromSwitchId;
        public string ToSwitchId;
        public float EstimatedBandwidthGbps;
        public string IntentKind;
    }
}
