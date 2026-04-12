using HarmonyLib;
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
using UnityEngine;
using Component = UnityEngine.Component;

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "0.1.9", "Big Texas Jerky")]
[assembly: MelonGame("Waseku", "Data Center")]

namespace AutoSwitch
{
    public sealed class AutoSwitchMod : MelonMod
    {
        private const float ScanIntervalSeconds = 8.0f;

        private const float SameFabricXZTolerance = 0.35f;
        private const float AdjacentYTolerance = 0.035f;
        private const float FallbackExpectedStep = 0.0444f;

        private const float OriginXZRejectRadius = 1.25f;
        private const float PositionBucketSize = 0.02f;
        private const int MaxObjectsPerPositionBucket = 6;

        private static readonly HashSet<string> AllowedSwitchNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "Switch32xQSFP",
            "Switch4xQSXP16xSFP",
            "Switch4xSFP",
            "Switch16CU"
        };

        private static string DebugFolderPath =>
            Path.Combine(MelonEnvironment.ModsDirectory, "AutoSwitch");

        private static string DebugLogPath =>
            Path.Combine(DebugFolderPath, "autoswitch-debug.log");

        private static readonly Dictionary<string, string> LastClusterSignatures = new(StringComparer.OrdinalIgnoreCase);

        // Fabric state
        private static readonly Dictionary<string, float> FabricIdToSpeed = new(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, HashSet<int>> FabricIdToCableIds = new(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<int, float> CableIdToFabricSpeed = new();
        private static readonly Dictionary<int, string> CableIdToFabricId = new();

        // Network registration state captured from NetworkMap.RegisterCableConnection
        private static readonly Dictionary<int, CableRegistrationInfo> RegisteredCables = new();
        private static readonly Dictionary<string, HashSet<int>> DeviceNameToCableIds = new(StringComparer.OrdinalIgnoreCase);

        // Device-name discovery on switch objects
        private static readonly Dictionary<int, HashSet<string>> SwitchInstanceIdToDeviceNames = new();
        private static readonly HashSet<string> LoggedDeviceNameMappings = new(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> LoggedPatches = new(StringComparer.OrdinalIgnoreCase);

        private float _nextScanTime;
        private string _lastSummary = string.Empty;
        private bool _loggedNearMissesThisScene;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();

            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(DebugLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] AutoSwitch debug log started.{Environment.NewLine}");

            InstallNativePatches();

            MelonLogger.Msg("[AutoSwitch] v0.1.9 active. NetworkMap cable registration path enabled.");
        }

        public override void OnSceneWasLoaded(int buildIndex, string sceneName)
        {
            if (buildIndex == 0)
                return;

            _nextScanTime = 0f;
            _lastSummary = string.Empty;
            _loggedNearMissesThisScene = false;

            LastClusterSignatures.Clear();
            FabricIdToSpeed.Clear();
            FabricIdToCableIds.Clear();
            CableIdToFabricSpeed.Clear();
            CableIdToFabricId.Clear();
            RegisteredCables.Clear();
            DeviceNameToCableIds.Clear();
            SwitchInstanceIdToDeviceNames.Clear();
            LoggedDeviceNameMappings.Clear();

            LogToFile($"Scene loaded: {sceneName} ({buildIndex})");
        }

        public override void OnUpdate()
        {
            try
            {
                if (Time.time < _nextScanTime)
                    return;

                _nextScanTime = Time.time + ScanIntervalSeconds;
                RunFabricScanAndApply();
            }
            catch (Exception ex)
            {
                MelonLogger.Error($"[AutoSwitch] OnUpdate exception: {ex}");
                LogToFile($"OnUpdate exception: {ex}");
                _nextScanTime = Time.time + 10.0f;
            }
        }

        private void RunFabricScanAndApply()
        {
            List<SwitchNode> rawCandidates = DiscoverPlacedLooseSwitchCandidates();
            List<SwitchNode> filtered = FilterJunkPileCandidates(rawCandidates);
            List<List<SwitchNode>> allFabrics = BuildWorldSpaceFabricGroups(filtered);

            List<List<SwitchNode>> activeFabrics = allFabrics
                .Where(f => f.Count >= 2)
                .OrderByDescending(f => f.Count)
                .ThenByDescending(f => f.Sum(x => x.EstimatedCapacityGbps))
                .ToList();

            int adjacencyPairs = 0;
            foreach (List<SwitchNode> fabric in activeFabrics)
            {
                List<SwitchNode> ordered = fabric
                    .OrderByDescending(s => s.WorldPosition.y)
                    .ThenBy(s => s.WorldPosition.x)
                    .ThenBy(s => s.WorldPosition.z)
                    .ToList();

                for (int i = 1; i < ordered.Count; i++)
                {
                    if (AreAdjacent(ordered[i - 1], ordered[i], out _))
                        adjacencyPairs++;
                }
            }

            ApplyFabricCombining(activeFabrics);
            RefreshSwitchDeviceNameMappings(activeFabrics);
            RebuildFabricCableMembership(activeFabrics);

            string summary =
                $"SCAN SUMMARY | rawCandidates={rawCandidates.Count} | filteredSwitches={filtered.Count} | activeFabrics={activeFabrics.Count} | adjacentPairs={adjacencyPairs} | registeredCables={RegisteredCables.Count} | trackedCableIds={CableIdToFabricSpeed.Count}";

            if (!string.Equals(summary, _lastSummary, StringComparison.Ordinal))
            {
                _lastSummary = summary;
                MelonLogger.Msg($"[AutoSwitch] {summary}");
                LogToFile(summary);

                int fabricIndex = 1;
                foreach (List<SwitchNode> fabric in activeFabrics)
                {
                    string fabricId = $"FABRIC-{fabricIndex:000}";
                    float totalEstimatedCapacity = fabric.Sum(f => f.EstimatedCapacityGbps);

                    string fabricLine =
                        $"FABRIC | id={fabricId} | members={fabric.Count} | estCapacityGbps={totalEstimatedCapacity.ToString("0.##", CultureInfo.InvariantCulture)} | trackedCableIds={GetTrackedCableCountForFabric(fabricId)} | switches={string.Join(", ", fabric.Select(s => $"{s.GameObject.name}@({s.WorldPosition.x:0.###},{s.WorldPosition.y:0.###},{s.WorldPosition.z:0.###})[ports={s.TotalPorts}]"))}";
                    LogToFile(fabricLine);
                    fabricIndex++;
                }
            }

            if (!_loggedNearMissesThisScene)
            {
                _loggedNearMissesThisScene = true;
                DumpNearMisses();
            }
        }

        private void InstallNativePatches()
        {
            try
            {
                Type networkMapType = AccessTools.TypeByName("Il2Cpp.NetworkMap");
                if (networkMapType != null)
                {
                    MethodInfo registerCable = AccessTools.Method(networkMapType, "RegisterCableConnection");
                    if (registerCable != null)
                    {
                        HarmonyInstance.Patch(
                            registerCable,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(NetworkMap_RegisterCableConnection_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap.RegisterCableConnection");
                    }

                    MethodInfo createLacp = AccessTools.Method(networkMapType, "CreateLACPGroup");
                    if (createLacp != null)
                    {
                        HarmonyInstance.Patch(
                            createLacp,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(NetworkMap_CreateLACPGroup_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap.CreateLACPGroup");
                    }

                    MethodInfo updateServerSpeed = AccessTools.Method(networkMapType, "UpdateCustomerServerCountAndSpeed");
                    if (updateServerSpeed != null)
                    {
                        HarmonyInstance.Patch(
                            updateServerSpeed,
                            prefix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(NetworkMap_UpdateCustomerServerCountAndSpeed_Prefix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap.UpdateCustomerServerCountAndSpeed");
                    }
                }

                Type lacpGroupType = AccessTools.TypeByName("Il2Cpp.NetworkMap+LACPGroup");
                if (lacpGroupType != null)
                {
                    MethodInfo aggSpeed = AccessTools.Method(lacpGroupType, "GetAggregatedSpeed");
                    if (aggSpeed != null)
                    {
                        HarmonyInstance.Patch(
                            aggSpeed,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(LACPGroup_GetAggregatedSpeed_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap+LACPGroup.GetAggregatedSpeed");
                    }
                }

                Type cableLinkType = AccessTools.TypeByName("Il2Cpp.CableLink");
                if (cableLinkType != null)
                {
                    MethodInfo setConnectionSpeed = AccessTools.Method(cableLinkType, "SetConnectionSpeed", new[] { typeof(float) });
                    if (setConnectionSpeed != null)
                    {
                        HarmonyInstance.Patch(
                            setConnectionSpeed,
                            prefix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(CableLink_SetConnectionSpeed_Prefix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched CableLink.SetConnectionSpeed");
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile($"InstallNativePatches failed: {ex}");
            }
        }

        private static void LogPatch(string message)
        {
            if (LoggedPatches.Add(message))
                LogToFile(message);
        }

        private void ApplyFabricCombining(List<List<SwitchNode>> activeFabrics)
        {
            int clusterIndex = 1;

            foreach (List<SwitchNode> fabric in activeFabrics)
            {
                string fabricId = $"FABRIC-{clusterIndex:000}";
                float totalEstimatedCapacity = fabric.Sum(x => x.EstimatedCapacityGbps);
                int totalPorts = fabric.Sum(x => x.TotalPorts);

                string signature = BuildFabricSignature(fabric, totalEstimatedCapacity, totalPorts);

                if (LastClusterSignatures.TryGetValue(fabricId, out string existing) &&
                    string.Equals(existing, signature, StringComparison.Ordinal))
                {
                    clusterIndex++;
                    continue;
                }

                LastClusterSignatures[fabricId] = signature;

                List<SwitchNode> ordered = fabric
                    .OrderByDescending(s => s.WorldPosition.y)
                    .ThenBy(s => s.WorldPosition.x)
                    .ThenBy(s => s.WorldPosition.z)
                    .ToList();

                for (int i = 0; i < ordered.Count; i++)
                {
                    SwitchNode node = ordered[i];
                    ApplyCombinedFabricToSwitch(node, fabricId, i + 1, ordered.Count, totalEstimatedCapacity, totalPorts);
                }

                LogToFile($"APPLY | fabricId={fabricId} | members={ordered.Count} | totalPorts={totalPorts} | estCapacityGbps={totalEstimatedCapacity.ToString("0.##", CultureInfo.InvariantCulture)}");
                clusterIndex++;
            }
        }

        private void RefreshSwitchDeviceNameMappings(List<List<SwitchNode>> activeFabrics)
        {
            SwitchInstanceIdToDeviceNames.Clear();

            foreach (List<SwitchNode> fabric in activeFabrics)
            {
                foreach (SwitchNode node in fabric)
                {
                    if (node?.GameObject == null)
                        continue;

                    int id = node.GameObject.GetInstanceID();
                    HashSet<string> names = DiscoverPossibleDeviceNames(node.GameObject);

                    if (names.Count == 0)
                        continue;

                    SwitchInstanceIdToDeviceNames[id] = names;

                    string joined = string.Join(", ", names.OrderBy(x => x));
                    string logKey = $"{id}:{joined}";
                    if (LoggedDeviceNameMappings.Add(logKey))
                        LogToFile($"DEVICE NAME MAP | switch={node.GameObject.name} | names=[{joined}]");
                }
            }
        }

        private void RebuildFabricCableMembership(List<List<SwitchNode>> activeFabrics)
        {
            FabricIdToSpeed.Clear();
            FabricIdToCableIds.Clear();
            CableIdToFabricSpeed.Clear();
            CableIdToFabricId.Clear();

            int clusterIndex = 1;

            foreach (List<SwitchNode> fabric in activeFabrics)
            {
                string fabricId = $"FABRIC-{clusterIndex:000}";
                float totalEstimatedCapacity = fabric.Sum(x => x.EstimatedCapacityGbps);

                FabricIdToSpeed[fabricId] = totalEstimatedCapacity;
                FabricIdToCableIds[fabricId] = new HashSet<int>();

                HashSet<string> allDeviceNames = new(StringComparer.OrdinalIgnoreCase);

                foreach (SwitchNode node in fabric)
                {
                    if (node?.GameObject == null)
                        continue;

                    int key = node.GameObject.GetInstanceID();
                    if (SwitchInstanceIdToDeviceNames.TryGetValue(key, out HashSet<string> names))
                    {
                        foreach (string n in names)
                            allDeviceNames.Add(n);
                    }
                }

                foreach ((int cableId, CableRegistrationInfo info) in RegisteredCables)
                {
                    bool aMatch = !string.IsNullOrWhiteSpace(info.DeviceA) && allDeviceNames.Contains(info.DeviceA);
                    bool bMatch = !string.IsNullOrWhiteSpace(info.DeviceB) && allDeviceNames.Contains(info.DeviceB);

                    if (!aMatch && !bMatch)
                        continue;

                    FabricIdToCableIds[fabricId].Add(cableId);
                    CableIdToFabricSpeed[cableId] = totalEstimatedCapacity;
                    CableIdToFabricId[cableId] = fabricId;
                }

                if (FabricIdToCableIds[fabricId].Count > 0)
                {
                    string cableList = string.Join(",", FabricIdToCableIds[fabricId].OrderBy(x => x));
                    string deviceList = string.Join(", ", allDeviceNames.OrderBy(x => x));
                    LogToFile($"FABRIC CABLE MAP | fabricId={fabricId} | speed={totalEstimatedCapacity.ToString("0.##", CultureInfo.InvariantCulture)} | deviceNames=[{deviceList}] | cableIds=[{cableList}]");
                }

                clusterIndex++;
            }
        }

        private static int GetTrackedCableCountForFabric(string fabricId)
        {
            if (FabricIdToCableIds.TryGetValue(fabricId, out HashSet<int> ids))
                return ids.Count;

            return 0;
        }

        private static string BuildFabricSignature(List<SwitchNode> fabric, float totalEstimatedCapacity, int totalPorts)
        {
            string members = string.Join("|", fabric
                .OrderBy(x => x.GameObject != null ? x.GameObject.GetInstanceID() : int.MinValue)
                .Select(x => x.GameObject != null ? x.GameObject.GetInstanceID().ToString(CultureInfo.InvariantCulture) : "null"));

            return $"{members}::{totalEstimatedCapacity.ToString("0.###", CultureInfo.InvariantCulture)}::{totalPorts.ToString(CultureInfo.InvariantCulture)}";
        }

        private static void ApplyCombinedFabricToSwitch(SwitchNode node, string fabricId, int memberIndex, int memberCount, float totalEstimatedCapacity, int totalPorts)
        {
            if (node == null || node.GameObject == null)
                return;

            FabricGroupTag tag = node.GameObject.GetComponent<FabricGroupTag>();
            if (tag == null)
                tag = node.GameObject.AddComponent<FabricGroupTag>();

            tag.FabricId = fabricId;
            tag.MemberIndex = memberIndex;
            tag.MemberCount = memberCount;
            tag.AggregatedBandwidth = totalEstimatedCapacity;
            tag.AggregatedPortCount = totalPorts;

            ApplyLikelyMembers(node.GameObject, fabricId, memberIndex, memberCount, totalEstimatedCapacity, totalPorts);

            foreach (Component c in node.GameObject.GetComponentsInChildren<Component>(true))
            {
                if (c == null)
                    continue;

                ApplyLikelyMembers(c, fabricId, memberIndex, memberCount, totalEstimatedCapacity, totalPorts);
            }
        }

        private static void ApplyLikelyMembers(object target, string fabricId, int memberIndex, int memberCount, float totalEstimatedCapacity, int totalPorts)
        {
            TrySetStringMember(target, "fabricGroupId", fabricId);
            TrySetStringMember(target, "switchGroupId", fabricId);
            TrySetStringMember(target, "lacpGroupId", fabricId);
            TrySetStringMember(target, "groupId", fabricId);

            TrySetBoolMember(target, "fabricEnabled", true);
            TrySetBoolMember(target, "autoLacp", true);
            TrySetBoolMember(target, "lacpEnabled", true);
            TrySetBoolMember(target, "switchCombined", true);
            TrySetBoolMember(target, "isCombinedSwitchFabric", true);

            TrySetFloatMember(target, "aggregatedBandwidth", totalEstimatedCapacity);
            TrySetFloatMember(target, "fabricBandwidth", totalEstimatedCapacity);
            TrySetFloatMember(target, "sharedBandwidth", totalEstimatedCapacity);
            TrySetFloatMember(target, "maxBandwidth", totalEstimatedCapacity);
            TrySetFloatMember(target, "throughput", totalEstimatedCapacity);
            TrySetFloatMember(target, "maxThroughput", totalEstimatedCapacity);
            TrySetFloatMember(target, "speed", totalEstimatedCapacity);

            TrySetIntMember(target, "aggregatedPortCount", totalPorts);
            TrySetIntMember(target, "fabricPortCount", totalPorts);
            TrySetIntMember(target, "memberIndex", memberIndex);
            TrySetIntMember(target, "memberCount", memberCount);
        }

        private List<SwitchNode> DiscoverPlacedLooseSwitchCandidates()
        {
            var results = new List<SwitchNode>();
            Transform[] allTransforms = UnityEngine.Object.FindObjectsOfType<Transform>();
            HashSet<int> seen = new();

            foreach (Transform t in allTransforms)
            {
                if (t == null || t.gameObject == null)
                    continue;

                GameObject go = t.gameObject;
                int id = go.GetInstanceID();

                if (seen.Contains(id))
                    continue;

                if (!LooksLikeAllowedSwitchRoot(go, out int qsfpCount, out int sfpCount))
                    continue;

                string parentChain = BuildParentChain(go.transform, 8);
                string parentName = go.transform.parent != null ? go.transform.parent.name : "<root>";

                if (parentChain.Contains("CustomerBase", StringComparison.OrdinalIgnoreCase) ||
                    parentChain.Contains("CustomerBases", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!parentChain.Contains("UsableObjects", StringComparison.OrdinalIgnoreCase))
                    continue;

                var node = new SwitchNode
                {
                    GameObject = go,
                    ParentName = parentName,
                    ParentChain = parentChain,
                    WorldPosition = go.transform.position,
                    LocalPosition = go.transform.localPosition,
                    ApproxHeight = EstimateObjectHeight(go),
                    QsfpPortCount = qsfpCount,
                    SfpPortCount = sfpCount
                };

                node.EstimatedCapacityGbps = EstimateCapacityGbps(node);

                results.Add(node);
                seen.Add(id);
            }

            return results;
        }

        private List<SwitchNode> FilterJunkPileCandidates(List<SwitchNode> candidates)
        {
            var filtered = new List<SwitchNode>();
            if (candidates == null || candidates.Count == 0)
                return filtered;

            Dictionary<string, int> bucketCounts = BuildPositionBuckets(candidates);

            foreach (SwitchNode node in candidates)
            {
                if (IsLikelyOriginJunk(node))
                    continue;

                string bucketKey = GetPositionBucketKey(node.WorldPosition);
                if (bucketCounts.TryGetValue(bucketKey, out int count) && count > MaxObjectsPerPositionBucket)
                    continue;

                filtered.Add(node);
            }

            return filtered;
        }

        private static Dictionary<string, int> BuildPositionBuckets(List<SwitchNode> candidates)
        {
            var buckets = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (SwitchNode node in candidates)
            {
                string key = GetPositionBucketKey(node.WorldPosition);
                if (!buckets.ContainsKey(key))
                    buckets[key] = 0;

                buckets[key]++;
            }

            return buckets;
        }

        private static string GetPositionBucketKey(Vector3 pos)
        {
            int bx = Mathf.RoundToInt(pos.x / PositionBucketSize);
            int by = Mathf.RoundToInt(pos.y / PositionBucketSize);
            int bz = Mathf.RoundToInt(pos.z / PositionBucketSize);
            return $"{bx}:{by}:{bz}";
        }

        private static bool IsLikelyOriginJunk(SwitchNode node)
        {
            float xzRadius = Mathf.Sqrt((node.WorldPosition.x * node.WorldPosition.x) + (node.WorldPosition.z * node.WorldPosition.z));
            return xzRadius <= OriginXZRejectRadius;
        }

        private static bool LooksLikeAllowedSwitchRoot(GameObject go, out int qsfpCount, out int sfpCount)
        {
            qsfpCount = 0;
            sfpCount = 0;

            if (go == null)
                return false;

            string baseName = NormalizeCloneName(go.name);
            if (!AllowedSwitchNames.Contains(baseName))
                return false;

            qsfpCount = CountPorts(go, "QSFP_port.");
            sfpCount = CountPorts(go, "SFP_port.");

            int totalPorts = qsfpCount + sfpCount;
            if (totalPorts < 4)
                return false;

            Transform parent = go.transform.parent;
            if (parent != null && parent.gameObject != null)
            {
                string parentBaseName = NormalizeCloneName(parent.gameObject.name);
                if (AllowedSwitchNames.Contains(parentBaseName))
                    return false;
            }

            return true;
        }

        private static List<List<SwitchNode>> BuildWorldSpaceFabricGroups(List<SwitchNode> switches)
        {
            var groups = new List<List<SwitchNode>>();
            var remaining = new HashSet<SwitchNode>(switches);

            while (remaining.Count > 0)
            {
                SwitchNode seed = remaining.First();
                remaining.Remove(seed);

                var group = new List<SwitchNode>();
                var queue = new Queue<SwitchNode>();
                queue.Enqueue(seed);

                while (queue.Count > 0)
                {
                    SwitchNode current = queue.Dequeue();
                    group.Add(current);

                    List<SwitchNode> neighbors = remaining
                        .Where(other => AreAdjacent(current, other, out _))
                        .ToList();

                    foreach (SwitchNode neighbor in neighbors)
                    {
                        remaining.Remove(neighbor);
                        queue.Enqueue(neighbor);
                    }
                }

                groups.Add(group);
            }

            return groups;
        }

        private static bool AreAdjacent(SwitchNode a, SwitchNode b, out string reason)
        {
            float dx = Mathf.Abs(a.WorldPosition.x - b.WorldPosition.x);
            float dz = Mathf.Abs(a.WorldPosition.z - b.WorldPosition.z);

            if (dx > SameFabricXZTolerance || dz > SameFabricXZTolerance)
            {
                reason = $"XZ too far | dx={dx:0.###} dz={dz:0.###}";
                return false;
            }

            float expectedStep = Mathf.Max(
                a.ApproxHeight > 0.001f ? a.ApproxHeight : FallbackExpectedStep,
                b.ApproxHeight > 0.001f ? b.ApproxHeight : FallbackExpectedStep
            );

            float actualStep = Mathf.Abs(a.WorldPosition.y - b.WorldPosition.y);

            if (Mathf.Abs(actualStep - expectedStep) > AdjacentYTolerance)
            {
                reason = $"Y gap mismatch | actual={actualStep:0.###} expected={expectedStep:0.###}";
                return false;
            }

            reason = "OK";
            return true;
        }

        private static float EstimateCapacityGbps(SwitchNode node)
        {
            float qsfpContribution = node.QsfpPortCount * 40f;
            float sfpContribution = node.SfpPortCount * 10f;
            return qsfpContribution + sfpContribution;
        }

        private static HashSet<string> DiscoverPossibleDeviceNames(GameObject go)
        {
            HashSet<string> result = new(StringComparer.OrdinalIgnoreCase);
            if (go == null)
                return result;

            foreach (Component c in go.GetComponentsInChildren<Component>(true))
            {
                if (c == null)
                    continue;

                Type t = c.GetType();

                TryAddStringMember(t, c, "deviceName", result);
                TryAddStringMember(t, c, "DeviceName", result);
                TryAddStringMember(t, c, "switchId", result);
                TryAddStringMember(t, c, "SwitchId", result);
                TryAddStringMember(t, c, "networkId", result);
                TryAddStringMember(t, c, "NetworkId", result);
                TryAddStringMember(t, c, "id", result);
                TryAddStringMember(t, c, "Id", result);

                foreach (FieldInfo f in SafeFields(t))
                {
                    if (f.FieldType != typeof(string))
                        continue;

                    string lower = f.Name.ToLowerInvariant();
                    if (!(lower.Contains("device") || lower.Contains("switch") || lower.Contains("network") || lower == "id"))
                        continue;

                    TryAddFieldString(f, c, result);
                }

                foreach (PropertyInfo p in SafeProperties(t))
                {
                    if (p.PropertyType != typeof(string) || !p.CanRead)
                        continue;

                    string lower = p.Name.ToLowerInvariant();
                    if (!(lower.Contains("device") || lower.Contains("switch") || lower.Contains("network") || lower == "id"))
                        continue;

                    TryAddPropertyString(p, c, result);
                }
            }

            return result;
        }

        private static IEnumerable<FieldInfo> SafeFields(Type t)
        {
            try { return t.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); }
            catch { return Array.Empty<FieldInfo>(); }
        }

        private static IEnumerable<PropertyInfo> SafeProperties(Type t)
        {
            try { return t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); }
            catch { return Array.Empty<PropertyInfo>(); }
        }

        private static void TryAddStringMember(Type t, object obj, string name, HashSet<string> result)
        {
            try
            {
                FieldInfo f = t.GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (f != null && f.FieldType == typeof(string))
                    AddIfPlausibleDeviceName(f.GetValue(obj) as string, result);

                PropertyInfo p = t.GetProperty(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (p != null && p.CanRead && p.PropertyType == typeof(string))
                    AddIfPlausibleDeviceName(p.GetValue(obj) as string, result);
            }
            catch
            {
            }
        }

        private static void TryAddFieldString(FieldInfo f, object obj, HashSet<string> result)
        {
            try
            {
                AddIfPlausibleDeviceName(f.GetValue(obj) as string, result);
            }
            catch
            {
            }
        }

        private static void TryAddPropertyString(PropertyInfo p, object obj, HashSet<string> result)
        {
            try
            {
                AddIfPlausibleDeviceName(p.GetValue(obj) as string, result);
            }
            catch
            {
            }
        }

        private static void AddIfPlausibleDeviceName(string value, HashSet<string> result)
        {
            if (string.IsNullOrWhiteSpace(value))
                return;

            string trimmed = value.Trim();
            if (trimmed.Length < 2)
                return;

            result.Add(trimmed);
        }

        private void DumpNearMisses()
        {
            try
            {
                Transform[] allTransforms = UnityEngine.Object.FindObjectsOfType<Transform>();

                LogToFile("NEAR MISS DUMP START");

                int written = 0;
                foreach (Transform t in allTransforms)
                {
                    if (t == null || t.gameObject == null)
                        continue;

                    string name = t.gameObject.name ?? string.Empty;
                    string lower = name.ToLowerInvariant();

                    if (!(lower.Contains("switch") || lower.Contains("qsfp") || lower.Contains("sfp")))
                        continue;

                    int qsfpCount = CountPorts(t.gameObject, "QSFP_port.");
                    int sfpCount = CountPorts(t.gameObject, "SFP_port.");

                    LogToFile(
                        $"NEAR MISS | name={t.gameObject.name} | normalized={NormalizeCloneName(t.gameObject.name)} | parent={t.parent?.name ?? "<root>"} | chain={BuildParentChain(t, 6)} | qsfp={qsfpCount} | sfp={sfpCount}"
                    );

                    written++;
                    if (written >= 60)
                        break;
                }

                LogToFile("NEAR MISS DUMP END");
            }
            catch (Exception ex)
            {
                LogToFile($"DumpNearMisses failed: {ex}");
            }
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
                    return prop.GetValue(target);

                FieldInfo field = t.GetField(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (field != null)
                    return field.GetValue(target);
            }
            catch
            {
            }

            return null;
        }

        private static void TrySetStringMember(object target, string memberName, string value) => TrySetMember(target, memberName, value);
        private static void TrySetBoolMember(object target, string memberName, bool value) => TrySetMember(target, memberName, value);
        private static void TrySetFloatMember(object target, string memberName, float value) => TrySetMember(target, memberName, value);
        private static void TrySetIntMember(object target, string memberName, int value) => TrySetMember(target, memberName, value);

        private static void TrySetMember(object target, string memberName, object value)
        {
            if (target == null || string.IsNullOrWhiteSpace(memberName))
                return;

            try
            {
                Type t = target.GetType();

                FieldInfo field = t.GetField(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (field != null && field.FieldType.IsAssignableFrom(value.GetType()))
                {
                    field.SetValue(target, value);
                    return;
                }

                PropertyInfo prop = t.GetProperty(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (prop != null && prop.CanWrite && prop.PropertyType.IsAssignableFrom(value.GetType()))
                {
                    prop.SetValue(target, value);
                }
            }
            catch
            {
            }
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

        private static int CountPorts(GameObject go, string prefix)
        {
            int count = 0;

            foreach (Transform t in go.GetComponentsInChildren<Transform>(true))
            {
                if (t == null)
                    continue;

                string n = t.name ?? string.Empty;
                if (n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    count++;
            }

            return count;
        }

        private static float EstimateObjectHeight(GameObject go)
        {
            if (go == null)
                return FallbackExpectedStep;

            Renderer[] renderers = go.GetComponentsInChildren<Renderer>(true);
            if (renderers == null || renderers.Length == 0)
                return FallbackExpectedStep;

            Bounds bounds = renderers[0].bounds;
            for (int i = 1; i < renderers.Length; i++)
                bounds.Encapsulate(renderers[i].bounds);

            return Mathf.Max(0.001f, bounds.size.y);
        }

        private static string BuildParentChain(Transform t, int depth)
        {
            if (t == null)
                return "<null>";

            List<string> parts = new();
            Transform current = t.parent;
            int count = 0;

            while (current != null && count < depth)
            {
                parts.Add(current.name);
                current = current.parent;
                count++;
            }

            if (parts.Count == 0)
                return "<root>";

            return string.Join(" <- ", parts);
        }

        private static void LogToFile(string message)
        {
            try
            {
                Directory.CreateDirectory(DebugFolderPath);
                File.AppendAllText(DebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] {message}{Environment.NewLine}");
            }
            catch
            {
            }
        }

        // Harmony patch handlers

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
                var info = new CableRegistrationInfo
                {
                    CableId = __0,
                    DeviceA = __5 ?? string.Empty,
                    DeviceB = __6 ?? string.Empty,
                    ExtraA = __9 ?? string.Empty,
                    ExtraB = __10 ?? string.Empty,
                    PortA = __7,
                    PortB = __8
                };

                RegisteredCables[__0] = info;

                RegisterDeviceName(__5, __0);
                RegisterDeviceName(__6, __0);
                RegisterDeviceName(__9, __0);
                RegisterDeviceName(__10, __0);

                LogToFile($"REGISTER CABLE | cableId={__0} | deviceA={info.DeviceA} | deviceB={info.DeviceB} | extraA={info.ExtraA} | extraB={info.ExtraB} | portA={__7} | portB={__8}");
            }
            catch (Exception ex)
            {
                LogToFile($"NetworkMap_RegisterCableConnection_Postfix failed: {ex}");
            }
        }

        private static void RegisterDeviceName(string name, int cableId)
        {
            if (string.IsNullOrWhiteSpace(name))
                return;

            string trimmed = name.Trim();
            if (!DeviceNameToCableIds.TryGetValue(trimmed, out HashSet<int> ids))
            {
                ids = new HashSet<int>();
                DeviceNameToCableIds[trimmed] = ids;
            }

            ids.Add(cableId);
        }

        private static void NetworkMap_CreateLACPGroup_Postfix(string __0, string __1, object __2, ref int __result)
        {
            try
            {
                List<int> cableIds = new();

                if (__2 is IEnumerable enumerable && __2 is not string)
                {
                    foreach (object item in enumerable)
                    {
                        if (item == null)
                            continue;

                        if (int.TryParse(item.ToString(), out int id))
                            cableIds.Add(id);
                    }
                }

                string joined = string.Join(",", cableIds.OrderBy(x => x));
                LogToFile($"CREATE LACP | groupId={__result} | deviceA={__0} | deviceB={__1} | cableIds=[{joined}]");

                float overrideSpeed = 0f;
                foreach (int cableId in cableIds)
                {
                    if (CableIdToFabricSpeed.TryGetValue(cableId, out float speed))
                        overrideSpeed = Mathf.Max(overrideSpeed, speed);
                }

                if (overrideSpeed > 0f)
                    LogToFile($"CREATE LACP SPEED HINT | groupId={__result} | overrideSpeed={overrideSpeed.ToString("0.##", CultureInfo.InvariantCulture)}");
            }
            catch (Exception ex)
            {
                LogToFile($"NetworkMap_CreateLACPGroup_Postfix failed: {ex}");
            }
        }

        private static void LACPGroup_GetAggregatedSpeed_Postfix(object __instance, ref float __result)
        {
            try
            {
                if (__instance == null)
                    return;

                object value = TryGetMemberValue(__instance, "cableIds");
                if (value == null)
                    return;

                float overrideSpeed = 0f;

                if (value is IEnumerable enumerable && value is not string)
                {
                    foreach (object item in enumerable)
                    {
                        if (item == null)
                            continue;

                        if (!int.TryParse(item.ToString(), out int cableId))
                            continue;

                        if (CableIdToFabricSpeed.TryGetValue(cableId, out float speed))
                            overrideSpeed = Mathf.Max(overrideSpeed, speed);
                    }
                }

                if (overrideSpeed > __result)
                    __result = overrideSpeed;
            }
            catch
            {
            }
        }

        private static void CableLink_SetConnectionSpeed_Prefix(object __instance, ref float __0)
        {
            try
            {
                if (__instance == null)
                    return;

                object value = TryGetMemberValue(__instance, "cableIDsOnLink");
                if (value == null)
                    return;

                float overrideSpeed = 0f;

                if (value is IEnumerable enumerable && value is not string)
                {
                    foreach (object item in enumerable)
                    {
                        if (item == null)
                            continue;

                        if (!int.TryParse(item.ToString(), out int cableId))
                            continue;

                        if (CableIdToFabricSpeed.TryGetValue(cableId, out float speed))
                            overrideSpeed = Mathf.Max(overrideSpeed, speed);
                    }
                }
                else if (int.TryParse(value.ToString(), out int oneId))
                {
                    if (CableIdToFabricSpeed.TryGetValue(oneId, out float speed))
                        overrideSpeed = Mathf.Max(overrideSpeed, speed);
                }

                if (overrideSpeed > __0)
                    __0 = overrideSpeed;
            }
            catch
            {
            }
        }

        private static void NetworkMap_UpdateCustomerServerCountAndSpeed_Prefix(ref float __2)
        {
            try
            {
                if (FabricIdToSpeed.Count == 0)
                    return;

                float maxFabricSpeed = FabricIdToSpeed.Values.Max();
                if (maxFabricSpeed > __2)
                    __2 = maxFabricSpeed;
            }
            catch
            {
            }
        }
    }

    public sealed class FabricGroupTag : MonoBehaviour
    {
        public FabricGroupTag(IntPtr ptr) : base(ptr) { }

        public string FabricId { get; set; }
        public int MemberIndex { get; set; }
        public int MemberCount { get; set; }
        public float AggregatedBandwidth { get; set; }
        public int AggregatedPortCount { get; set; }
    }

    internal sealed class SwitchNode
    {
        public GameObject GameObject;
        public string ParentName;
        public string ParentChain;
        public Vector3 WorldPosition;
        public Vector3 LocalPosition;
        public float ApproxHeight;
        public int QsfpPortCount;
        public int SfpPortCount;
        public float EstimatedCapacityGbps;

        public int TotalPorts => QsfpPortCount + SfpPortCount;
    }

    internal sealed class CableRegistrationInfo
    {
        public int CableId;
        public string DeviceA;
        public string DeviceB;
        public string ExtraA;
        public string ExtraB;
        public int PortA;
        public int PortB;
    }
}