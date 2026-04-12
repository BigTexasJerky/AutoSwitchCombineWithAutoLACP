using HarmonyLib;
using Il2CppInterop.Runtime.Injection;
using MelonLoader;
using MelonLoader.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using UnityEngine;
using Component = UnityEngine.Component;

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "0.1.7", "Big Texas Jerky")]
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

        private static readonly string[] InterestingNeedles =
        {
            "lacp",
            "group",
            "bond",
            "portchannel",
            "channel",
            "bandwidth",
            "throughput",
            "speed",
            "transfer",
            "traffic",
            "packet",
            "link",
            "network",
            "cable",
            "devicea",
            "deviceb",
            "groupid"
        };

        private static readonly string[] RecalcMethodNeedles =
        {
            "refresh",
            "recalc",
            "recalculate",
            "rebuild",
            "recompute",
            "update",
            "apply",
            "sync",
            "reload",
            "bandwidth",
            "throughput",
            "network",
            "traffic",
            "speed",
            "link"
        };

        private static string DebugFolderPath =>
            Path.Combine(MelonEnvironment.ModsDirectory, "AutoSwitch");

        private static string DebugLogPath =>
            Path.Combine(DebugFolderPath, "autoswitch-debug.log");

        private float _nextScanTime;
        private string _lastSummary = string.Empty;
        private bool _loggedNearMissesThisScene;
        private bool _loggedNativeTypeScanThisScene;

        private static readonly Dictionary<string, string> LastClusterSignatures = new(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> LoggedComponentTypes = new(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> LoggedTypeMembers = new(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> LoggedNativeTypes = new(StringComparer.OrdinalIgnoreCase);

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();

            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(DebugLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] AutoSwitch debug log started.{Environment.NewLine}");

            MelonLogger.Msg("[AutoSwitch] v0.1.7 active. Cluster mode with throughput patch scaffolding.");
        }

        public override void OnSceneWasLoaded(int buildIndex, string sceneName)
        {
            if (buildIndex == 0)
                return;

            _nextScanTime = 0f;
            _lastSummary = string.Empty;
            _loggedNearMissesThisScene = false;
            _loggedNativeTypeScanThisScene = false;

            LastClusterSignatures.Clear();
            LoggedComponentTypes.Clear();
            LoggedTypeMembers.Clear();
            LoggedNativeTypes.Clear();

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
            IntrospectActiveFabrics(activeFabrics);

            if (!_loggedNativeTypeScanThisScene)
            {
                _loggedNativeTypeScanThisScene = true;
                ScanAssemblyForInterestingTypes();
            }

            string summary =
                $"SCAN SUMMARY | rawCandidates={rawCandidates.Count} | filteredSwitches={filtered.Count} | activeFabrics={activeFabrics.Count} | adjacentPairs={adjacencyPairs}";

            if (!string.Equals(summary, _lastSummary, StringComparison.Ordinal))
            {
                _lastSummary = summary;
                MelonLogger.Msg($"[AutoSwitch] {summary}");
                LogToFile(summary);

                int fabricIndex = 1;
                foreach (List<SwitchNode> fabric in activeFabrics)
                {
                    float totalEstimatedCapacity = fabric.Sum(f => f.EstimatedCapacityGbps);

                    string fabricLine =
                        $"FABRIC | id=FABRIC-{fabricIndex:000} | members={fabric.Count} | estCapacityGbps={totalEstimatedCapacity.ToString("0.##", CultureInfo.InvariantCulture)} | switches={string.Join(", ", fabric.Select(s => $"{s.GameObject.name}@({s.WorldPosition.x:0.###},{s.WorldPosition.y:0.###},{s.WorldPosition.z:0.###})[ports={s.TotalPorts}]"))}";
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

                foreach (SwitchNode node in ordered)
                {
                    TryInvokeRecalcMethods(node.GameObject);
                }

                LogToFile($"APPLY | fabricId={fabricId} | members={ordered.Count} | totalPorts={totalPorts} | estCapacityGbps={totalEstimatedCapacity.ToString("0.##", CultureInfo.InvariantCulture)}");
                clusterIndex++;
            }
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

        private void IntrospectActiveFabrics(List<List<SwitchNode>> activeFabrics)
        {
            foreach (List<SwitchNode> fabric in activeFabrics)
            {
                foreach (SwitchNode node in fabric)
                {
                    if (node?.GameObject == null)
                        continue;

                    Component[] components = node.GameObject.GetComponentsInChildren<Component>(true);
                    foreach (Component component in components)
                    {
                        if (component == null)
                            continue;

                        Type t = component.GetType();
                        string typeName = t.FullName ?? t.Name;

                        if (!LoggedComponentTypes.Contains(typeName))
                        {
                            LoggedComponentTypes.Add(typeName);
                            LogToFile($"COMPONENT TYPE | {typeName}");
                        }

                        if (HasInterestingName(typeName))
                        {
                            LogInterestingMembers(t);
                        }
                        else
                        {
                            if (TypeContainsInterestingMembers(t))
                                LogInterestingMembers(t);
                        }
                    }
                }
            }
        }

        private void ScanAssemblyForInterestingTypes()
        {
            try
            {
                Assembly assembly = typeof(AutoSwitchMod).Assembly;
                LogToFile($"LOCAL ASSEMBLY | {assembly.GetName().Name}");

                foreach (Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    string asmName = asm.GetName().Name ?? string.Empty;

                    if (!asmName.Equals("Assembly-CSharp", StringComparison.OrdinalIgnoreCase) &&
                        !asmName.Contains("Assembly-CSharp", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    LogToFile($"INTERESTING TYPE SCAN | assembly={asmName}");

                    foreach (Type t in SafeGetTypes(asm))
                    {
                        if (t == null)
                            continue;

                        string fullName = t.FullName ?? t.Name;
                        if (!HasInterestingName(fullName))
                            continue;

                        if (LoggedNativeTypes.Add(fullName))
                        {
                            LogToFile($"NATIVE TYPE | {fullName}");
                            LogInterestingMembers(t);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile($"ScanAssemblyForInterestingTypes failed: {ex}");
            }
        }

        private static IEnumerable<Type> SafeGetTypes(Assembly asm)
        {
            try
            {
                return asm.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                return ex.Types.Where(t => t != null);
            }
            catch
            {
                return Array.Empty<Type>();
            }
        }

        private static bool HasInterestingName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            string lower = name.ToLowerInvariant();
            return InterestingNeedles.Any(n => lower.Contains(n));
        }

        private static bool TypeContainsInterestingMembers(Type t)
        {
            try
            {
                foreach (FieldInfo field in t.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
                {
                    if (HasInterestingName(field.Name))
                        return true;
                }

                foreach (PropertyInfo prop in t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
                {
                    if (HasInterestingName(prop.Name))
                        return true;
                }

                foreach (MethodInfo method in t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
                {
                    if (HasInterestingName(method.Name))
                        return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static void LogInterestingMembers(Type t)
        {
            string typeName = t.FullName ?? t.Name;
            if (!LoggedTypeMembers.Add(typeName))
                return;

            try
            {
                foreach (FieldInfo field in t.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
                {
                    if (!HasInterestingName(field.Name))
                        continue;

                    LogToFileStatic($"FIELD | {typeName} | {field.FieldType.Name} {field.Name}");
                }

                foreach (PropertyInfo prop in t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
                {
                    if (!HasInterestingName(prop.Name))
                        continue;

                    LogToFileStatic($"PROPERTY | {typeName} | {prop.PropertyType.Name} {prop.Name}");
                }

                foreach (MethodInfo method in t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
                {
                    if (!HasInterestingName(method.Name))
                        continue;

                    LogToFileStatic($"METHOD | {typeName} | {method.Name}");
                }
            }
            catch (Exception ex)
            {
                LogToFileStatic($"LogInterestingMembers failed for {typeName}: {ex.Message}");
            }
        }

        private static void TryInvokeRecalcMethods(GameObject go)
        {
            if (go == null)
                return;

            foreach (Component component in go.GetComponentsInChildren<Component>(true))
            {
                if (component == null)
                    continue;

                Type t = component.GetType();

                MethodInfo[] methods;
                try
                {
                    methods = t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                }
                catch
                {
                    continue;
                }

                foreach (MethodInfo method in methods)
                {
                    if (method == null)
                        continue;

                    if (method.GetParameters().Length != 0)
                        continue;

                    string lower = method.Name.ToLowerInvariant();
                    if (!RecalcMethodNeedles.Any(n => lower.Contains(n)))
                        continue;

                    try
                    {
                        method.Invoke(component, null);
                    }
                    catch
                    {
                    }
                }
            }
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

        private static void TrySetStringMember(object target, string memberName, string value)
        {
            TrySetMember(target, memberName, value);
        }

        private static void TrySetBoolMember(object target, string memberName, bool value)
        {
            TrySetMember(target, memberName, value);
        }

        private static void TrySetFloatMember(object target, string memberName, float value)
        {
            TrySetMember(target, memberName, value);
        }

        private static void TrySetIntMember(object target, string memberName, int value)
        {
            TrySetMember(target, memberName, value);
        }

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

        private static void LogToFileStatic(string message)
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
}